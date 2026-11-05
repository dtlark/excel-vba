"""Runs raw_invoices.py from start to finish on made-up files. No Excel needed.

Run them from the python folder with:  python -m unittest discover -s tests
"""
from __future__ import annotations

import io
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo  # noqa: E402

import raw_invoices  # noqa: E402

VENDOR = ["LINE", "ITEM", "DESCRIPTION", "ORDERED", "SHIPPED", "PRICE/UNIT", "TAX", "FREGHT",
          "AMOUNT", "SHIP TO"]
CONS_HEADERS = ["Key", "Notes", "Checked", "Invoice"] + VENDOR  # 14 columns, A to N
KEY_FORMULA = 'Consumables[[#This Row],[Invoice]]&"-"&Consumables[[#This Row],[LINE]]'


def make_raw(path: Path, tab: str = "123456", headers=VENDOR) -> None:
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = tab
    ws.append(headers)
    ws.append([1, "A-100", "Gloves", 10, 10, 2.5, 1.25, 0, 26.25, 12345])
    ws.append([2, "B-200", "Glasses", "4", "4", "3.10", 0, "1.00", 13.4, 12345])  # numbers as text
    ws.append([None, None, "Subtotal for PO 55", None, None, None, None, None, 39.65, None])
    ws.append([None] * 10)  # an empty row
    ws.append([3, "C-300", "Vest", 1, 1, 9.99, 0.5, 0, 10.5, "0012345"])  # 9.99 + 0.5 = 10.49: ERROR
    ws.append([None, None, "Grand Total", None, None, None, None, None, 50.15, None])
    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=10).number_format = "0000000"  # SHIP TO padded with zeros
        ws.cell(row=row, column=1).fill = PatternFill(fill_type="solid", start_color="0000FF")
    ws.cell(row=3, column=6).number_format = "@"  # a Text-formatted cell
    book.save(path)


def make_master(path: Path, totals_row: bool = False, below: str | None = None) -> None:
    book = openpyxl.Workbook()
    book.active.title = "Summary"
    book.active["A1"] = "Master List"
    ws = book.create_sheet("Consolidated Consumables")
    ws.append(CONS_HEADERS)
    ws.append(["=" + KEY_FORMULA, "", "", 111111, 1, "Z-1", "Old row", 1, 1, 5, 0, 0, 5, 999])
    last = 2
    if totals_row:
        ws.append([None] * 8 + ["Total", "=SUBTOTAL(109,Consumables[AMOUNT])"] + [None] * 4)
        last = 3
    if below:
        ws[below] = "something else"
    table = Table(displayName="Consumables", ref=f"A1:N{last}", totalsRowCount=1 if totals_row else None)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    table._initialise_columns()
    for column, name in zip(table.tableColumns, CONS_HEADERS):
        column.name = name
    table.tableColumns[0].calculatedColumnFormula = TableFormula(attr_text=KEY_FORMULA)
    ws.add_table(table)
    ws.conditional_formatting.add("M2:M2", CellIsRule(operator="greaterThan", formula=["100"],
                                                      fill=PatternFill(fill_type="solid", start_color="FFFF00")))
    book.save(path)


class EndToEndTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp())
        fy = self.root / "FY26"
        self.raw = fy / "Excels" / "Raw"
        self.raw.mkdir(parents=True)
        self.invoices = fy / "Excels"
        self.master = fy / "FY26 Master List.xlsx"
        self.settings = self.root / "settings.ini"
        self.settings.write_text(
            "[Settings]\n"
            "Fiscal year = 26\n"
            f"RAW folder = {self.root}/{{FY}}/Excels/Raw\n"
            "Done folder =\n"
            f"Invoice folder = {self.root}/{{FY}}/Excels\n"
            f"Master List file = {self.root}/{{FY}}/{{FY}} Master List\n", encoding="utf-8")

    def run_script(self, answer: str = "y"):
        output = io.StringIO()
        with mock.patch.object(raw_invoices, "SETTINGS_FILE", self.settings), \
                mock.patch("builtins.input", return_value=answer), redirect_stdout(output):
            code = raw_invoices.main()
        return code, output.getvalue()

    def test_one_file_all_the_way_through(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_master(self.master)
        code, output = self.run_script()
        self.assertEqual(code, 0, output)

        # The Invoice file
        invoice_path = self.invoices / "Invoice 123456 (ABC).xlsx"
        self.assertTrue(invoice_path.exists(), output)
        ws = openpyxl.load_workbook(invoice_path).active
        self.check_invoice_tab(ws)

        # The Master List: the same tab at the end, with an orange tab
        master = openpyxl.load_workbook(self.master)
        self.assertEqual(master.sheetnames[-1], "123456")
        new_tab = master["123456"]
        self.assertTrue(new_tab.sheet_properties.tabColor.rgb.endswith("E97132"))
        self.check_invoice_tab(new_tab)

        # The Consumables table grew by the 3 data rows
        cons = master["Consolidated Consumables"]
        table = cons.tables["Consumables"]
        self.assertEqual(table.ref, "A1:N5")
        self.assertEqual(table.autoFilter.ref, "A1:N5")
        for row, line in zip((3, 4, 5), (1, 2, 3)):
            self.assertEqual(cons.cell(row=row, column=1).value, "=" + KEY_FORMULA)  # formula column filled
            self.assertEqual(cons.cell(row=row, column=4).value, 123456)  # the tab name, as a number
            self.assertEqual(cons.cell(row=row, column=5).value, line)
        self.assertEqual(cons.cell(row=4, column=9).value, 4)  # SHIPPED "4" became a number
        self.assertEqual(cons.cell(row=4, column=10).value, 3.1)  # PRICE/UNIT "3.10" too
        self.assertEqual(cons.cell(row=4, column=10).number_format, raw_invoices.FMT_ACCOUNTING)
        self.assertEqual(cons.cell(row=5, column=14).value, 12345)  # "0012345" lost its zeros
        self.assertEqual(cons.cell(row=5, column=14).number_format, "0")
        cf_ranges = [str(rule_set.sqref) for rule_set in cons.conditional_formatting]
        self.assertIn("M2:M5", cf_ranges)  # conditional formatting followed the table down

        # The RAW file is in Done, and there's a backup of the Master List
        self.assertFalse((self.raw / "123456 ABC.xlsx").exists())
        self.assertTrue((self.raw / "Done" / "123456 ABC.xlsx").exists())
        self.assertEqual(len(list((self.master.parent / "Backups").iterdir())), 1)

        # Running again with the same file stops, lists every reason, and changes nothing
        make_raw(self.raw / "123456 ABC.xlsx")
        before = self.master.read_bytes()
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("STOPPED: found duplicates", output)
        for reason in ("the Done folder already has a file with this name",
                       "the Master List already has a tab named 123456",
                       "the Consumables table already has rows for 123456",
                       "the Invoice folder already has Invoice 123456 (ABC).xlsx"):
            self.assertIn(reason, output)
        self.assertEqual(self.master.read_bytes(), before)
        self.assertTrue((self.raw / "123456 ABC.xlsx").exists())

    def check_invoice_tab(self, ws):
        self.assertEqual(ws.title, "123456")
        self.assertEqual(list(ws.tables), ["Inv_123456"])
        table = ws.tables["Inv_123456"]
        self.assertEqual(table.ref, "A3:M6")  # 3 data rows; subtotal, grand total and empty rows gone
        self.assertEqual(table.tableStyleInfo.name, "TableStyleMedium1")
        self.assertEqual([c.value for c in ws[3]], VENDOR + ["Check Number", " ", "V"])
        self.assertEqual([ws.cell(row=r, column=1).value for r in (4, 5, 6)], [1, 2, 3])
        self.assertIsNone(ws.cell(row=4, column=1).fill.fgColor.rgb if ws.cell(row=4, column=1).fill.fill_type else None)
        self.assertEqual(ws.cell(row=5, column=5).value, 4)  # numbers stored as text became numbers
        self.assertEqual(ws.cell(row=5, column=6).value, 3.1)
        self.assertEqual(ws.cell(row=5, column=8).value, 1.0)
        self.assertEqual(ws.cell(row=4, column=5).number_format, "0.00")  # SHIPPED
        self.assertEqual(ws.cell(row=4, column=4).number_format, "0.00")  # ORDERED
        for column in (6, 7, 8, 9, 11, 13):  # PRICE/UNIT, TAX, FREGHT, AMOUNT, Check Number, V
            self.assertEqual(ws.cell(row=4, column=column).number_format, raw_invoices.FMT_ACCOUNTING)
        self.assertEqual(ws.cell(row=6, column=10).number_format, "0")  # SHIP TO without zeros
        self.assertEqual(ws["K4"].value, "=Inv_123456[[#This Row],[SHIPPED]]*Inv_123456[[#This Row],"
                                         "[PRICE/UNIT]]+Inv_123456[[#This Row],[TAX]]"
                                         "+Inv_123456[[#This Row],[FREGHT]]")
        self.assertEqual(ws["L5"].value, '=IF(ROUND(Inv_123456[[#This Row],[AMOUNT]]-Inv_123456[[#This Row],'
                                         '[Check Number]],6)=0,"","ERROR")')
        self.assertEqual(ws["M6"].value, '=IF(L6="ERROR",Inv_123456[[#This Row],[Check Number]]'
                                         '-Inv_123456[[#This Row],[AMOUNT]],"")')
        self.assertEqual(ws["I2"].value, "=SUBTOTAL(9,Inv_123456[AMOUNT])")
        self.assertEqual(ws["I2"].number_format, raw_invoices.FMT_ACCOUNTING)
        for cell in ("K3", "L3", "M3"):
            self.assertTrue(ws[cell].fill.fgColor.rgb.endswith("E97132"), cell)
        self.assertIn("L4:L6", [str(rule_set.sqref) for rule_set in ws.conditional_formatting])

    def test_totals_row_moves_down(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_master(self.master, totals_row=True)
        code, output = self.run_script()
        self.assertEqual(code, 0, output)
        cons = openpyxl.load_workbook(self.master)["Consolidated Consumables"]
        table = cons.tables["Consumables"]
        self.assertEqual(table.ref, "A1:N6")
        self.assertEqual(table.autoFilter.ref, "A1:N5")
        self.assertEqual(cons["J6"].value, "=SUBTOTAL(109,Consumables[AMOUNT])")
        self.assertEqual(cons["E5"].value, 3)

    def test_something_below_the_table_skips_the_file(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_master(self.master, below="B4")
        before = self.master.read_bytes()
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("there's something in the rows right below the Consumables table (cell B4)", output)
        self.assertEqual(self.master.read_bytes(), before)
        self.assertTrue((self.raw / "123456 ABC.xlsx").exists())
        self.assertFalse((self.invoices / "Invoice 123456 (ABC).xlsx").exists())

    def test_missing_column_skips_just_that_file(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_raw(self.raw / "222222.xlsx", tab="222222",
                 headers=[h if h != "AMOUNT" else "TOTAL" for h in VENDOR])
        make_master(self.master)
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("222222.xlsx: row 1 has no AMOUNT column. Row 1 has: LINE, ITEM", output)
        self.assertTrue((self.invoices / "Invoice 123456 (ABC).xlsx").exists())
        self.assertTrue((self.raw / "222222.xlsx").exists())
        self.assertIn("123456", openpyxl.load_workbook(self.master).sheetnames)

    def test_stops_when_the_master_list_is_open(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_master(self.master)
        (self.master.parent / "~$FY26 Master List.xlsx").write_text("lock")
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("is open in Excel", output)
        self.assertTrue((self.raw / "123456 ABC.xlsx").exists())

    def test_answering_no_changes_nothing(self):
        make_raw(self.raw / "123456 ABC.xlsx")
        make_master(self.master)
        before = self.master.read_bytes()
        code, _ = self.run_script(answer="n")
        self.assertEqual(code, 0)
        self.assertEqual(self.master.read_bytes(), before)
        self.assertTrue((self.raw / "123456 ABC.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
