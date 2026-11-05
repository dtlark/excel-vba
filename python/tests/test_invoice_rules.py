"""Tests for invoice_rules.py. They don't need Excel.

Run them from the python folder with:  python -m unittest discover -s tests
"""
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import invoice_rules as rules  # noqa: E402
import raw_invoices  # noqa: E402  (imports fine without xlwings)


class TextToNumberTests(unittest.TestCase):
    def check(self, text, expected):
        result = rules.text_to_number(text)
        if expected is None:
            self.assertIsNone(result, text)
        else:
            self.assertIsNotNone(result, text)
            self.assertAlmostEqual(result[0], expected[0], places=12, msg=text)
            self.assertEqual(result[1], expected[1], text)

    def test_numbers(self):
        cases = {
            "12": (12, 0), "12.50": (12.5, 2), " 12.50 ": (12.5, 2), "1,234.56": (1234.56, 2),
            "1,234,567": (1234567, 0), "$12.50": (12.5, 2), "$ 1,234.00": (1234, 2),
            "-12.5": (-12.5, 1), "+3": (3, 0), "12.50-": (-12.5, 2), "(12.50)": (-12.5, 2),
            "$ (5.00)": (-5, 2), "$-5": (-5, 0), "-$5": (-5, 0), ".75": (0.75, 2),
            "0.50": (0.5, 2), "0": (0, 0), "5.": (5, 0), "654321": (654321, 0),
        }
        for text, expected in cases.items():
            self.check(text, expected)

    def test_invisible_characters(self):
        for text in ("12.50\t", "12.50\r\n", "\xa012.50\xa0", "\u200b12.50", "\ufeff12.50"):
            self.check(text, (12.5, 2))

    def test_leading_zeros_are_dropped(self):
        self.check("00123", (123, 0))
        self.check("007", (7, 0))
        self.check("00.50", (0.5, 2))

    def test_stays_text(self):
        for text in ("1234567890123456", "12,34", ",123", "1.2.3", "10/15/2026", "12%", "1e5",
                     "N/A", "", "-", ".", "$", "$$5", "(-5)", "1 234", "ABC-123", "\u0663"):
            self.check(text, None)


class NumberFormatTests(unittest.TestCase):
    def test_zero_pad_formats(self):
        for fmt in ("0000000", "00000", "00000-0000", "[Red]000000", "000000;-000000"):
            self.assertTrue(rules.is_zero_pad_format(fmt), fmt)
        for fmt in ("General", "0", "0.00", "#,##0.00", "m/d/yyyy", "@", "0%",
                    raw_invoices.FMT_ACCOUNTING, "[$-409]mmm-yy", None):
            self.assertFalse(rules.is_zero_pad_format(fmt), fmt)

    def test_number_format_keeps_decimals(self):
        cases = {"0000000": "0", "0000.00": "0.00", "[Red]000.0": "0.0", "00000-0000": "0"}
        for fmt, expected in cases.items():
            self.assertEqual(rules.number_format_with_decimals(rules.format_decimals(fmt)), expected)


class FileNameTests(unittest.TestCase):
    def test_invoice_file_names(self):
        cases = {
            "123456 ABC.xlsx": "Invoice 123456 (ABC).xlsx",
            "123456_ABC.xlsx": "Invoice 123456 (ABC).xlsx",
            "123456 - ABC.xlsx": "Invoice 123456 (ABC).xlsx",
            "123456ABC.xlsx": "Invoice 123456 (ABC).xlsx",
            "123456 (ABC).xlsx": "Invoice 123456 (ABC).xlsx",
            "123456.xlsx": "Invoice 123456.xlsx",
        }
        for raw_name, expected in cases.items():
            name, note = rules.invoice_file_name(raw_name, "123456")
            self.assertEqual(name, expected, raw_name)
            self.assertEqual(note, "")

    def test_file_name_not_starting_with_tab_name(self):
        name, note = rules.invoice_file_name("Vendor invoice.xlsx", "123456")
        self.assertEqual(name, "Invoice 123456 (Vendor invoice).xlsx")
        self.assertIn("doesn't start with its tab name", note)

    def test_table_names(self):
        self.assertEqual(rules.table_name_for("123456"), "Inv_123456")
        self.assertEqual(rules.table_name_for("Oct 2026"), "Inv_Oct_2026")

    def test_master_candidates(self):
        base, files = rules.master_candidates(r"C:\Inv\FY26\FY26 Master List")
        self.assertEqual(base, r"C:\Inv\FY26\FY26 Master List")
        self.assertEqual(files[0], r"C:\Inv\FY26\FY26 Master List.xlsx")
        base, files = rules.master_candidates(r"C:\Inv\FY26\FY26 Master List.xlsm")
        self.assertEqual(base, r"C:\Inv\FY26\FY26 Master List")
        self.assertEqual(files[0], r"C:\Inv\FY26\FY26 Master List.xlsm")
        base, _ = rules.master_candidates(r"C:\Inv.Dept\Master List")  # dot in a folder name
        self.assertEqual(base, r"C:\Inv.Dept\Master List")

    def test_list_raw_files(self):
        with tempfile.TemporaryDirectory() as folder:
            for name in ("b 2.xlsx", "A 1.XLSX", "~$b 2.xlsx", "notes.txt", "c.xls"):
                (Path(folder) / name).write_text("x")
            (Path(folder) / "Done").mkdir()
            self.assertEqual(rules.list_raw_files(Path(folder)), ["A 1.XLSX", "b 2.xlsx", "c.xls"])


class RowAndHeaderTests(unittest.TestCase):
    def test_unwanted_rows(self):
        self.assertTrue(rules.is_unwanted_row(["", "Subtotal For PO 55", None]))
        self.assertTrue(rules.is_unwanted_row([None, "GRAND TOTAL", 99.0]))
        self.assertTrue(rules.is_unwanted_row([None, "", "  "]))
        self.assertFalse(rules.is_unwanted_row(["Widget", 3.0, 12.5]))
        self.assertFalse(rules.is_unwanted_row([None, "#N/A", None]))  # an Excel error is a value
        self.assertFalse(rules.is_unwanted_row([datetime(2026, 10, 1), None]))

    def test_find_header(self):
        headers = ["ITEM", "Ordered", " SHIPPED ", "PRICE/UNIT", "TAX", "FREGHT", "AMOUNT\xa0",
                   "SHIP\nTO", 2026.0, None]
        self.assertEqual(rules.find_header(headers, "SHIPPED"), 2)
        self.assertEqual(rules.find_header(headers, "FREIGHT|FREGHT"), 5)
        self.assertEqual(rules.find_header(headers, "AMOUNT"), 6)
        self.assertEqual(rules.find_header(headers, "ORDERED"), 1)
        self.assertEqual(rules.find_header(headers, "SHIP TO"), 7)
        self.assertEqual(rules.find_header(headers, "2026"), 8)
        self.assertIsNone(rules.find_header(headers, "QTY"))
        self.assertEqual(rules.first_spelling("FREIGHT|FREGHT"), "FREIGHT")
        self.assertTrue(rules.header_list(headers).startswith("ITEM, ORDERED, SHIPPED, PRICE/UNIT"))

    def test_column_has_value(self):
        self.assertTrue(rules.column_has_value([None, 654321.0, "x"], "654321"))
        self.assertTrue(rules.column_has_value(["654321 "], "654321"))
        self.assertFalse(rules.column_has_value([6543210.0, "65432"], "654321"))


class FormulaTests(unittest.TestCase):
    # In the file, Excel stores [@[SHIPPED]] as Inv_123456[[#This Row],[SHIPPED]].
    def test_formulas_match_the_macro(self):
        t = "Inv_123456"
        self.assertEqual(
            rules.check_number_formula(t, "SHIPPED", "PRICE/UNIT", "TAX", "FREGHT"),
            "=Inv_123456[[#This Row],[SHIPPED]]*Inv_123456[[#This Row],[PRICE/UNIT]]"
            "+Inv_123456[[#This Row],[TAX]]+Inv_123456[[#This Row],[FREGHT]]")
        self.assertEqual(
            rules.error_formula(t, "AMOUNT", "Check Number"),
            '=IF(ROUND(Inv_123456[[#This Row],[AMOUNT]]-Inv_123456[[#This Row],[Check Number]],6)=0,'
            '"","ERROR")')
        self.assertEqual(
            rules.v_formula("P4", t, "Check Number", "AMOUNT"),
            '=IF(P4="ERROR",Inv_123456[[#This Row],[Check Number]]-Inv_123456[[#This Row],[AMOUNT]],"")')
        self.assertEqual(rules.subtotal_formula(t, "AMOUNT"), "=SUBTOTAL(9,Inv_123456[AMOUNT])")

    def test_escaping(self):
        self.assertEqual(rules.escape_column_name("PO #"), "PO '#")
        self.assertEqual(rules.escape_column_name("Qty [ea]"), "Qty '[ea']")
        self.assertEqual(rules.escape_column_name("Rep's"), "Rep''s")


class TableHeaderTests(unittest.TestCase):
    def test_names_like_excel(self):
        names = rules.table_headers(["ITEM", None, "AMOUNT", "amount", 2026.0], ["Check Number", " ", "V"])
        self.assertEqual(names, ["ITEM", "Column2", "AMOUNT", "amount2", "2026", "Check Number", " ", "V"])

    def test_added_name_that_repeats_a_vendor_name(self):
        self.assertEqual(rules.table_headers(["V"], ["Check Number", " ", "V"])[-1], "V2")

    def test_tab_name_as_number(self):
        self.assertEqual(rules.as_number_if_whole("654321"), 654321)
        self.assertEqual(rules.as_number_if_whole("65-4321"), "65-4321")


class SettingsTests(unittest.TestCase):
    def write(self, text):
        folder = tempfile.mkdtemp()
        path = Path(folder) / "settings.ini"
        path.write_text(text, encoding="utf-8")
        return path

    def test_fiscal_year_fills_in_the_paths(self):
        path = self.write(
            "[Settings]\n"
            "Fiscal year = FY27\n"
            'RAW folder = "C:\\Inv\\{FY}\\Excels\\Raw\\"\n'
            "Done folder =\n"
            "Invoice folder = C:\\Inv\\{fy}\\Excels\n"
            "Master List file = C:\\Inv\\{FY}\\{FY} Master List\n")
        settings = rules.load_settings(path)
        self.assertEqual(str(settings.raw_folder), str(Path("C:\\Inv\\FY27\\Excels\\Raw")))
        self.assertEqual(settings.done_folder, Path("C:\\Inv\\FY27\\Excels\\Raw") / "Done")
        self.assertEqual(str(settings.invoice_folder), str(Path("C:\\Inv\\FY27\\Excels")))
        self.assertEqual(settings.master_path, "C:\\Inv\\FY27\\FY27 Master List")

    def test_missing_settings_are_listed(self):
        path = self.write("[Settings]\nRAW folder = C:\\Inv\\{FY}\\Raw\n")
        with self.assertRaises(rules.SettingsError) as caught:
            rules.load_settings(path)
        message = str(caught.exception)
        for label in ("Fiscal year", "Invoice folder", "Master List file"):
            self.assertIn(label, message)

    def test_example_file_loads(self):
        example = Path(__file__).resolve().parents[1] / "settings.example.ini"
        settings = rules.load_settings(example)
        self.assertEqual(settings.master_path, "C:\\Invoices\\FY26\\FY26 Master List")

    def test_missing_file(self):
        with self.assertRaises(rules.SettingsError):
            rules.load_settings(Path(tempfile.mkdtemp()) / "settings.ini")


class MoveToDoneTests(unittest.TestCase):
    def test_moves_and_never_overwrites(self):
        with tempfile.TemporaryDirectory() as folder:
            raw, done = Path(folder) / "raw", Path(folder) / "raw" / "Done"
            done.mkdir(parents=True)
            (raw / "123456 ABC.xlsx").write_text("new")
            (done / "123456 ABC.xlsx").write_text("old")
            self.assertEqual(raw_invoices.move_to_done(raw / "123456 ABC.xlsx", done), "")
            self.assertFalse((raw / "123456 ABC.xlsx").exists())
            self.assertEqual((done / "123456 ABC.xlsx").read_text(), "old")
            self.assertEqual(len(list(done.glob("123456 ABC *.xlsx"))), 1)


if __name__ == "__main__":
    unittest.main()
