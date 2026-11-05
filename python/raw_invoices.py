"""Process the vendor's RAW invoice files: the Python version of vba/RawInvoices.bas.

It does the same job as the macro without Excel. It reads and writes the .xlsx
files directly (with openpyxl), so Excel doesn't need to be running.

For every Excel file in the RAW folder (in file name order) it:
  1. Opens it. Its tab name (e.g. 123456) is the <Tab Name>.
  2. Takes the header row and data rows (values and number formats only),
     leaving out rows that contain "subtotal for" or "grand total", and empty
     rows. Numbers stored as text become real numbers, and leading-zero number
     formats become Number.
  3. Builds a tab with the header on row 3 and a table named Inv_<Tab Name>
     (White, Table Style Medium 1), plus three columns with orange headers:
       Check Number  = SHIPPED x PRICE/UNIT + TAX + FREIGHT
       " " (blank)   = ERROR when AMOUNT isn't equal to Check Number
       V             = Check Number - AMOUNT on ERROR rows
  4. Puts =SUBTOTAL(9,Inv_<Tab Name>[AMOUNT]) on row 2 above AMOUNT.
  5. Saves it as "Invoice <Tab Name> (<rest of the RAW file name>).xlsx".
  6. Adds the same tab to the end of the Master List, with an orange tab.
  7. Adds the data rows (the vendor's columns only) to the bottom of the Master
     List's Consumables table: the tab name in its 4th column, the data from its 5th.
  8. Once the Master List is saved, moves the RAW file into the Done folder.

Any duplicate stops the whole run before anything changes.

Differences from the macro:
  - The Master List must be CLOSED in Excel while this runs.
  - It SAVES the Master List at the end, after copying it into a "Backups"
    folder next to it.
  - openpyxl can't keep everything Excel can. If the Master List has Power
    Query queries, slicers, buttons or shapes, they may be lost when it's
    saved. That's what the backup is for.
  - Formulas show their results when the files are opened in Excel.

Run it:  python raw_invoices.py      (or double-click "Run RAW invoices.bat")
The settings are in settings.ini next to this file.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from collections import OrderedDict
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import invoice_rules as rules

try:
    import openpyxl
    from openpyxl.formatting.formatting import ConditionalFormatting
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.formula.translate import Translator
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter, range_boundaries
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:  # checked in main(), so the rest of this file still imports
    openpyxl = None

HERE = Path(__file__).resolve().parent
SETTINGS_FILE = HERE / "settings.ini"

# Column headers looked for in row 1 of the RAW file (capitals don't matter).
# Put other spellings after a |, like the vendor's "FREGHT".
H_SHIPPED = "SHIPPED"
H_PRICE = "PRICE/UNIT"
H_TAX = "TAX"
H_FREIGHT = "FREIGHT|FREGHT"
H_AMOUNT = "AMOUNT"
H_ORDERED = "ORDERED"  # only formatted, not used in a formula, so a file without it isn't skipped
ADDED_HEADERS = ["Check Number", " ", "V"]  # the blank column's name is a space

TABLE_STYLE = "TableStyleMedium1"  # "White, Table Style Medium 1" in Excel's Table Design gallery
FMT_NUMBER = "0.00"  # Excel's "Number" with 2 decimals
FMT_ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'  # "Accounting" with $

ORANGE = "E97132"
ERROR_TEXT = "9C0006"
ERROR_FILL = "FFC7CE"

# The Master List table that collects every invoice's data rows, and where they go in it
CONSUMABLES_TABLE = "Consumables"
CONSUMABLES_TAB = "Consolidated Consumables"  # only used in messages
CONS_TAB_NAME_COLUMN = 4  # the tab name goes in the table's 4th column
CONS_FIRST_DATA_COLUMN = 5  # the vendor's columns start in its 5th


class Skip(Exception):
    """A reason to leave one RAW file in RAW and carry on with the others."""


class Stop(Exception):
    """A reason to stop the whole run before anything is changed."""


@dataclass
class Invoice:
    """One RAW file, read and cleaned up, ready to be written out."""
    raw_path: Path
    tab_name: str
    table_name: str
    headers: list[str]  # every column name, including the three added ones
    n_cols: int  # how many of those are the vendor's
    rows: list[list]  # the vendor's data rows
    formats: list[list[str]]  # the number format for each of those cells
    columns: dict[str, int]  # where SHIPPED, PRICE/UNIT, TAX, FREIGHT and AMOUNT are (from 1)
    invoice_file: str
    notes: list[str] = field(default_factory=list)


@dataclass
class MasterList:
    path: Path
    book: object
    cons_sheet: object
    cons_table: object


def main() -> int:
    if openpyxl is None:
        print("This needs the openpyxl package. Install it with:\n  pip install openpyxl")
        return 1
    try:
        settings = rules.load_settings(SETTINGS_FILE)
    except rules.SettingsError as problem:
        print(problem)
        return 1
    for label, folder in (("RAW folder", settings.raw_folder), ("Invoice folder", settings.invoice_folder)):
        if not folder.is_dir():
            print(f"Can't find the {label}:\n  {folder}\nCheck settings.ini.")
            return 1

    files = rules.list_raw_files(settings.raw_folder)
    if not files:
        print(f"There are no Excel files in RAW:\n  {settings.raw_folder}")
        return 0
    print(f"Found {len(files)} file(s) in RAW:")
    for name in files:
        print(f"  {name}")
    if not ask_yes_no(f"Process these {len(files)} file(s)?"):
        return 0

    try:
        print("Opening the Master List...")
        master = open_master(settings)
        # Any duplicate stops the whole run before anything is changed.
        print("Checking for duplicates...")
        duplicates = find_duplicates(files, settings, master)
    except Stop as problem:
        print(problem)
        return 1
    if duplicates:
        print("\nSTOPPED: found duplicates, so nothing was processed.")
        for line in duplicates:
            print(f"  {line}")
        print("\nTake those files out of RAW, or remove the old copies (from the Master List, "
              "the Done folder or the Invoice folder), then run again.")
        return 1

    done: list[Invoice] = []
    skipped: list[str] = []
    created: list[Path] = []  # Invoice files saved in this run
    for i, name in enumerate(files, start=1):
        print(f"[{i}/{len(files)}] {name}")
        raw_path = settings.raw_folder / name
        # Steps 1-5 only touch this file's own Invoice file, so a problem skips just this file.
        try:
            invoice = read_invoice(raw_path)
            check_master_has_room(master, invoice)
            invoice_path = settings.invoice_folder / invoice.invoice_file
            if invoice_path.exists():  # the duplicate check normally catches this
                raise Skip(f"the Invoice folder already has {invoice.invoice_file}, "
                           "so nothing was changed for this file.")
            save_invoice_workbook(invoice, invoice_path)
            created.append(invoice_path)
        except Skip as reason:
            skipped.append(f"{name}: {reason}")
            continue
        except Exception as error:
            skipped.append(f"{name}: error: {describe_error(error)}")
            continue
        # Steps 6-7 change the Master List. A problem there stops everything, so the
        # Master List is never saved half done.
        try:
            add_to_master(master, invoice)
        except Exception as error:
            remove_files(created)
            print(f"\nSTOPPED: something went wrong adding {name} to the Master List: "
                  f"{describe_error(error)}\nNothing was changed: the Master List wasn't saved, "
                  "and the Invoice files from this run were deleted.")
            return 1
        done.append(invoice)

    if done:
        print("Saving the Master List...")
        try:
            backup = backup_master(master.path)
            save_master(master)
        except Exception as error:
            remove_files(created)
            print(f"\nSTOPPED: couldn't save the Master List: {describe_error(error)}\n"
                  "Nothing was changed: the Invoice files from this run were deleted, and the "
                  "RAW files are still in RAW.")
            return 1
        # ----- 8. Only now that the Master List is saved, move the RAW files to Done -----
        settings.done_folder.mkdir(parents=True, exist_ok=True)
        for invoice in done:
            note = move_to_done(invoice.raw_path, settings.done_folder)
            if note:
                invoice.notes.append(note)

    print(f"\nProcessed {len(done)} file(s).")
    for invoice in done:
        print(f"  {invoice.raw_path.name}  ->  {invoice.invoice_file}")
        for note in invoice.notes:
            print(f"      Note: {note}")
    if skipped:
        print(f"\nLeft in RAW: {len(skipped)} file(s).")
        for line in skipped:
            print(f"  {line}")
    if done:
        print(f"\nThe Master List was saved. A copy of it from before this run is in:\n  {backup}")
    return 1 if skipped else 0


# ---------------------------------------------------------------- reading a RAW file


def read_invoice(raw_path: Path) -> Invoice:
    """Reads a RAW file and cleans it up. Raises Skip when it can't be used."""
    if is_open_in_excel(raw_path):
        raise Skip("it's open in Excel. Close it and run again.")
    if raw_path.suffix.lower() == ".xls":
        raise Skip("it's an old .xls file. Open it in Excel, save it as .xlsx, and run again.")
    try:
        values_book = openpyxl.load_workbook(raw_path, data_only=True)
        formulas_book = openpyxl.load_workbook(raw_path)
    except Exception as error:
        raise Skip(f"it can't be read ({describe_error(error)}).")
    ws = values_book.worksheets[0]
    ws_formulas = formulas_book.worksheets[0]
    tab_name = ws.title

    header_values = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1), ())]
    while header_values and not rules.cell_text(header_values[-1]).strip():
        header_values.pop()
    n_cols = len(header_values)
    if n_cols == 0:
        raise Skip("row 1 has no headers.")

    rows: list[list] = []
    formats: list[list[str]] = []
    for cells, formula_cells in zip(ws.iter_rows(min_row=2, max_col=n_cols),
                                    ws_formulas.iter_rows(min_row=2, max_col=n_cols)):
        values = [cell.value for cell in cells]
        if rules.is_unwanted_row(values):  # "subtotal for", "grand total" and empty rows
            continue
        for cell, formula_cell in zip(cells, formula_cells):
            if cell.value is None and isinstance(formula_cell.value, str) and formula_cell.value.startswith("="):
                raise Skip("it has formulas with no saved results. Open it in Excel, save it, "
                           "and run again.")
        rows.append(values)
        formats.append([cell.number_format for cell in cells])
    if not rows:
        raise Skip("no data rows left after removing subtotal, grand total and empty rows.")

    for values, fmts in zip(rows, formats):
        for i, value in enumerate(values):
            if isinstance(value, str):  # numbers stored as text become real numbers
                parsed = rules.text_to_number(value)
                if parsed is not None:
                    number, decimals = parsed
                    values[i] = int(number) if decimals == 0 else number
                    if fmts[i] == "@":  # a Text format would keep it text
                        fmts[i] = rules.number_format_with_decimals(decimals)
            if rules.is_zero_pad_format(fmts[i]):  # e.g. SHIP TO 0012345
                fmts[i] = rules.number_format_with_decimals(rules.format_decimals(fmts[i]))

    columns, missing = {}, []
    for key, spellings in (("shipped", H_SHIPPED), ("price", H_PRICE), ("tax", H_TAX),
                           ("freight", H_FREIGHT), ("amount", H_AMOUNT)):
        index = rules.find_header(header_values, spellings)
        if index is None:
            missing.append(rules.first_spelling(spellings))
        else:
            columns[key] = index + 1
    if missing:
        raise Skip(f"row 1 has no {', '.join(missing)} column. "
                   f"Row 1 has: {rules.header_list(header_values)}")

    notes = []
    invoice_file, note = rules.invoice_file_name(raw_path.name, tab_name)
    if note:
        notes.append(note)

    # Number formats: Number for the quantities, Accounting for the money
    def format_column(column: int, number_format: str) -> None:
        for fmts in formats:
            fmts[column - 1] = number_format

    format_column(columns["shipped"], FMT_NUMBER)
    ordered = rules.find_header(header_values, H_ORDERED)
    if ordered is not None:
        format_column(ordered + 1, FMT_NUMBER)
    else:
        notes.append(f"no {rules.first_spelling(H_ORDERED)} column, so it wasn't formatted. "
                     f"Row 1 has: {rules.header_list(header_values)}")
    for key in ("price", "tax", "freight", "amount"):
        format_column(columns[key], FMT_ACCOUNTING)

    return Invoice(raw_path=raw_path, tab_name=tab_name, table_name=rules.table_name_for(tab_name),
                   headers=rules.table_headers(header_values, ADDED_HEADERS), n_cols=n_cols,
                   rows=rows, formats=formats, columns=columns, invoice_file=invoice_file, notes=notes)


# ---------------------------------------------------------------- writing the invoice tab


def save_invoice_workbook(invoice: Invoice, path: Path) -> None:
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = invoice.tab_name
    build_invoice_sheet(ws, invoice)
    book.calculation.fullCalcOnLoad = True  # Excel works out the formulas when it opens it
    book.save(path)


def build_invoice_sheet(ws, invoice: Invoice) -> None:
    """Writes the Inv_ table and everything around it onto an empty tab."""
    names, table = invoice.headers, invoice.table_name
    check_col, error_col, v_col = invoice.n_cols + 1, invoice.n_cols + 2, invoice.n_cols + 3
    last_row = 3 + len(invoice.rows)
    amount = names[invoice.columns["amount"] - 1]
    check_number = names[check_col - 1]
    error_letter = get_column_letter(error_col)

    for column, name in enumerate(names, start=1):  # header on row 3
        ws.cell(row=3, column=column, value=name)
    for row, (values, fmts) in enumerate(zip(invoice.rows, invoice.formats), start=4):
        for column, (value, fmt) in enumerate(zip(values, fmts), start=1):
            cell = ws.cell(row=row, column=column)
            write_value(cell, value)
            cell.number_format = fmt
        ws.cell(row=row, column=check_col).value = rules.check_number_formula(
            table, names[invoice.columns["shipped"] - 1], names[invoice.columns["price"] - 1],
            names[invoice.columns["tax"] - 1], names[invoice.columns["freight"] - 1])
        ws.cell(row=row, column=error_col).value = rules.error_formula(table, amount, check_number)
        ws.cell(row=row, column=v_col).value = rules.v_formula(
            f"{error_letter}{row}", table, check_number, amount)
        ws.cell(row=row, column=check_col).number_format = FMT_ACCOUNTING
        ws.cell(row=row, column=v_col).number_format = FMT_ACCOUNTING

    subtotal = ws.cell(row=2, column=invoice.columns["amount"])  # row 2, above AMOUNT
    subtotal.value = rules.subtotal_formula(table, amount)
    subtotal.number_format = FMT_ACCOUNTING

    orange = PatternFill(fill_type="solid", start_color=ORANGE, end_color=ORANGE)
    for column in (check_col, error_col, v_col):
        ws.cell(row=3, column=column).fill = orange
    ws.conditional_formatting.add(
        f"{error_letter}4:{error_letter}{last_row}",
        CellIsRule(operator="equal", formula=['"ERROR"'], font=Font(color=ERROR_TEXT),
                   fill=PatternFill(fill_type="solid", start_color=ERROR_FILL, end_color=ERROR_FILL)))

    inv_table = Table(displayName=table, ref=f"A3:{get_column_letter(v_col)}{last_row}")
    inv_table.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showFirstColumn=False,
                                              showLastColumn=False, showRowStripes=True,
                                              showColumnStripes=False)
    ws.add_table(inv_table)
    set_column_widths(ws, invoice)


def set_column_widths(ws, invoice: Invoice) -> None:
    """Roughly what Excel's AutoFit would do (openpyxl can't measure text)."""
    for column, name in enumerate(invoice.headers, start=1):
        width = len(name.strip()) + 3  # room for the filter button
        if column <= invoice.n_cols:
            for values, fmts in zip(invoice.rows, invoice.formats):
                width = max(width, shown_length(values[column - 1], fmts[column - 1]) + 1)
        else:
            width = max(width, 14 if column != invoice.n_cols + 2 else 8)
        ws.column_dimensions[get_column_letter(column)].width = min(width, 50)


def shown_length(value, number_format: str) -> int:
    if value is None:
        return 0
    if isinstance(value, bool):
        return 5
    if isinstance(value, (int, float)):
        if number_format == FMT_ACCOUNTING:
            return len(f"{value:,.2f}") + 4
        decimals = rules.format_decimals(number_format) if number_format != "General" else 0
        return len(f"{value:.{decimals}f}") if number_format != "General" else len(rules.cell_text(value))
    if isinstance(value, datetime):
        return 10
    return max(len(line) for line in str(value).splitlines() or [""])


def write_value(cell, value) -> None:
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # the vendor's text, not a formula


# ---------------------------------------------------------------- the Master List


def open_master(settings) -> MasterList:
    _, candidates = rules.master_candidates(settings.master_path)
    path = next((Path(c) for c in candidates if Path(c).is_file()), None)
    if path is None:
        raise Stop(f"Can't find the Master List:\n  {settings.master_path}\nCheck settings.ini.")
    if path.suffix.lower() in (".xls", ".xlsb"):
        raise Stop(f"{path.name} is a {path.suffix} file, which this script can't edit. "
                   "Save it as .xlsx (or .xlsm) and run again.")
    if is_open_in_excel(path):
        raise Stop(f"{path.name} is open in Excel. Close it (save it first if you've made changes), "
                   "then run again.")
    if not os.access(path, os.W_OK):
        raise Stop(f"{path.name} is read-only, so it couldn't be saved.")
    try:
        book = openpyxl.load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm", keep_links=True)
    except Exception as error:
        raise Stop(f"Can't read {path.name}: {describe_error(error)}")
    if book.security is not None and book.security.lockStructure:
        raise Stop(f"{path.name} has its structure protected, so tabs can't be added.\n"
                   "Unprotect it (Review > Protect Workbook), save it, and run again.")
    found = find_table(book, CONSUMABLES_TABLE)
    if found is None:
        raise Stop(f"{path.name} has no table named {CONSUMABLES_TABLE} (it should be on the "
                   f"{CONSUMABLES_TAB} tab), so nothing was processed.")
    return MasterList(path=path, book=book, cons_sheet=found[0], cons_table=found[1])


def find_table(book, table_name: str):
    """(sheet, table) for the table with this name anywhere in the workbook, or None."""
    for ws in book.worksheets:
        for table in ws.tables.values():  # (openpyxl's tables.items() gives addresses, not tables)
            if table.displayName.lower() == table_name.lower():
                return ws, table
    return None


def find_duplicates(files: list[str], settings, master: MasterList) -> list[str]:
    """Checks every RAW file before anything is changed. Returns one line per duplicate
    file, listing every reason. A file is a duplicate when:
      - the Done folder already has a file with the same name,
      - the Master List already has its tab or its Inv_ table,
      - the Consumables table already has rows for its tab name,
      - the Invoice folder already has the Invoice file it would make, or
      - another RAW file in this run has the same tab name."""
    sheet_names = {name.lower() for name in master.book.sheetnames}
    consumables_tabs = consumables_column_values(master, CONS_TAB_NAME_COLUMN)
    lines, seen = [], set()
    for name in files:
        reasons = []
        if (settings.done_folder / name).exists():
            reasons.append("the Done folder already has a file with this name")
        tab_name = raw_tab_name(settings.raw_folder / name)  # "" when it can't be read (skipped later)
        if tab_name:
            if tab_name.lower() in sheet_names:
                reasons.append(f"the Master List already has a tab named {tab_name}")
            elif find_table(master.book, rules.table_name_for(tab_name)) is not None:
                reasons.append(f"the Master List already has a table named {rules.table_name_for(tab_name)}")
            if rules.column_has_value(consumables_tabs, tab_name):
                reasons.append(f"the {CONSUMABLES_TABLE} table already has rows for {tab_name}")
            invoice_file, _ = rules.invoice_file_name(name, tab_name)
            if (settings.invoice_folder / invoice_file).exists():
                reasons.append(f"the Invoice folder already has {invoice_file}")
            if tab_name.lower() in seen:
                reasons.append(f"another file in RAW is also tab {tab_name}")
            seen.add(tab_name.lower())
        if reasons:
            lines.append(f"{name}: " + "; ".join(reasons))
    return lines


def raw_tab_name(path: Path) -> str:
    """The name of a RAW file's tab, or "" if the file can't be read."""
    if path.suffix.lower() == ".xls" or is_open_in_excel(path):
        return ""
    try:
        book = openpyxl.load_workbook(path, read_only=True)
        try:
            return book.sheetnames[0]
        finally:
            book.close()
    except Exception:
        return ""


@dataclass
class TableShape:
    min_col: int
    min_row: int
    max_col: int
    max_row: int
    first_data: int
    last_data: int
    totals: int


def table_shape(table) -> TableShape:
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    header = 1 if table.headerRowCount is None else table.headerRowCount
    totals = table.totalsRowCount or 0
    return TableShape(min_col, min_row, max_col, max_row, min_row + header, max_row - totals, totals)


def consumables_column_values(master: MasterList, table_column: int) -> list:
    shape = table_shape(master.cons_table)
    column = shape.min_col + table_column - 1
    return [master.cons_sheet.cell(row=r, column=column).value
            for r in range(shape.first_data, shape.last_data + 1)]


def consumables_plan(master: MasterList, invoice: Invoice) -> tuple[int, int]:
    """(first row to write, how many rows the table grows by) for adding this invoice."""
    ws, shape = master.cons_sheet, table_shape(master.cons_table)
    tab_col = shape.min_col + CONS_TAB_NAME_COLUMN - 1
    last_written_col = shape.min_col + CONS_FIRST_DATA_COLUMN - 2 + invoice.n_cols
    # A table with no data yet has one empty row: fill that one first.
    blank_first_row = shape.last_data == shape.first_data and all(
        ws.cell(row=shape.first_data, column=c).value in (None, "")
        for c in range(tab_col, last_written_col + 1))
    start = shape.first_data if blank_first_row else shape.last_data + 1
    grow = max(0, start + len(invoice.rows) - 1 - shape.last_data)
    return start, grow


def check_master_has_room(master: MasterList, invoice: Invoice) -> None:
    """Raises Skip when this invoice can't go into the Master List."""
    if invoice.tab_name.lower() in {name.lower() for name in master.book.sheetnames}:
        raise Skip(f"the Master List already has a tab named {invoice.tab_name}.")
    if find_table(master.book, invoice.table_name) is not None:
        raise Skip(f"the Master List already has a table named {invoice.table_name}.")
    shape = table_shape(master.cons_table)
    have = shape.max_col - shape.min_col + 1
    needed = CONS_FIRST_DATA_COLUMN - 1 + invoice.n_cols
    if have < needed:
        raise Skip(f"the {CONSUMABLES_TABLE} table has {have} columns, but this file's "
                   f"{invoice.n_cols} columns need {needed} (they start in its column "
                   f"{CONS_FIRST_DATA_COLUMN}).")
    _, grow = consumables_plan(master, invoice)
    for row in range(shape.max_row + 1, shape.max_row + grow + 1):
        for column in range(shape.min_col, shape.max_col + 1):
            if master.cons_sheet.cell(row=row, column=column).value not in (None, ""):
                raise Skip(f"there's something in the rows right below the {CONSUMABLES_TABLE} "
                           f"table (cell {get_column_letter(column)}{row}), so its rows can't be added. "
                           "Move it out of the way and run again.")


def add_to_master(master: MasterList, invoice: Invoice) -> None:
    # ----- 6. The same tab at the end of the Master List, with an orange tab -----
    ws = master.book.create_sheet(title=invoice.tab_name)
    build_invoice_sheet(ws, invoice)
    ws.sheet_properties.tabColor = ORANGE
    # ----- 7. The data rows at the bottom of the Consumables table -----
    append_to_consumables(master, invoice)


def append_to_consumables(master: MasterList, invoice: Invoice) -> None:
    """Adds the invoice's data rows (the vendor's columns only) to the bottom of the
    Consumables table: the tab name in its 4th column and the vendor's columns from
    its 5th. New rows get the formatting and formulas of the row above, like Excel
    does when a table grows."""
    ws, table = master.cons_sheet, master.cons_table
    shape = table_shape(table)
    start, grow = consumables_plan(master, invoice)
    first_col, last_col = shape.min_col, shape.max_col
    tab_col = first_col + CONS_TAB_NAME_COLUMN - 1
    data_col = first_col + CONS_FIRST_DATA_COLUMN - 1

    if grow and shape.totals:  # move the totals row down out of the way
        ws.move_range(f"{get_column_letter(first_col)}{shape.max_row}:"
                      f"{get_column_letter(last_col)}{shape.max_row}", rows=grow, translate=True)

    template = shape.last_data  # the row new rows copy their formatting and formulas from
    columns = list(table.tableColumns)
    for i, (values, fmts) in enumerate(zip(invoice.rows, invoice.formats)):
        row = start + i
        for column in range(first_col, last_col + 1):
            cell = ws.cell(row=row, column=column)
            source = ws.cell(row=template, column=column)
            if row != template:
                cell._style = copy(source._style)
            if column == tab_col or data_col <= column < data_col + invoice.n_cols:
                continue
            definition = columns[column - first_col] if column - first_col < len(columns) else None
            formula = getattr(getattr(definition, "calculatedColumnFormula", None), "attr_text", None)
            if formula:
                cell.value = "=" + formula
            elif row != template and isinstance(source.value, str) and source.value.startswith("="):
                cell.value = Translator(source.value, origin=source.coordinate).translate_formula(cell.coordinate)
        ws.cell(row=row, column=tab_col).value = rules.as_number_if_whole(invoice.tab_name)
        for j in range(invoice.n_cols):
            cell = ws.cell(row=row, column=data_col + j)
            write_value(cell, values[j])
            cell.number_format = fmts[j]

    if grow:
        new_max_row = shape.max_row + grow
        table.ref = f"{get_column_letter(first_col)}{shape.min_row}:{get_column_letter(last_col)}{new_max_row}"
        if table.autoFilter is not None:
            table.autoFilter.ref = (f"{get_column_letter(first_col)}{shape.min_row}:"
                                    f"{get_column_letter(last_col)}{new_max_row - shape.totals}")
        extend_ranges(ws, shape, grow)


def extend_ranges(ws, shape: TableShape, grow: int) -> None:
    """Conditional formatting and data validation that ran to the table's last row
    follow it down, like they would in Excel."""
    def extended(sqref) -> MultiCellRange:
        ranges = []
        for rng in MultiCellRange(str(sqref)).ranges:
            rng = CellRange(rng.coord)
            if rng.max_row == shape.last_data and rng.min_col >= shape.min_col and rng.max_col <= shape.max_col:
                rng.expand(down=grow)
            ranges.append(rng.coord)
        return MultiCellRange(" ".join(ranges))

    try:
        # openpyxl keeps conditional formatting in a dictionary keyed by its range, so
        # it's rebuilt with the new ranges instead of changing the ranges in place.
        rebuilt = OrderedDict()
        for rule_set, rule_list in ws.conditional_formatting._cf_rules.items():
            key = ConditionalFormatting(sqref=extended(rule_set.sqref), pivot=rule_set.pivot)
            rebuilt.setdefault(key, []).extend(rule_list)
        ws.conditional_formatting._cf_rules = rebuilt
    except Exception:
        pass  # nice to have; never worth stopping for
    try:
        for validation in ws.data_validations.dataValidation:
            validation.sqref = extended(validation.sqref)
    except Exception:
        pass


def backup_master(path: Path) -> Path:
    folder = path.parent / "Backups"
    folder.mkdir(exist_ok=True)
    target = folder / f"{path.stem} {datetime.now():%Y-%m-%d %H%M%S}{path.suffix}"
    shutil.copy2(path, target)
    return target


def save_master(master: MasterList) -> None:
    """Saves to a temporary file first, then swaps it in, so a failed save can't
    leave a broken Master List behind."""
    master.book.calculation.fullCalcOnLoad = True  # Excel works out the new formulas when it opens it
    handle, temp_name = tempfile.mkstemp(dir=master.path.parent, suffix=master.path.suffix)
    os.close(handle)
    try:
        master.book.save(temp_name)
        os.replace(temp_name, master.path)
    finally:
        if os.path.exists(temp_name):
            os.remove(temp_name)


# ---------------------------------------------------------------- files


def is_open_in_excel(path: Path) -> bool:
    """Excel keeps a hidden "~$" lock file next to a workbook while it's open."""
    return any((path.parent / f"~${name}").exists() for name in (path.name, path.name[2:]))


def move_to_done(raw_path: Path, done_folder: Path) -> str:
    """Moves the RAW file into the Done folder (adds the date and time if that name is
    taken). Returns a note if it couldn't."""
    target = done_folder / raw_path.name
    if target.exists():
        target = done_folder / f"{raw_path.stem} {datetime.now():%Y-%m-%d %H%M%S}{raw_path.suffix}"
    try:
        shutil.move(str(raw_path), str(target))
        return ""
    except OSError as error:
        return f"couldn't move the RAW file to Done ({error}), so move it yourself"


def remove_files(paths: list[Path]) -> None:
    for path in paths:
        try:
            path.unlink()
        except OSError:
            pass


def describe_error(error: Exception) -> str:
    return str(error) or type(error).__name__


def ask_yes_no(question: str) -> bool:
    while True:
        answer = input(f"{question} (y/n) ").strip().lower()
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no", ""):
            return False


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nStopped.")
        sys.exit(1)
