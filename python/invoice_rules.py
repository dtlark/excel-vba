"""The rules for turning a vendor's RAW invoice file into an Invoice workbook.

Nothing in here reads or writes files except load_settings and list_raw_files.
It all works on plain Python values, so it's easy to test. raw_invoices.py uses
these rules, and they match the VBA macro (vba/RawInvoices.bas) rule for rule.
"""
from __future__ import annotations

import configparser
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

# ---------------------------------------------------------------- settings

FY_TAG = "{FY}"
_FY_PATTERN = re.compile(re.escape(FY_TAG), re.IGNORECASE)


class SettingsError(Exception):
    """The settings file is missing, or something in it needs filling in."""


@dataclass
class Settings:
    raw_folder: Path
    done_folder: Path
    invoice_folder: Path
    master_path: str  # may or may not end in .xlsx


def load_settings(path: Path) -> Settings:
    """Reads settings.ini: the same five settings as the macro's Settings tab."""
    if not path.is_file():
        raise SettingsError(
            f"Can't find the settings file:\n  {path}\n"
            "Copy settings.example.ini to settings.ini in the same folder and fill it in."
        )
    parser = configparser.ConfigParser(interpolation=None, delimiters=("=",))
    parser.read(path, encoding="utf-8-sig")
    if not parser.has_section("Settings"):
        raise SettingsError(f"{path.name} needs a [Settings] line above the settings.")
    section = parser["Settings"]

    fiscal_year = normalize_fiscal_year(clean_setting(section.get("Fiscal year", "")))
    values, uses_fy = {}, False
    for label in ("RAW folder", "Done folder", "Invoice folder", "Master List file"):
        values[label], used = with_fiscal_year(clean_setting(section.get(label, "")), fiscal_year)
        uses_fy = uses_fy or used

    missing = []
    if uses_fy and not fiscal_year:
        missing.append(f"Fiscal year (a path uses {FY_TAG})")
    if not values["RAW folder"]:
        missing.append("RAW folder")
    if not values["Invoice folder"]:
        missing.append("Invoice folder")
    if "\\" not in values["Master List file"] and "/" not in values["Master List file"]:
        missing.append("Master List file (the full path)")
    if missing:
        raise SettingsError(
            f"Fill these in in {path.name}, then run again:\n" + "\n".join(f"  {m}" for m in missing)
        )

    raw_folder = Path(values["RAW folder"])
    done_folder = Path(values["Done folder"]) if values["Done folder"] else raw_folder / "Done"
    return Settings(raw_folder, done_folder, Path(values["Invoice folder"]), values["Master List file"])


def clean_setting(value: str) -> str:
    """Drops the quotes "Copy as path" adds, and a final \\ or /."""
    text = str(value or "").replace('"', "").strip()
    return text.rstrip("\\/")


def normalize_fiscal_year(value: str) -> str:
    """'26' and 'FY26' both mean fiscal year 26."""
    text = value.strip()
    if text[:2].upper() == "FY":
        text = text[2:].strip()
    return text


def with_fiscal_year(path: str, fiscal_year: str) -> tuple[str, bool]:
    """Puts "FY" + the fiscal year wherever the path says {FY}. Also says whether it did."""
    used = bool(_FY_PATTERN.search(path))
    return _FY_PATTERN.sub(lambda _: "FY" + fiscal_year, path), used


# ---------------------------------------------------------------- file names


def list_raw_files(folder: Path) -> list[str]:
    """The Excel files in the RAW folder, by name. Skips Excel's "~$" lock files."""
    names = [
        p.name
        for p in folder.iterdir()
        if p.is_file() and ".xls" in p.suffix.lower() and not p.name.startswith("~$")
    ]
    return sorted(names, key=str.lower)


def name_without_extension(file_name: str) -> str:
    dot = file_name.rfind(".")
    return file_name[:dot] if dot >= 0 else file_name


def master_candidates(master_path: str) -> tuple[str, list[str]]:
    """The Master List path without its extension, and the files to try, in order."""
    dot = master_path.rfind(".")
    separator = max(master_path.rfind("\\"), master_path.rfind("/"))
    if dot > separator and master_path[dot:].lower().startswith(".xls"):
        base, first = master_path[:dot], [master_path]
    else:
        base, first = master_path, []
    return base, first + [base + ext for ext in (".xlsx", ".xlsm", ".xlsb", ".xls")]


def table_name_for(tab_name: str) -> str:
    """'123456' -> 'Inv_123456'. Table names allow only letters, numbers, _ and ."""
    return "Inv_" + re.sub(r"[^A-Za-z0-9_.]", "_", tab_name)


def raw_postfix(raw_name: str, tab_name: str) -> tuple[str, str]:
    """The part of the RAW file name after the tab name, and a note if there's a problem.

    '123456 Oct.xlsx' -> 'Oct', '123456 - Oct.xlsx' -> 'Oct', '123456.xlsx' -> ''
    """
    base = name_without_extension(raw_name)
    note = ""
    if base[: len(tab_name)].lower() == tab_name.lower():
        rest = base[len(tab_name):]
    else:
        rest = base
        note = (
            f"the file name doesn't start with its tab name ({tab_name}), "
            "so the whole file name was used"
        )
    return rest.lstrip(" _-").strip(" "), note


def invoice_file_name(raw_name: str, tab_name: str) -> tuple[str, str]:
    """'Invoice 123456 (ABC).xlsx', with no () when nothing follows the tab name."""
    postfix, note = raw_postfix(raw_name, tab_name)
    name = f"Invoice {tab_name}"
    if postfix:
        if not (postfix.startswith("(") and postfix.endswith(")")):
            postfix = f"({postfix})"  # don't double up "(ABC)"
        name += f" {postfix}"
    return name + ".xlsx", note


# ---------------------------------------------------------------- rows and headers


def cell_text(value) -> str:
    """A cell value as Excel would show it in plain text (12.0 -> '12')."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def is_unwanted_row(values: list) -> bool:
    """True for rows with "subtotal for" or "grand total" in any cell, and empty rows."""
    has_value = False
    for value in values:
        text = cell_text(value).strip(" ").lower()
        if text:
            has_value = True
        if "subtotal for" in text or "grand total" in text:
            return True
    return not has_value


def clean_header(value) -> str:
    """Header text in capitals, with line breaks, tabs and non-breaking spaces as
    plain spaces, and no doubled, leading or trailing spaces."""
    text = cell_text(value)
    for ch in ("\r", "\n", "\t", "\xa0"):
        text = text.replace(ch, " ")
    while "  " in text:
        text = text.replace("  ", " ")
    return text.strip(" ").upper()


def find_header(headers: list, spellings: str) -> int | None:
    """Position (from 0) of the first header matching any spelling, like 'FREIGHT|FREGHT'."""
    cleaned = [clean_header(h) for h in headers]
    for spelling in spellings.split("|"):
        target = clean_header(spelling)
        if target in cleaned:
            return cleaned.index(target)
    return None


def first_spelling(spellings: str) -> str:
    return spellings.split("|")[0]


def header_list(headers: list) -> str:
    """The headers for messages, like 'ITEM, DESCRIPTION, SHIPPED, ...'"""
    return ", ".join(clean_header(h) for h in headers)


def column_has_value(values: list, tab_name: str) -> bool:
    """True when a column holds this tab name, as a number or as text."""
    target = tab_name.strip().lower()
    return any(cell_text(v).strip().lower() == target for v in values)


# ---------------------------------------------------------------- numbers

_DIGITS = re.compile(r"[0-9]+")


def _all_digits(text: str) -> bool:
    return bool(_DIGITS.fullmatch(text))


def text_to_number(text: str) -> tuple[float, int] | None:
    """(number, decimal places) for a number written as text, or None.

    Handles spaces (also tabs, line breaks and invisible spaces), a $ sign, commas
    between thousands, and negatives written as -12.50, 12.50- or (12.50). Leading
    zeros are dropped: '00123' becomes 123. Stays text: more than 15 digits (too
    long for Excel to hold exactly), dates, percents and anything else.
    """
    s = text
    for ch in ("\t", "\r", "\n", "\xa0"):
        s = s.replace(ch, " ")
    s = s.replace("​", "").replace("﻿", "").strip(" ")
    if s.count("$") > 1:
        return None
    s = s.replace("$", "").strip(" ")
    negative = False
    if len(s) >= 2 and s[0] == "(" and s[-1] == ")":  # (12.50)
        negative, s = True, s[1:-1].strip(" ")
    elif len(s) >= 2 and s[-1] == "-":  # 12.50-
        negative, s = True, s[:-1].strip(" ")
    elif len(s) >= 2 and s[0] in "-+":  # -12.50
        negative, s = s[0] == "-", s[1:].strip(" ")
    if not s:
        return None

    # What's left may only be digits, commas between groups of 3, and one dot.
    if "." in s:
        dot = s.index(".")
        int_part, frac_part = s[:dot], s[dot + 1:]
        if frac_part and not _all_digits(frac_part):
            return None
    else:
        int_part, frac_part = s, ""
    if "," in int_part:
        groups = int_part.split(",")
        if len(groups[0]) > 3 or not _all_digits(groups[0]):
            return None
        if any(len(g) != 3 or not _all_digits(g) for g in groups[1:]):
            return None
        int_part = int_part.replace(",", "")
    elif int_part and not _all_digits(int_part):
        return None
    if not int_part and not frac_part:
        return None
    if len(int_part) + len(frac_part) > 15:  # too long for Excel
        return None

    number = float(f"{int_part or '0'}.{frac_part or '0'}")
    return (-number if negative else number), len(frac_part)


def first_format_section(number_format: str) -> str:
    """The part of a number format before the first ; without a leading [Red] or [$-409]."""
    fmt = number_format.split(";")[0]
    while fmt.startswith("[") and "]" in fmt:
        fmt = fmt[fmt.index("]") + 1:]
    return fmt


def is_zero_pad_format(number_format) -> bool:
    """True for a format that pads numbers with leading zeros, like '0000000'."""
    return isinstance(number_format, str) and first_format_section(number_format)[:2] == "00"


def format_decimals(number_format: str) -> int:
    """How many decimal places a number format shows: '0000.00' -> 2, '0000000' -> 0."""
    fmt = first_format_section(number_format)
    if "." not in fmt:
        return 0
    count = 0
    for ch in fmt[fmt.index(".") + 1:]:
        if ch not in "0#":
            break
        count += 1
    return count


def number_format_with_decimals(decimals: int) -> str:
    """Excel's Number format with this many decimal places: 0 -> '0', 2 -> '0.00'."""
    return "0" if decimals <= 0 else "0." + "0" * decimals


# ---------------------------------------------------------------- formulas


def escape_column_name(name: str) -> str:
    """Table formulas need a ' in front of [ ] # and ' when they appear in a column name."""
    return name.replace("'", "''").replace("[", "'[").replace("]", "']").replace("#", "'#")


def this_row(table: str, column: str) -> str:
    """A this-row reference to a table column, the way it's stored in the file:
    Inv_123456[[#This Row],[PRICE/UNIT]]. Excel shows it as [@[PRICE/UNIT]]."""
    return f"{table}[[#This Row],[{escape_column_name(column)}]]"


def check_number_formula(table: str, shipped: str, price: str, tax: str, freight: str) -> str:
    """Check Number = (SHIPPED x PRICE/UNIT) + TAX + FREIGHT"""
    return (f"={this_row(table, shipped)}*{this_row(table, price)}"
            f"+{this_row(table, tax)}+{this_row(table, freight)}")


def error_formula(table: str, amount: str, check_number: str) -> str:
    """ERROR when AMOUNT and Check Number differ. Compared to 6 decimal places so
    computer rounding noise isn't flagged; any real difference is, even a fraction
    of a cent."""
    return (f'=IF(ROUND({this_row(table, amount)}-{this_row(table, check_number)},6)=0,'
            f'"","ERROR")')


def v_formula(error_cell: str, table: str, check_number: str, amount: str) -> str:
    """V = Check Number - AMOUNT on ERROR rows. The blank column's name is a space,
    so V points at its cell (like P4) instead of using its name."""
    return (f'=IF({error_cell}="ERROR",'
            f'{this_row(table, check_number)}-{this_row(table, amount)},"")')


def subtotal_formula(table: str, amount: str) -> str:
    return f"=SUBTOTAL(9,{table}[{escape_column_name(amount)}])"


# ---------------------------------------------------------------- tables


def table_headers(vendor_headers: list, added: list[str]) -> list[str]:
    """The column names for the Inv_ table, the way Excel would name them: a blank
    vendor header becomes Column1, Column2..., and a repeated name gets a number
    (AMOUNT, AMOUNT2). The added names (like the blank " " column) are kept as is,
    unless they repeat a vendor name."""
    names: list[str] = []
    used: set[str] = set()

    def unique(name: str) -> str:
        candidate, n = name, 1
        while candidate.lower() in used:
            n += 1
            candidate = f"{name}{n}"
        used.add(candidate.lower())
        return candidate

    for position, value in enumerate(vendor_headers, start=1):
        text = cell_text(value)
        names.append(unique(text if text.strip() else f"Column{position}"))
    for name in added:
        names.append(unique(name))
    return names


def as_number_if_whole(text: str):
    """'123456' -> 123456 (so a tab name lands as a number), anything else unchanged."""
    return int(text) if _all_digits(text) and len(text) <= 15 else text
