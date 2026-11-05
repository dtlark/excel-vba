# Python version

`raw_invoices.py` does the same job as the VBA macro
([vba/RawInvoices.bas](../vba/RawInvoices.bas)), but without Excel. It reads and
writes the `.xlsx` files directly, so it's fast and Excel doesn't need to be
running. The rules are the same: subtotal and grand total rows removed, numbers
stored as text converted, the `Inv_` table with Check Number, ERROR and V, the
Invoice file, the new tab in the Master List, the Consumables rows, duplicates
stopping the run, and so on.

## How it differs from the macro

- **Close the Master List in Excel before you run it.** If the Master List is
  open, the script stops before doing anything.
- **It saves the Master List itself** at the end, after copying it into a
  `Backups` folder next to the Master List. (The macro leaves it open and
  unsaved instead.)
- **It can't keep everything Excel can.** If the Master List has Power Query
  queries, slicers, buttons or shapes, they may be lost when the script saves
  it. If that happens, put the backup back and use the macro instead.
- **A run never leaves things half done.** RAW files only move to Done after
  the Master List is saved. If something goes wrong partway through, the
  Master List isn't saved, and the Invoice files from that run are deleted.
- If there's anything in the rows right below the Consumables table, that file
  is skipped. Excel would push it down, but the script won't move it.
- Old `.xls` RAW files are skipped. Open them in Excel and save them as `.xlsx`.
- Formulas show their results once the files are opened in Excel.

## Set up (once)

1. Install Python 3.9 or newer, from your company's software center or
   [python.org](https://www.python.org/downloads/). On python.org's installer,
   tick **Add python.exe to PATH**.
2. Get this folder onto your computer: on GitHub, click **Code → Download ZIP**,
   unzip it, and keep the `python` folder somewhere handy.
3. Open the `python` folder in File Explorer, type `cmd` in the address bar and
   press Enter. In the window that opens, run:

   ```
   pip install -r requirements.txt
   ```

   That installs `openpyxl`, the library that reads and writes Excel files.
4. Copy `settings.example.ini` to `settings.ini` in the same folder, open it in
   Notepad, and fill in your paths. They're the same as on the macro's Settings
   tab, including `{FY}` and the Fiscal year. Save it.

## Each time

1. Save the vendor's Excel files into the RAW folder.
2. Close the Master List in Excel.
3. Double-click **Run RAW invoices.bat** (or run `python raw_invoices.py`).
4. It lists the files it found. Type `y` and press Enter to go ahead.
5. Read the summary at the end, then open the Master List to check the new
   tabs and Consumables rows.

## Settings

`settings.ini` holds the same five settings as the macro's Settings tab:

| Setting | What to put |
|---|---|
| Fiscal year | Just the number, like `26`. Anywhere a path says `{FY}`, the script puts `FY26` |
| RAW folder | The folder you save the vendor's Excel files into |
| Done folder | Where processed RAW files go. Leave blank for a `Done` folder inside the RAW folder |
| Invoice folder | The folder the `Invoice <Tab Name>` files are saved in |
| Master List file | The Master List's full path. The `.xlsx` can be left off |

`settings.ini` isn't uploaded to GitHub, so your paths stay on your computer.

The column names (`SHIPPED`, `FREIGHT|FREGHT`, `ORDERED`, ...), the table style,
the number formats and the Consumables columns are near the top of
`raw_invoices.py`, the same as in the macro.

## Tests

The tests build made-up RAW files and a made-up Master List, run the whole
script on them, and check the results. They don't need Excel:

```
python -m unittest discover -s tests
```
