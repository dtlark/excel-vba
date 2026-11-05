# VBA macros

| File | What it does |
|---|---|
| `RawInvoices.bas` | Turns the vendor's RAW files into Invoice workbooks and adds each one to the Master List. |
| `PdfDiagnose.bas` | Shows what Word's PDF conversion produces (PDF experiment). |

## RawInvoices.bas

For each Excel file in the RAW folder, the `ProcessRawInvoices` macro:

1. Copies the header and data rows to a new workbook, with the header on row 3.
   Rows containing "subtotal for" or "grand total", and empty rows, are left out.
   Only values and number formats are copied, so the vendor's colors and fonts
   don't cover up the table style. Numbers stored as text become real numbers.
2. Makes it a table named `Inv_<Tab Name>` in **White, Table Style Medium 1**
   (set by `TABLE_STYLE` near the top of the code), adds the **Check Number**, blank and
   **V** columns with orange headers, and highlights ERROR cells in red.
3. Puts `=SUBTOTAL(9, Inv_<Tab Name>[AMOUNT])` on row 2 above AMOUNT.
   Formats SHIPPED and ORDERED as **Number** (2 decimals), and PRICE/UNIT, TAX,
   FREIGHT, AMOUNT, Check Number, V and that subtotal as **Accounting**.
4. Saves it as `Invoice <Tab Name> (<Postfix>).xlsx` in the Invoice folder.
5. Copies the tab to the end of the Master List.
6. Adds the table's data rows to the bottom of the Master List's **Consumables**
   table (on the Consolidated Consumables tab). Only the vendor's columns go in,
   not Check Number, the blank column or V. The tab name goes in the table's 4th
   column, and the vendor's columns start in its 5th.
7. Moves the RAW file into the Done folder.

You only need Excel open with the macro workbook. The macro opens and closes the
RAW and Invoice files itself. The Master List can be open or closed: if it's
closed, the macro opens it. The macro never saves the Master List, so you can
check the new tabs and Consumables rows first.

### Set up (once)

1. In Excel, create a blank workbook and save it as **Excel Macro-Enabled
   Workbook (\*.xlsm)**, for example `Invoice Macros.xlsm`. Save it anywhere
   except the RAW folder.
2. Press **Alt+F11**, then click **Insert → Module**.
3. Paste in the contents of `RawInvoices.bas`. On GitHub, open the file and
   click **Copy raw file**. (Or use **File → Import File** with the downloaded file.)
4. Close the VBA editor. Press **Alt+F8**, pick **ProcessRawInvoices**, and
   click **Run**. The first time, it adds a **Settings** sheet and stops.
5. On the Settings sheet, fill in column B:

   | Setting | What to put |
   |---|---|
   | Fiscal year | Just the number, like `26`. See below |
   | RAW folder | The folder you save the vendor's Excel files into |
   | Done folder | Where processed RAW files go. Leave blank for a `Done` folder inside the RAW folder |
   | Invoice folder | The folder the `Invoice <Tab Name>` files are saved in |
   | Master List file | The Master List's full path. The `.xlsx` can be left off |

   **Fiscal year in paths:** anywhere a path says `{FY}`, the macro puts `FY`
   and the Fiscal year number. With Fiscal year `26`,
   `C:\Invoices\{FY}\{FY} Master List` means `C:\Invoices\FY26\FY26 Master List`.
   Next year, change Fiscal year to `27` and every path follows.

   To copy a path, right-click the folder or file in File Explorer and choose
   **Copy as path** (on Windows 10, hold Shift while right-clicking). The quotes
   it adds are removed automatically.
6. Save the workbook.

### Each time

1. Save the vendor's Excel files into the RAW folder.
2. Open `Invoice Macros.xlsm`. If a yellow bar says macros are disabled, click
   **Enable Content**.
3. Press **Alt+F8**, pick **ProcessRawInvoices**, and click **Run**. It lists
   the files it found and asks before starting.
4. Read the message at the end. Any file that couldn't be processed is still in
   RAW, with the reason.
5. Check the new tabs at the end of the Master List and the new rows at the
   bottom of the Consumables table, then save it.

### First run

Test with one file before trusting it with a batch:

1. Make a backup copy of the Master List.
2. Leave just one file in RAW and run the macro.
3. Compare the Invoice file and the new Master List tab with what you'd make by hand.

To undo a run: close the Master List without saving (that also undoes the
Consumables rows), move the RAW file back out of `Done`, and delete the Invoice
file.

### Duplicates stop the run

Before changing anything, the macro checks every file in RAW. If any of them is
a duplicate, it stops with an error that lists them, and nothing in that run is
processed. A file is a duplicate when:

- the Done folder already has a file with the same name,
- the Master List already has its tab, or its `Inv_` table,
- the Consumables table already has rows with its tab name in the 4th column, or
- another file in RAW has the same tab name.

Take those files out of RAW (or remove the old copies), then run again.

### When a file is skipped

The file stays in RAW, the other files are still processed, and the message at
the end says why:

- Row 1 is missing SHIPPED, PRICE/UNIT, TAX, FREIGHT (or FREGHT) or AMOUNT.
- The Consumables table doesn't have enough columns for the file's data.
- The file is open in Excel.
- There are no data rows.

If the Master List has no table named Consumables, the macro stops before
processing anything.

### Updating the macro

When there's a new version, press **Alt+F11**, click in the module, select all
the code (**Ctrl+A**), paste the new version over it, and save. Your paths stay
on the Settings sheet, so there's nothing to re-enter. For a new fiscal year,
change the Fiscal year number on the Settings sheet (if your paths use `{FY}`).
If your Settings sheet was made before the Fiscal year setting existed, the
macro adds that row at the bottom the next time it runs.

### Details

- **Invoice file name:** the part of the RAW file name after the tab name goes
  in parentheses, without leading spaces, `_` or `-`. `123456 ABC.xlsx` and
  `123456 - ABC.xlsx` both become `Invoice 123456 (ABC).xlsx`. If nothing comes
  after the tab name, there are no parentheses: `123456.xlsx` becomes
  `Invoice 123456.xlsx`.
- **Numbers stored as text:** text like `12.50`, `1,234`, `$12.50`, `-12.50`,
  `12.50-` or `(12.50)` becomes a real number, in every column. Some text is
  left alone on purpose: values with leading zeros like `00123` (usually codes,
  and converting would drop the zeros), anything over 15 digits (too long for
  Excel to store exactly), dates, and percents. The message at the end says how
  many cells were converted.
- **ERROR check:** AMOUNT and Check Number are compared to 6 decimal places, so
  computer rounding noise (like 0.00000000000001) isn't flagged. Any real
  difference is, even a fraction of a cent.
- **Columns are found by header name**, ignoring capitals, extra spaces,
  non-breaking spaces and line breaks, so it still works if the vendor moves a
  column. If a column can't be found, the message lists the headers row 1 does
  have. The names are near the top of the code. To accept another spelling,
  add it after a `|` on that column's line:

  ```vb
  Private Const H_FREIGHT As String = "FREIGHT|FREGHT"
  ```

## PdfDiagnose.bas

This macro opens a PDF in Word and copies exactly what Word made into Excel,
without fixing anything, so you can see what goes wrong in the conversion.

### Set up (once)

1. Download `PdfDiagnose.bas` (open it on GitHub, then click **Download raw file**).
2. In Excel, create a blank workbook and save it as **Excel Macro-Enabled Workbook (\*.xlsm)**.
3. Press **Alt+F11** to open the VBA editor.
4. Click **File → Import File…** and pick `PdfDiagnose.bas`.
   (Or click **Insert → Module** and paste in the file's contents.)
5. Close the VBA editor.

### Run it

1. Press **Alt+F8**, pick **DiagnosePdfConversion**, and click **Run**.
2. Pick the PDF. Word opens and converts it, which can take a minute. If Word asks
   "Word will now convert your PDF to an editable Word document…", click **OK**.
   The question can show up behind the Excel window.
3. When it finishes, the workbook has three new sheets:
   - **Summary**: how many tables Word made, how many rows each has, and how many problem rows and cells.
   - **Raw Tables**: every Word table, one Excel cell per Word cell. Column C counts the cells in each row.
     - **Pink**: a row with fewer cells than the fullest row (usually merged cells).
     - **Yellow**: a cell holding more than one line (often two rows or values jammed together).
   - **Raw Text**: any text Word did *not* put in a table, split into columns at tab stops.

Word stays open on the converted document so you can compare the two. Close it
without saving. To have Word close automatically, change `SHOW_WORD` to `False`
at the top of the code.

Try `sample-table.pdf` first. Its data is made up, so you can share all of the
output.
