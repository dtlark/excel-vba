# VBA: PDF → Word → Excel

Word can open a PDF and convert it to an editable document, but the table comes
out wrong. If it comes out wrong the same way every time, a macro can undo it.

## Step 1: see what the conversion produces (`PdfDiagnose.bas`)

This macro opens a PDF in Word and copies exactly what Word made into Excel,
without fixing anything, so you can see what goes wrong.

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

### What to report back

- The numbers on the **Summary** sheet.
- What a broken row looks like next to how it should look. For example: "the
  header row has 8 cells but data rows have 7, and cell F holds both Qty and
  Unit Price".

Try `sample-table.pdf` first. Its data is made up, so you can share all of the
output.

## Step 2: fix the table (to do)

Once the pattern is known, a second macro will convert the PDF, undo the
known errors, and write a clean table to a sheet.
