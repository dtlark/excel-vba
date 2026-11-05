# excel-vba

Excel automation for processing a vendor's invoice files. VBA now, with a
Python version planned.

## What's here

- `vba/RawInvoices.bas`: turns the vendor's RAW Excel files into Invoice
  workbooks and adds each one to a Master List workbook. Setup and use are in
  [vba/README.md](vba/README.md).
- `vba/PdfDiagnose.bas`: shows how Word converts a PDF table (PDF experiment).
- `sample-table.pdf`: a made-up PDF table for testing PDF parsing.
