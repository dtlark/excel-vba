# excel-vba

Excel automation for processing a vendor's invoice files, in two versions that
do the same job: a VBA macro and a Python script.

## What's here

- `vba/RawInvoices.bas`: turns the vendor's RAW Excel files into Invoice
  workbooks and adds each one to a Master List workbook. Setup and use are in
  [vba/README.md](vba/README.md).
- `python/raw_invoices.py`: the same job in Python, without Excel. It reads and
  writes the `.xlsx` files directly. Setup, use and how it differs from the
  macro are in [python/README.md](python/README.md).
- `vba/PdfDiagnose.bas`: shows how Word converts a PDF table (PDF experiment).
- `sample-table.pdf`: a made-up PDF table for testing PDF parsing.
