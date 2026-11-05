Option Explicit

' =====================================================================
'  RAW invoices: vendor file -> Invoice workbook -> Master List
'
'  For every Excel file in the RAW folder (in file name order):
'    1. Opens it. Its tab name (e.g. 123456) is the <Tab Name>.
'    2. Copies the header row and data rows (values and number formats only)
'       into a new one-tab workbook, header on row 3, leaving out rows that
'       contain "subtotal for" or "grand total", and empty rows. Numbers
'       stored as text become real numbers.
'    3. Makes it a table named Inv_<Tab Name> and adds three columns
'       with orange headers:
'         Check Number  = SHIPPED x PRICE/UNIT + TAX + FREIGHT
'         " " (blank)   = ERROR when AMOUNT isn't equal to Check Number
'         V             = Check Number - AMOUNT on ERROR rows
'    4. Puts =SUBTOTAL(9,Inv_<Tab Name>[AMOUNT]) on row 2 above AMOUNT.
'    5. Saves it as "Invoice <Tab Name> (<rest of the RAW file name>).xlsx"
'       in the Invoice folder, e.g. "Invoice 123456 (ABC).xlsx".
'    6. Copies the tab to the end of the Master List.
'    7. Adds the table's data rows (the vendor's columns, not the three added
'       ones) to the bottom of the Master List's Consumables table, on the
'       Consolidated Consumables tab: the tab name in the table's 4th column
'       and the vendor's columns from its 5th column on.
'    8. Moves the RAW file into the Done folder.
'
'  A file that can't be processed stays in RAW, and the message at the
'  end says why. The Master List is NOT saved, so you can check the new
'  tabs and Consumables rows first and then save it yourself.
'
'  The folders and the Master List file go on a sheet named "Settings"
'  in the workbook that holds this macro. The first run adds that sheet.
'
'  To run: Alt+F8, pick ProcessRawInvoices, click Run.
' =====================================================================

Private Const APP_TITLE As String = "RAW invoices"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const SET_RAW As String = "RAW folder"
Private Const SET_DONE As String = "Done folder"
Private Const SET_INVOICE As String = "Invoice folder"
Private Const SET_MASTER As String = "Master List file"
Private Const SET_FY As String = "Fiscal year"
' Anywhere a path on the Settings sheet says {FY}, the macro puts "FY" and the Fiscal
' year setting, e.g. {FY} -> FY26. Change the Fiscal year to 27 and every path follows.
Private Const FY_TAG As String = "{FY}"
Private Const FY_NOTE As String = "Just the number, like 26. Wherever a path says {FY}, " & _
                                  "the macro puts FY and this number, like FY26"

' Column headers the macro looks for in row 1 of the RAW file (capitals don't matter).
' Put other spellings after a |, like the vendor's "FREGHT" below.
Private Const H_SHIPPED As String = "SHIPPED"
Private Const H_PRICE As String = "PRICE/UNIT"
Private Const H_TAX As String = "TAX"
Private Const H_FREIGHT As String = "FREIGHT|FREGHT"
Private Const H_AMOUNT As String = "AMOUNT"
' ORDERED is only formatted, not used in a formula, so a file without it isn't skipped
Private Const H_ORDERED As String = "ORDERED"

' The Master List table that collects every invoice's data rows, and where they go in it
Private Const CONSUMABLES_TABLE As String = "Consumables"
Private Const CONSUMABLES_TAB As String = "Consolidated Consumables"   ' only used in messages
Private Const CONS_TAB_NAME_COLUMN As Long = 4        ' the tab name goes in the 4th column
Private Const CONS_FIRST_DATA_COLUMN As Long = 5      ' the vendor's columns start in the 5th

' Style for the Inv_ table: "White, Table Style Medium 1" in Excel's Table Design gallery.
' Other styles follow the same pattern, e.g. "TableStyleMedium2" or "TableStyleLight9".
Private Const TABLE_STYLE As String = "TableStyleMedium1"

' Number formats: Excel's "Number" with 2 decimals, and "Accounting" with a $ sign
Private Const FMT_NUMBER As String = "0.00"
Private Const FMT_ACCOUNTING As String = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

Private Const ORANGE As Long = 3305961        ' #E97132 = RGB(233, 113, 50)
Private Const ERROR_TEXT As Long = 393372     ' #9C0006 = RGB(156, 0, 6)
Private Const ERROR_FILL As Long = 13551615   ' #FFC7CE = RGB(255, 199, 206)

' Filled in from the Settings sheet when the macro starts
Private rawFolder As String
Private doneFolder As String
Private invoiceFolder As String
Private masterPath As String

Public Sub ProcessRawInvoices()
    Dim files() As String, nFiles As Long, i As Long
    Dim master As Workbook, consLo As ListObject
    Dim result As String, note As String, invoiceFile As String
    Dim fileList As String, doneList As String, skipList As String, summary As String
    Dim duplicates As String
    Dim nDone As Long, nSkipped As Long, oldCalc As Long, calcChanged As Boolean

    On Error GoTo Failed
    If Not LoadSettings() Then Exit Sub                     ' LoadSettings already said why

    If Not FolderExists(rawFolder) Then
        MsgBox "Can't find the RAW folder:" & vbLf & rawFolder & vbLf & vbLf & _
               "Check the Settings sheet.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Not FolderExists(invoiceFolder) Then
        MsgBox "Can't find the Invoice folder:" & vbLf & invoiceFolder & vbLf & vbLf & _
               "Check the Settings sheet.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    nFiles = ListRawFiles(files)
    If nFiles = 0 Then
        MsgBox "There are no Excel files in RAW:" & vbLf & rawFolder, vbInformation, APP_TITLE
        Exit Sub
    End If
    For i = 1 To nFiles
        If i <= 15 Then fileList = fileList & vbLf & "  " & files(i)
    Next i
    If nFiles > 15 Then fileList = fileList & vbLf & "  ...and " & (nFiles - 15) & " more"
    If MsgBox("Process these " & nFiles & " file(s) from RAW?" & vbLf & fileList, _
              vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub

    Set master = GetMasterWorkbook()
    If master Is Nothing Then Exit Sub                      ' GetMasterWorkbook already said why
    If master.ReadOnly Then
        MsgBox master.Name & " is open read-only, so the new tabs couldn't be saved." & vbLf & _
               "Someone else may have it open. Try again once they've closed it.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If master.ProtectStructure Then
        MsgBox master.Name & " has its structure protected, so tabs can't be added." & vbLf & _
               "Unprotect it (Review > Protect Workbook) and run again.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    Set consLo = FindTable(master, CONSUMABLES_TABLE)
    If consLo Is Nothing Then
        MsgBox master.Name & " has no table named " & CONSUMABLES_TABLE & " (it should be on the " & _
               CONSUMABLES_TAB & " tab), so nothing was processed.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    ' Any duplicate stops the whole run before anything is changed.
    duplicates = FindDuplicates(files, nFiles, master, consLo)
    If Len(duplicates) > 0 Then
        MsgBox "Stopped: found duplicates, so nothing was processed." & vbLf & duplicates & _
               vbLf & vbLf & "Take those files out of RAW, or remove the old copies " & _
               "(from the Master List, the Done folder or the Invoice folder), then run again.", _
               vbCritical, APP_TITLE
        Exit Sub
    End If
    If Not FolderExists(doneFolder) Then MkDir doneFolder

    Application.ScreenUpdating = False
    oldCalc = Application.Calculation                       ' manual while working, for speed
    Application.Calculation = xlCalculationManual
    calcChanged = True
    For i = 1 To nFiles
        Application.StatusBar = APP_TITLE & ": file " & i & " of " & nFiles & ": " & files(i)
        invoiceFile = ""
        note = ""
        result = ProcessOneFile(rawFolder & "\" & files(i), master, consLo, invoiceFile, note)
        If Len(result) = 0 Then
            nDone = nDone + 1
            doneList = doneList & vbLf & "  " & files(i) & "  ->  " & invoiceFile
            If Len(note) > 0 Then doneList = doneList & vbLf & "      Note: " & note
        Else
            nSkipped = nSkipped + 1
            skipList = skipList & vbLf & "  " & files(i) & ": " & result
        End If
    Next i
    Application.Calculation = oldCalc
    calcChanged = False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    master.Activate

    summary = "Processed " & nDone & " file(s)." & doneList
    If nSkipped > 0 Then
        summary = summary & vbLf & vbLf & "Left in RAW: " & nSkipped & " file(s)." & skipList
    End If
    If nDone > 0 Then
        summary = summary & vbLf & vbLf & "The Master List isn't saved yet. Check the new tabs and " & _
                  "the rows added to " & CONSUMABLES_TABLE & ", then save it."
    End If
    MsgBox summary, IIf(nSkipped > 0, vbExclamation, vbInformation), APP_TITLE
    Exit Sub

Failed:
    If calcChanged Then Application.Calculation = oldCalc
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Stopped with an error: " & Err.Description & vbLf & vbLf & _
           "Files finished before it stopped: " & nDone & doneList, vbExclamation, APP_TITLE
End Sub

' Processes one RAW file. Returns "" when it worked, otherwise why it was skipped.
' invoiceFile gets the saved file's name; note gets anything worth mentioning.
Private Function ProcessOneFile(ByVal rawPath As String, ByVal master As Workbook, _
                                ByVal consLo As ListObject, _
                                ByRef invoiceFile As String, ByRef note As String) As String
    Dim wbRaw As Workbook, wbInv As Workbook
    Dim wsRaw As Worksheet, ws As Worksheet
    Dim lo As ListObject, checkCol As ListColumn, errorCol As ListColumn, vCol As ListColumn
    Dim fc As FormatCondition
    Dim rawName As String, tabName As String, tableName As String
    Dim invoicePath As String, missing As String, notFound As String, errorCell As String
    Dim msg As String
    Dim lastRow As Long, lastCol As Long, nRows As Long, nKept As Long
    Dim colShipped As Long, colPrice As Long, colTax As Long, colFreight As Long, colAmount As Long
    Dim colOrdered As Long

    On Error GoTo Failed
    Application.DisplayAlerts = False                       ' no "replace this file?" prompts
    rawName = Mid$(rawPath, InStrRev(rawPath, "\") + 1)

    If IsWorkbookOpen(rawName) Then
        msg = "it's open in Excel. Close it and run again."
        GoTo Cleanup
    End If

    ' ----- 1. Open the RAW file -----
    Set wbRaw = Workbooks.Open(Filename:=rawPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set wsRaw = wbRaw.Worksheets(1)
    tabName = wsRaw.Name
    tableName = TableNameFor(tabName)

    If SheetExists(master, tabName) Then
        msg = "the Master List already has a tab named " & tabName & "."
        GoTo Cleanup
    End If
    If TableExists(master, tableName) Then
        msg = "the Master List already has a table named " & tableName & "."
        GoTo Cleanup
    End If

    ClearFilters wsRaw                                      ' rows hidden by a filter wouldn't copy
    lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    lastRow = LastUsedRow(wsRaw)
    If lastRow < 2 Then
        msg = "there's nothing under the header row."
        GoTo Cleanup
    End If
    If consLo.ListColumns.Count < CONS_FIRST_DATA_COLUMN - 1 + lastCol Then
        msg = "the " & CONSUMABLES_TABLE & " table has " & consLo.ListColumns.Count & _
              " columns, but this file's " & lastCol & " columns need " & _
              (CONS_FIRST_DATA_COLUMN - 1 + lastCol) & " (they start in its column " & _
              CONS_FIRST_DATA_COLUMN & ")."
        GoTo Cleanup
    End If

    ' ----- 2. Copy header + data into a new workbook, header on row 3 -----
    Set wbInv = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbInv.Worksheets(1)
    ws.Name = tabName
    ' Values and number formats only (dates stay dates). The vendor's colors, fonts and
    ' borders stay behind, so they don't cover up the table style.
    wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(lastRow, lastCol)).Copy
    ws.Range("A3").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    wbRaw.Close SaveChanges:=False
    Set wbRaw = Nothing

    nRows = lastRow - 1                                     ' data rows, before cleanup
    lastRow = 3 + nRows
    MakePlainData ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol))
    nKept = nRows - DeleteUnwantedRows(ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)))
    If nKept = 0 Then
        msg = "no data rows left after removing subtotal, grand total and empty rows."
        GoTo Cleanup
    End If
    lastRow = 3 + nKept
    ConvertTextNumbers ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol))   ' quietly, no note

    colShipped = FindHeader(ws, lastCol, H_SHIPPED, missing)
    colPrice = FindHeader(ws, lastCol, H_PRICE, missing)
    colTax = FindHeader(ws, lastCol, H_TAX, missing)
    colFreight = FindHeader(ws, lastCol, H_FREIGHT, missing)
    colAmount = FindHeader(ws, lastCol, H_AMOUNT, missing)
    If Len(missing) > 0 Then
        msg = "row 1 has no " & missing & " column. Row 1 has: " & HeaderList(ws, lastCol)
        GoTo Cleanup
    End If
    colOrdered = FindHeader(ws, lastCol, H_ORDERED, notFound)  ' optional, so no skip if missing

    ' ----- 3. Make the table and add the three new columns -----
    ws.Cells(3, lastCol + 1).Value = "Check Number"
    ws.Cells(3, lastCol + 2).Value = " "                    ' a space, so the header looks blank
    ws.Cells(3, lastCol + 3).Value = "V"
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, XlListObjectHasHeaders:=xlYes, _
        Source:=ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol + 3)))
    lo.Name = tableName
    lo.TableStyle = TABLE_STYLE
    Set checkCol = lo.ListColumns(lastCol + 1)
    Set errorCol = lo.ListColumns(lastCol + 2)
    Set vCol = lo.ListColumns(lastCol + 3)

    ' Check Number = (SHIPPED x PRICE/UNIT) + TAX + FREIGHT
    checkCol.DataBodyRange.Formula = "=" & ThisRow(lo, colShipped) & "*" & ThisRow(lo, colPrice) & _
        "+" & ThisRow(lo, colTax) & "+" & ThisRow(lo, colFreight)

    ' ERROR when AMOUNT and Check Number differ. Compared to 6 decimal places so computer
    ' rounding noise (like 0.00000000000001) isn't flagged; any real difference is,
    ' even a fraction of a cent.
    errorCol.DataBodyRange.Formula = "=IF(ROUND(" & ThisRow(lo, colAmount) & "-" & _
        ThisRow(lo, checkCol.Index) & ",6)=0,"""",""ERROR"")"

    ' V = Check Number - AMOUNT on ERROR rows. The blank column's name is a space,
    ' so V points at its cell (like P4) instead of using its name.
    errorCell = errorCol.DataBodyRange.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    vCol.DataBodyRange.Formula = "=IF(" & errorCell & "=""ERROR""," & _
        ThisRow(lo, checkCol.Index) & "-" & ThisRow(lo, colAmount) & ","""")"

    ' ----- 4. Row 2: SUBTOTAL above AMOUNT -----
    ws.Cells(2, colAmount).Formula = "=SUBTOTAL(9," & lo.Name & "[" & _
        EscapeColumnName(lo.ListColumns(colAmount).Name) & "])"

    ' Number formats: Number for the quantities, Accounting for the money
    lo.ListColumns(colShipped).DataBodyRange.NumberFormat = FMT_NUMBER
    If colOrdered > 0 Then
        lo.ListColumns(colOrdered).DataBodyRange.NumberFormat = FMT_NUMBER
    Else
        AddNote note, "no " & notFound & " column, so it wasn't formatted. Row 1 has: " & _
                      HeaderList(ws, lastCol)
    End If
    lo.ListColumns(colPrice).DataBodyRange.NumberFormat = FMT_ACCOUNTING
    lo.ListColumns(colTax).DataBodyRange.NumberFormat = FMT_ACCOUNTING
    lo.ListColumns(colFreight).DataBodyRange.NumberFormat = FMT_ACCOUNTING
    lo.ListColumns(colAmount).DataBodyRange.NumberFormat = FMT_ACCOUNTING
    checkCol.DataBodyRange.NumberFormat = FMT_ACCOUNTING
    vCol.DataBodyRange.NumberFormat = FMT_ACCOUNTING
    ws.Cells(2, colAmount).NumberFormat = FMT_ACCOUNTING     ' the SUBTOTAL above the table

    ' Orange headers, red ERROR cells
    ws.Range(ws.Cells(3, lastCol + 1), ws.Cells(3, lastCol + 3)).Interior.Color = ORANGE
    Set fc = errorCol.DataBodyRange.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""ERROR""")
    fc.Font.Color = ERROR_TEXT
    fc.Interior.Color = ERROR_FILL
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol + 3)).Columns.AutoFit
    If errorCol.Range.ColumnWidth < 8 Then errorCol.Range.ColumnWidth = 8
    If vCol.Range.ColumnWidth < 8 Then vCol.Range.ColumnWidth = 8

    ' ----- 5. Save the Invoice workbook (never over an existing one) -----
    invoiceFile = InvoiceFileName(rawName, tabName, note)
    invoicePath = invoiceFolder & "\" & invoiceFile
    If FileExists(invoicePath) Then                         ' the duplicate check normally catches this
        msg = "the Invoice folder already has " & invoiceFile & ", so nothing was changed for this file."
        GoTo Cleanup
    End If
    wbInv.SaveAs Filename:=invoicePath, FileFormat:=xlOpenXMLWorkbook

    ' ----- 6. Copy the tab to the end of the Master List, with an orange tab -----
    ws.Copy After:=master.Sheets(master.Sheets.Count)
    master.Worksheets(tabName).Tab.Color = ORANGE

    ' ----- 7. Add the data rows to the bottom of the Consumables table -----
    AppendToConsumables consLo, lo, lastCol, tabName
    wbInv.Close SaveChanges:=False
    Set wbInv = Nothing

    ' ----- 8. Move the RAW file to Done -----
    MoveToDone rawPath, rawName, note

Cleanup:
    On Error Resume Next
    If Not wbRaw Is Nothing Then wbRaw.Close SaveChanges:=False
    If Not wbInv Is Nothing Then wbInv.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ProcessOneFile = msg
    Exit Function

Failed:
    msg = "error: " & Err.Description
    Resume Cleanup
End Function

' Reads the Settings sheet in this workbook, adding the sheet if it isn't there.
' Returns False (after saying what to do) when something needs filling in.
Private Function LoadSettings() As Boolean
    Dim ws As Worksheet, missing As String, fiscalYear As String, usesFY As Boolean, r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SETTINGS_SHEET
        ws.Range("A1:C1").Value = Array("Setting", "Value", "What to put here")
        ws.Range("A2:C2").Value = Array(SET_FY, "", FY_NOTE)
        ws.Range("A3:C3").Value = Array(SET_RAW, "", "Folder the vendor's Excel files are saved into")
        ws.Range("A4:C4").Value = Array(SET_DONE, "", "Where processed RAW files go. Blank = a Done folder inside the RAW folder")
        ws.Range("A5:C5").Value = Array(SET_INVOICE, "", "Folder the Invoice <Tab Name> files are saved in")
        ws.Range("A6:C6").Value = Array(SET_MASTER, "", "Full path of the Master List workbook (the .xlsx can be left off)")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:C").AutoFit
        ws.Columns("B").ColumnWidth = 70
        ws.Activate
        ws.Range("B2").Select
        MsgBox "I added a Settings sheet to this workbook." & vbLf & vbLf & _
               "Fill in column B (the fiscal year, the folders and the Master List file), " & _
               "save this workbook, and run the macro again.", vbInformation, APP_TITLE
        Exit Function
    End If

    ' Settings sheets made before the Fiscal year setting existed get its row added at the end.
    If ws.Columns(1).Find(What:=SET_FY, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).Value = SET_FY
        ws.Cells(r, 3).Value = FY_NOTE
    End If

    ' "26" and "FY26" both mean fiscal year 26
    fiscalYear = ReadSetting(ws, SET_FY)
    If UCase$(Left$(fiscalYear, 2)) = "FY" Then fiscalYear = Trim$(Mid$(fiscalYear, 3))

    rawFolder = WithFiscalYear(ReadSetting(ws, SET_RAW), fiscalYear, usesFY)
    doneFolder = WithFiscalYear(ReadSetting(ws, SET_DONE), fiscalYear, usesFY)
    invoiceFolder = WithFiscalYear(ReadSetting(ws, SET_INVOICE), fiscalYear, usesFY)
    masterPath = WithFiscalYear(ReadSetting(ws, SET_MASTER), fiscalYear, usesFY)
    If Len(doneFolder) = 0 And Len(rawFolder) > 0 Then doneFolder = rawFolder & "\Done"

    If usesFY And Len(fiscalYear) = 0 Then
        missing = missing & vbLf & "  " & SET_FY & " (a path uses " & FY_TAG & ")"
    End If
    If Len(rawFolder) = 0 Then missing = missing & vbLf & "  " & SET_RAW
    If Len(invoiceFolder) = 0 Then missing = missing & vbLf & "  " & SET_INVOICE
    If InStr(masterPath, "\") = 0 Then missing = missing & vbLf & "  " & SET_MASTER & " (the full path)"
    If Len(missing) > 0 Then
        ws.Activate
        MsgBox "Fill these in on the Settings sheet, in column B, then run again:" & missing, _
               vbExclamation, APP_TITLE
        Exit Function
    End If
    LoadSettings = True
End Function

' The path with every {FY} replaced by "FY" and the fiscal year, e.g. FY26.
' Sets usesFY to True when the path had a {FY} in it.
Private Function WithFiscalYear(ByVal path As String, ByVal fiscalYear As String, _
                                ByRef usesFY As Boolean) As String
    If InStr(1, path, FY_TAG, vbTextCompare) > 0 Then
        usesFY = True
        path = Replace(path, FY_TAG, "FY" & fiscalYear, 1, -1, vbTextCompare)
    End If
    WithFiscalYear = path
End Function

' The value in column B next to a label in column A, without quotes or a final \
Private Function ReadSetting(ByVal ws As Worksheet, ByVal label As String) As String
    Dim found As Range, v As Variant, s As String
    Set found = ws.Columns(1).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If found Is Nothing Then Exit Function
    v = found.Offset(0, 1).Value
    If IsError(v) Then Exit Function
    s = Trim$(Replace(CStr(v), """", ""))                   ' "Copy as path" adds quotes
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    ReadSetting = s
End Function

' Fills files(1..n) with the Excel files in the RAW folder, sorted by name. Returns n.
Private Function ListRawFiles(ByRef files() As String) As Long
    Dim f As String, n As Long, i As Long, j As Long, tmp As String
    ReDim files(1 To 1)
    f = Dir(rawFolder & "\*.xls*")
    Do While Len(f) > 0
        ' Skip Excel's temporary "~$" lock files, and this macro's own workbook
        If Left$(f, 2) <> "~$" And _
           StrComp(rawFolder & "\" & f, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            n = n + 1
            ReDim Preserve files(1 To n)
            files(n) = f
        End If
        f = Dir()
    Loop
    For i = 2 To n                                          ' insertion sort by name
        tmp = files(i)
        j = i - 1
        Do While j >= 1
            If StrComp(files(j), tmp, vbTextCompare) <= 0 Then Exit Do
            files(j + 1) = files(j)
            j = j - 1
        Loop
        files(j + 1) = tmp
    Next i
    ListRawFiles = n
End Function

' Uses the Master List if it's already open, otherwise opens it.
Private Function GetMasterWorkbook() As Workbook
    Dim wb As Workbook, ext As Variant
    Dim basePath As String, masterName As String, masterFolder As String, dotPos As Long

    ' The setting may or may not end in .xlsx / .xlsm
    dotPos = InStrRev(masterPath, ".")
    If dotPos > InStrRev(masterPath, "\") And LCase$(Mid$(masterPath, dotPos)) Like ".xls*" Then
        basePath = Left$(masterPath, dotPos - 1)
    Else
        basePath = masterPath
    End If
    masterFolder = Left$(basePath, InStrRev(basePath, "\") - 1)
    masterName = Mid$(basePath, InStrRev(basePath, "\") + 1)

    For Each wb In Application.Workbooks
        If StrComp(NameWithoutExtension(wb.Name), masterName, vbTextCompare) = 0 Then
            If StrComp(wb.Path, masterFolder, vbTextCompare) <> 0 Then
                If MsgBox("The open " & wb.Name & " is in" & vbLf & wb.Path & vbLf & vbLf & _
                          "instead of" & vbLf & masterFolder & vbLf & vbLf & "Use it anyway?", _
                          vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Function
            End If
            Set GetMasterWorkbook = wb
            Exit Function
        End If
    Next wb

    If FileExists(masterPath) Then
        Set GetMasterWorkbook = Workbooks.Open(Filename:=masterPath)
        Exit Function
    End If
    For Each ext In Array(".xlsx", ".xlsm", ".xlsb", ".xls")
        If FileExists(basePath & ext) Then
            Set GetMasterWorkbook = Workbooks.Open(Filename:=basePath & ext)
            Exit Function
        End If
    Next ext

    MsgBox "Can't find the Master List:" & vbLf & masterPath & vbLf & vbLf & _
           "Check the Settings sheet.", vbExclamation, APP_TITLE
End Function

' Shows any rows hidden by a filter, because copying skips filtered-out rows.
Private Sub ClearFilters(ByVal ws As Worksheet)
    Dim tbl As ListObject
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    For Each tbl In ws.ListObjects
        tbl.AutoFilter.ShowAllData
    Next tbl
    On Error GoTo 0
End Sub

' Turns the copied block into plain values: no table, no merged cells, no formulas.
Private Sub MakePlainData(ByVal rng As Range)
    Dim formulaCells As Range, area As Range
    Do While rng.Worksheet.ListObjects.Count > 0
        rng.Worksheet.ListObjects(1).Unlist
    Loop
    rng.UnMerge
    On Error Resume Next                                    ' SpecialCells errors when there are none
    Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not formulaCells Is Nothing Then
        For Each area In formulaCells.Areas
            area.Value = area.Value
        Next area
    End If
End Sub

' Deletes rows containing "subtotal for" or "grand total" (any capitals), and empty
' rows. Returns how many rows it deleted.
Private Function DeleteUnwantedRows(ByVal rng As Range) As Long
    Dim v As Variant, r As Long, c As Long, s As String
    Dim keep As Boolean, hasValue As Boolean, toDelete As Range
    v = ValuesOf(rng)
    For r = 1 To UBound(v, 1)
        keep = True
        hasValue = False
        For c = 1 To UBound(v, 2)
            If IsError(v(r, c)) Then
                hasValue = True
            Else
                s = LCase$(Trim$(CStr(v(r, c))))
                If Len(s) > 0 Then hasValue = True
                If InStr(s, "subtotal for") > 0 Or InStr(s, "grand total") > 0 Then keep = False
            End If
        Next c
        If Not hasValue Then keep = False
        If Not keep Then
            If toDelete Is Nothing Then
                Set toDelete = rng.Rows(r).EntireRow
            Else
                Set toDelete = Union(toDelete, rng.Rows(r).EntireRow)
            End If
            DeleteUnwantedRows = DeleteUnwantedRows + 1
        End If
    Next r
    If Not toDelete Is Nothing Then toDelete.Delete
End Function

' Turns numbers stored as text (like "12.50", "1,234" or "$ (5.00)") into real numbers
' and returns how many cells it changed. Only those cells are rewritten, so any other
' text is left exactly as it was.
Private Function ConvertTextNumbers(ByVal rng As Range) As Long
    Dim v As Variant, r As Long, c As Long, n As Double
    v = ValuesOf(rng)
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If VarType(v(r, c)) = vbString Then
                If TextToNumber(CStr(v(r, c)), n) Then
                    With rng.Cells(r, c)
                        If .NumberFormat = "@" Then .NumberFormat = "General"   ' Text format keeps text
                        .Value = n
                    End With
                    ConvertTextNumbers = ConvertTextNumbers + 1
                End If
            End If
        Next c
    Next r
End Function

' True when s is a number written as text, and puts the number in result. Handles
' spaces (also tabs, line breaks and invisible spaces), a $ sign, commas between
' thousands, and negatives written as -12.50, 12.50- or (12.50). Leading zeros are
' dropped: "00123" becomes 123. Stays text: more than 15 digits (too long for Excel
' to hold exactly), dates, percents and anything else.
Private Function TextToNumber(ByVal s As String, ByRef result As Double) As Boolean
    Dim negative As Boolean, intPart As String, fracPart As String, dotPos As Long
    Dim groups As Variant, g As Long

    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")                           ' non-breaking space
    s = Replace(s, ChrW(8203), "")                          ' zero-width space
    s = Replace(s, ChrW(65279), "")                         ' byte-order mark
    s = Trim$(s)
    If Len(s) - Len(Replace(s, "$", "")) > 1 Then Exit Function
    s = Trim$(Replace(s, "$", ""))
    If Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then           ' (12.50)
        negative = True
        s = Trim$(Mid$(s, 2, Len(s) - 2))
    ElseIf Len(s) >= 2 And Right$(s, 1) = "-" Then                             ' 12.50-
        negative = True
        s = Trim$(Left$(s, Len(s) - 1))
    ElseIf Len(s) >= 2 And (Left$(s, 1) = "-" Or Left$(s, 1) = "+") Then       ' -12.50
        negative = (Left$(s, 1) = "-")
        s = Trim$(Mid$(s, 2))
    End If
    If Len(s) = 0 Then Exit Function

    ' What's left may only be digits, commas between groups of 3, and one dot.
    dotPos = InStr(s, ".")
    If dotPos > 0 Then
        intPart = Left$(s, dotPos - 1)
        fracPart = Mid$(s, dotPos + 1)
        If Len(fracPart) > 0 And Not IsAllDigits(fracPart) Then Exit Function
    Else
        intPart = s
    End If
    If InStr(intPart, ",") > 0 Then
        groups = Split(intPart, ",")
        If Len(groups(0)) > 3 Or Not IsAllDigits(CStr(groups(0))) Then Exit Function
        For g = 1 To UBound(groups)
            If Len(groups(g)) <> 3 Or Not IsAllDigits(CStr(groups(g))) Then Exit Function
        Next g
        intPart = Replace(intPart, ",", "")
    ElseIf Len(intPart) > 0 Then
        If Not IsAllDigits(intPart) Then Exit Function
    End If
    If Len(intPart) + Len(fracPart) = 0 Then Exit Function
    If Len(intPart) + Len(fracPart) > 15 Then Exit Function                 ' too long for Excel

    result = Val(intPart & "." & fracPart)                  ' Val always reads "." as the decimal point
    If negative Then result = -result
    TextToNumber = True
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    IsAllDigits = Len(s) > 0 And Not (s Like "*[!0-9]*")
End Function

' A range's values as a 2-D array, even when it's a single cell.
Private Function ValuesOf(ByVal rng As Range) As Variant
    Dim v As Variant
    If rng.Cells.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = rng.Value
        ValuesOf = v
    Else
        ValuesOf = rng.Value
    End If
End Function

' Column number of a header on row 3. Capitals, extra spaces, non-breaking spaces and
' line breaks inside the header don't matter.
' "spellings" lists every accepted spelling, separated by |, e.g. "FREIGHT|FREGHT".
' Returns 0 and adds the first spelling to "missing" when none are there.
Private Function FindHeader(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal spellings As String, _
                            ByRef missing As String) As Long
    Dim c As Long, v As Variant, spelling As Variant
    For Each spelling In Split(spellings, "|")
        For c = 1 To lastCol
            v = ws.Cells(3, c).Value
            If Not IsError(v) Then
                If CleanHeader(CStr(v)) = CleanHeader(CStr(spelling)) Then
                    FindHeader = c
                    Exit Function
                End If
            End If
        Next c
    Next spelling
    If Len(missing) > 0 Then missing = missing & ", "
    missing = missing & Split(spellings, "|")(0)
End Function

' Header text in capitals, with line breaks, tabs and non-breaking spaces turned into
' plain spaces, and no doubled, leading or trailing spaces.
Private Function CleanHeader(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")                           ' non-breaking space
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanHeader = UCase$(Trim$(s))
End Function

' The headers on row 3 for messages, like "ITEM, DESCRIPTION, SHIPPED, ..."
Private Function HeaderList(ByVal ws As Worksheet, ByVal lastCol As Long) As String
    Dim c As Long, v As Variant
    For c = 1 To lastCol
        v = ws.Cells(3, c).Value
        If c > 1 Then HeaderList = HeaderList & ", "
        If IsError(v) Then
            HeaderList = HeaderList & "(error)"
        Else
            HeaderList = HeaderList & CleanHeader(CStr(v))
        End If
    Next c
End Function

' A this-row reference to a table column, like [@[PRICE/UNIT]]
Private Function ThisRow(ByVal lo As ListObject, ByVal colIndex As Long) As String
    ThisRow = "[@[" & EscapeColumnName(lo.ListColumns(colIndex).Name) & "]]"
End Function

' Table formulas need a ' in front of [ ] # and ' when they appear in a column name.
Private Function EscapeColumnName(ByVal colName As String) As String
    colName = Replace(colName, "'", "''")
    colName = Replace(colName, "[", "'[")
    colName = Replace(colName, "]", "']")
    colName = Replace(colName, "#", "'#")
    EscapeColumnName = colName
End Function

' "123456" -> "Inv_123456". Table names allow only letters, numbers, _ and .
Private Function TableNameFor(ByVal tabName As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(tabName)
        ch = Mid$(tabName, i, 1)
        If ch Like "[A-Za-z0-9_.]" Then out = out & ch Else out = out & "_"
    Next i
    TableNameFor = "Inv_" & out
End Function

' The part of the RAW file name after the tab name:
' "123456 Oct.xlsx" -> "Oct", "123456 - Oct.xlsx" -> "Oct", "123456.xlsx" -> ""
Private Function RawPostfix(ByVal rawName As String, ByVal tabName As String, _
                            ByRef note As String) As String
    Dim baseName As String, rest As String
    baseName = NameWithoutExtension(rawName)
    If StrComp(Left$(baseName, Len(tabName)), tabName, vbTextCompare) = 0 Then
        rest = Mid$(baseName, Len(tabName) + 1)
    Else
        rest = baseName
        AddNote note, "the file name doesn't start with its tab name (" & tabName & _
                      "), so the whole file name was used"
    End If
    Do While Len(rest) > 0
        If InStr(" _-", Left$(rest, 1)) = 0 Then Exit Do
        rest = Mid$(rest, 2)
    Loop
    RawPostfix = Trim$(rest)
End Function

' Moves the RAW file into the Done folder (adds the date and time if that name is taken).
Private Sub MoveToDone(ByVal rawPath As String, ByVal rawName As String, ByRef note As String)
    Dim target As String
    target = doneFolder & "\" & rawName
    If FileExists(target) Then
        target = doneFolder & "\" & NameWithoutExtension(rawName) & " " & _
                 Format$(Now, "yyyy-mm-dd hhnnss") & Mid$(rawName, InStrRev(rawName, "."))
    End If
    On Error Resume Next
    Name rawPath As target
    If Err.Number <> 0 Then
        AddNote note, "couldn't move the RAW file to Done (" & Err.Description & "), so move it yourself"
    End If
    On Error GoTo 0
End Sub

Private Sub AddNote(ByRef note As String, ByVal addition As String)
    If Len(note) > 0 Then note = note & "; "
    note = note & addition
End Sub

Private Function NameWithoutExtension(ByVal fileName As String) As String
    If InStrRev(fileName, ".") > 0 Then
        NameWithoutExtension = Left$(fileName, InStrRev(fileName, ".") - 1)
    Else
        NameWithoutExtension = fileName
    End If
End Function

Private Function IsWorkbookOpen(ByVal fileName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sh
End Function

' Adds the invoice table's data rows (only its first nCols columns, the vendor's own)
' to the bottom of the Consumables table: the tab name in its 4th column and the
' vendor's columns from its 5th column on.
Private Sub AppendToConsumables(ByVal consLo As ListObject, ByVal invLo As ListObject, _
                                ByVal nCols As Long, ByVal tabName As String)
    Dim nAdd As Long, firstNew As Long, reuseBlankRow As Boolean

    nAdd = invLo.ListRows.Count
    ' A table with no data yet usually has one empty row: fill that one first.
    If consLo.ListRows.Count = 1 Then
        reuseBlankRow = (Application.WorksheetFunction.CountA( _
            consLo.DataBodyRange.Cells(1, CONS_TAB_NAME_COLUMN).Resize(1, nCols + 1)) = 0)
    End If
    If reuseBlankRow Then
        firstNew = 1
    Else
        firstNew = consLo.ListRows.Count + 1
    End If
    Do While consLo.ListRows.Count < firstNew + nAdd - 1
        consLo.ListRows.Add                                 ' formula columns fill in by themselves
    Loop

    invLo.DataBodyRange.Resize(nAdd, nCols).Copy
    consLo.DataBodyRange.Cells(firstNew, CONS_FIRST_DATA_COLUMN).PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    consLo.DataBodyRange.Cells(firstNew, CONS_TAB_NAME_COLUMN).Resize(nAdd, 1).Value = tabName

    ' If any of those Consumables columns are formatted as Text, the values above land as
    ' text (including the tab name). Turn them into real numbers.
    ConvertTextNumbers consLo.DataBodyRange.Cells(firstNew, CONS_TAB_NAME_COLUMN).Resize(nAdd, nCols + 1)
End Sub

' Checks every RAW file before anything is changed. Returns one line per duplicate
' file, listing every reason ("" when there are none). A file is a duplicate when:
'   - the Done folder already has a file with the same name,
'   - the Master List already has its tab or its Inv_ table,
'   - the Consumables table already has rows for its tab name,
'   - the Invoice folder already has the Invoice file it would make, or
'   - another RAW file in this run has the same tab name.
Private Function FindDuplicates(ByRef files() As String, ByVal nFiles As Long, _
                                ByVal master As Workbook, ByVal consLo As ListObject) As String
    Dim i As Long, tabName As String, reasons As String, seen As String
    Dim invoiceFile As String, ignored As String

    Application.ScreenUpdating = False
    For i = 1 To nFiles
        Application.StatusBar = APP_TITLE & ": checking for duplicates, file " & i & " of " & nFiles
        reasons = ""
        If FileExists(doneFolder & "\" & files(i)) Then
            AddNote reasons, "the Done folder already has a file with this name"
        End If
        tabName = ""
        If Not IsWorkbookOpen(files(i)) Then tabName = RawTabName(rawFolder & "\" & files(i))
        If Len(tabName) > 0 Then                            ' (open files get skipped later anyway)
            If SheetExists(master, tabName) Then
                AddNote reasons, "the Master List already has a tab named " & tabName
            ElseIf TableExists(master, TableNameFor(tabName)) Then
                AddNote reasons, "the Master List already has a table named " & TableNameFor(tabName)
            End If
            If ConsumablesHas(consLo, tabName) Then
                AddNote reasons, "the " & CONSUMABLES_TABLE & " table already has rows for " & tabName
            End If
            invoiceFile = InvoiceFileName(files(i), tabName, ignored)
            If FileExists(invoiceFolder & "\" & invoiceFile) Then
                AddNote reasons, "the Invoice folder already has " & invoiceFile
            End If
            If InStr(1, seen, "|" & tabName & "|", vbTextCompare) > 0 Then
                AddNote reasons, "another file in RAW is also tab " & tabName
            End If
            seen = seen & "|" & tabName & "|"
        End If
        If Len(reasons) > 0 Then FindDuplicates = FindDuplicates & vbLf & "  " & files(i) & ": " & reasons
    Next i
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Function

' "Invoice 123456 (ABC).xlsx" from the RAW file name and its tab name. No parentheses
' when nothing follows the tab name in the RAW file name.
Private Function InvoiceFileName(ByVal rawName As String, ByVal tabName As String, _
                                 ByRef note As String) As String
    Dim postfixText As String
    postfixText = RawPostfix(rawName, tabName, note)
    InvoiceFileName = "Invoice " & tabName
    If Len(postfixText) > 0 Then
        If Not (Left$(postfixText, 1) = "(" And Right$(postfixText, 1) = ")") Then
            postfixText = "(" & postfixText & ")"           ' don't double up "(ABC)"
        End If
        InvoiceFileName = InvoiceFileName & " " & postfixText
    End If
    InvoiceFileName = InvoiceFileName & ".xlsx"
End Function

' The name of a RAW file's tab, or "" if the file can't be opened.
Private Function RawTabName(ByVal path As String) As String
    Dim wb As Workbook
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    RawTabName = wb.Worksheets(1).Name
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Exit Function
Failed:
    Resume Cleanup
End Function

' True when the Consumables table's tab-name column already holds this tab name
' (as a number or as text).
Private Function ConsumablesHas(ByVal consLo As ListObject, ByVal tabName As String) As Boolean
    If consLo.DataBodyRange Is Nothing Then Exit Function
    ConsumablesHas = Application.WorksheetFunction.CountIf( _
        consLo.DataBodyRange.Columns(CONS_TAB_NAME_COLUMN), tabName) > 0
End Function

' The table with this name anywhere in the workbook, or Nothing.
Private Function FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim sh As Worksheet, tbl As ListObject
    For Each sh In wb.Worksheets
        For Each tbl In sh.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = tbl
                Exit Function
            End If
        Next tbl
    Next sh
End Function

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    TableExists = Not FindTable(wb, tableName) Is Nothing
End Function

' Last row with anything in it (0 if the sheet is empty). Includes hidden rows.
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LastUsedRow = f.Row
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    Dim attr As Long
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    attr = GetAttr(path)
    If Err.Number = 0 Then FolderExists = (attr And vbDirectory) = vbDirectory
End Function

Private Function FileExists(ByVal path As String) As Boolean
    Dim attr As Long
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    attr = GetAttr(path)
    If Err.Number = 0 Then FileExists = (attr And vbDirectory) = 0
End Function
