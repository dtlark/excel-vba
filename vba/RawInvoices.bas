Option Explicit

' =====================================================================
'  RAW invoices: vendor file -> Invoice workbook -> Master List
'
'  For every Excel file in the RAW folder (in file name order):
'    1. Opens it. Its tab name (e.g. 123456) is the <Tab Name>.
'    2. Copies the header row and data rows into a new one-tab workbook,
'       header on row 3, leaving out rows that contain "subtotal for" or
'       "grand total", and empty rows.
'    3. Makes it a table named Inv_<Tab Name> and adds three columns
'       with orange headers:
'         Check Number  = SHIPPED x PRICE/UNIT + TAX + FREIGHT
'         " " (blank)   = ERROR when AMOUNT isn't equal to Check Number
'         V             = Check Number - AMOUNT on ERROR rows
'    4. Puts =SUBTOTAL(9,Inv_<Tab Name>[AMOUNT]) on row 2 above AMOUNT.
'    5. Saves it as "Invoice <Tab Name> <rest of the RAW file name>.xlsx"
'       in the Invoice folder.
'    6. Copies the tab to the end of the Master List.
'    7. Moves the RAW file into the Done folder.
'
'  A file that can't be processed stays in RAW, and the message at the
'  end says why. The Master List is NOT saved, so you can check the new
'  tabs first and then save it yourself.
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
    Dim master As Workbook
    Dim result As String, note As String, invoiceFile As String
    Dim fileList As String, doneList As String, skipList As String, summary As String
    Dim nDone As Long, nSkipped As Long

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
    If Not FolderExists(doneFolder) Then MkDir doneFolder

    Application.ScreenUpdating = False
    For i = 1 To nFiles
        Application.StatusBar = APP_TITLE & ": file " & i & " of " & nFiles & ": " & files(i)
        invoiceFile = ""
        note = ""
        result = ProcessOneFile(rawFolder & "\" & files(i), master, invoiceFile, note)
        If Len(result) = 0 Then
            nDone = nDone + 1
            doneList = doneList & vbLf & "  " & files(i) & "  ->  " & invoiceFile
            If Len(note) > 0 Then doneList = doneList & vbLf & "      Note: " & note
        Else
            nSkipped = nSkipped + 1
            skipList = skipList & vbLf & "  " & files(i) & ": " & result
        End If
    Next i
    Application.StatusBar = False
    Application.ScreenUpdating = True
    master.Activate

    summary = "Processed " & nDone & " file(s)." & doneList
    If nSkipped > 0 Then
        summary = summary & vbLf & vbLf & "Left in RAW: " & nSkipped & " file(s)." & skipList
    End If
    If nDone > 0 Then
        summary = summary & vbLf & vbLf & "The Master List isn't saved yet. Check the new tabs, then save it."
    End If
    MsgBox summary, IIf(nSkipped > 0, vbExclamation, vbInformation), APP_TITLE
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Stopped with an error: " & Err.Description & vbLf & vbLf & _
           "Files finished before it stopped: " & nDone & doneList, vbExclamation, APP_TITLE
End Sub

' Processes one RAW file. Returns "" when it worked, otherwise why it was skipped.
' invoiceFile gets the saved file's name; note gets anything worth mentioning.
Private Function ProcessOneFile(ByVal rawPath As String, ByVal master As Workbook, _
                                ByRef invoiceFile As String, ByRef note As String) As String
    Dim wbRaw As Workbook, wbInv As Workbook
    Dim wsRaw As Worksheet, ws As Worksheet
    Dim lo As ListObject, checkCol As ListColumn, errorCol As ListColumn, vCol As ListColumn
    Dim fc As FormatCondition
    Dim rawName As String, tabName As String, tableName As String, postfixText As String
    Dim invoicePath As String, missing As String, amountFormat As String, errorCell As String
    Dim msg As String
    Dim lastRow As Long, lastCol As Long, nRows As Long, nKept As Long
    Dim colShipped As Long, colPrice As Long, colTax As Long, colFreight As Long, colAmount As Long

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

    ' ----- 2. Copy header + data into a new workbook, header on row 3 -----
    Set wbInv = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbInv.Worksheets(1)
    ws.Name = tabName
    wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(lastRow, lastCol)).Copy Destination:=ws.Range("A3")
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

    colShipped = FindHeader(ws, lastCol, "SHIPPED", missing)
    colPrice = FindHeader(ws, lastCol, "PRICE/UNIT", missing)
    colTax = FindHeader(ws, lastCol, "TAX", missing)
    colFreight = FindHeader(ws, lastCol, "FREIGHT", missing)
    colAmount = FindHeader(ws, lastCol, "AMOUNT", missing)
    If Len(missing) > 0 Then
        msg = "row 1 has no " & missing & " column."
        GoTo Cleanup
    End If

    ' ----- 3. Make the table and add the three new columns -----
    ws.Cells(3, lastCol + 1).Value = "Check Number"
    ws.Cells(3, lastCol + 2).Value = " "                    ' a space, so the header looks blank
    ws.Cells(3, lastCol + 3).Value = "V"
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, XlListObjectHasHeaders:=xlYes, _
        Source:=ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol + 3)))
    lo.Name = tableName
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

    ' Formatting: money format like AMOUNT, orange headers, red ERROR cells
    amountFormat = lo.ListColumns(colAmount).DataBodyRange.Cells(1, 1).NumberFormat
    ws.Cells(2, colAmount).NumberFormat = amountFormat
    checkCol.DataBodyRange.NumberFormat = amountFormat
    vCol.DataBodyRange.NumberFormat = amountFormat
    ws.Range(ws.Cells(3, lastCol + 1), ws.Cells(3, lastCol + 3)).Interior.Color = ORANGE
    Set fc = errorCol.DataBodyRange.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""ERROR""")
    fc.Font.Color = ERROR_TEXT
    fc.Interior.Color = ERROR_FILL
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol + 3)).Columns.AutoFit
    If errorCol.Range.ColumnWidth < 8 Then errorCol.Range.ColumnWidth = 8
    If vCol.Range.ColumnWidth < 8 Then vCol.Range.ColumnWidth = 8

    ' ----- 5. Save the Invoice workbook -----
    postfixText = RawPostfix(rawName, tabName, note)
    invoiceFile = "Invoice " & tabName
    If Len(postfixText) > 0 Then invoiceFile = invoiceFile & " " & postfixText
    invoiceFile = invoiceFile & ".xlsx"
    invoicePath = invoiceFolder & "\" & invoiceFile
    If FileExists(invoicePath) Then AddNote note, "replaced the existing " & invoiceFile
    wbInv.SaveAs Filename:=invoicePath, FileFormat:=xlOpenXMLWorkbook

    ' ----- 6. Copy the tab to the end of the Master List -----
    ws.Copy After:=master.Sheets(master.Sheets.Count)
    wbInv.Close SaveChanges:=False
    Set wbInv = Nothing

    ' ----- 7. Move the RAW file to Done -----
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
    Dim ws As Worksheet, missing As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SETTINGS_SHEET
        ws.Range("A1:C1").Value = Array("Setting", "Value", "What to put here")
        ws.Range("A2:C2").Value = Array(SET_RAW, "", "Folder the vendor's Excel files are saved into")
        ws.Range("A3:C3").Value = Array(SET_DONE, "", "Where processed RAW files go. Blank = a Done folder inside the RAW folder")
        ws.Range("A4:C4").Value = Array(SET_INVOICE, "", "Folder the Invoice <Tab Name> files are saved in")
        ws.Range("A5:C5").Value = Array(SET_MASTER, "", "Full path of the Master List workbook (the .xlsx can be left off)")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:C").AutoFit
        ws.Columns("B").ColumnWidth = 70
        ws.Activate
        ws.Range("B2").Select
        MsgBox "I added a Settings sheet to this workbook." & vbLf & vbLf & _
               "Fill in column B (the folders and the Master List file), save this workbook, " & _
               "and run the macro again.", vbInformation, APP_TITLE
        Exit Function
    End If

    rawFolder = ReadSetting(ws, SET_RAW)
    doneFolder = ReadSetting(ws, SET_DONE)
    invoiceFolder = ReadSetting(ws, SET_INVOICE)
    masterPath = ReadSetting(ws, SET_MASTER)
    If Len(doneFolder) = 0 And Len(rawFolder) > 0 Then doneFolder = rawFolder & "\Done"

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

' Column number of a header on row 3, ignoring capitals and extra spaces.
' Returns 0 and adds the header to "missing" when it isn't there.
Private Function FindHeader(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal header As String, _
                            ByRef missing As String) As Long
    Dim c As Long, v As Variant
    For c = 1 To lastCol
        v = ws.Cells(3, c).Value
        If Not IsError(v) Then
            If StrComp(Trim$(CStr(v)), header, vbTextCompare) = 0 Then
                FindHeader = c
                Exit Function
            End If
        End If
    Next c
    If Len(missing) > 0 Then missing = missing & ", "
    missing = missing & header
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

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim sh As Worksheet, tbl As ListObject
    For Each sh In wb.Worksheets
        For Each tbl In sh.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                TableExists = True
                Exit Function
            End If
        Next tbl
    Next sh
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
