Option Explicit

' =====================================================================
'  PDF -> Word -> Excel, step 1: see what Word's conversion produces
'
'  Opens a PDF in Word (Word converts it into an editable document),
'  then copies exactly what Word made into this workbook, unchanged.
'  Nothing is fixed yet: the point is to see HOW the conversion goes
'  wrong, so the fix in step 2 can undo it.
'
'  Creates (or clears) three sheets:
'    Summary     - counts: tables, rows, cells per row, text outside tables
'    Raw Tables  - every Word table, one Excel cell per Word cell
'                    pink   = a row with fewer cells than the fullest row
'                    yellow = a cell holding more than one line
'    Raw Text    - text Word did NOT put in a table, split at tab stops
'
'  Needs Excel and Word 2013 or later on Windows.
'  Run it from Excel: Alt+F8, pick DiagnosePdfConversion, click Run.
' =====================================================================

' True  = leave Word open on the converted document so you can compare.
' False = close Word when finished.
Private Const SHOW_WORD As Boolean = True

Private Const PINK As Long = 13551615     ' RGB(255, 199, 206)
Private Const YELLOW As Long = 10284031   ' RGB(255, 235, 156)

Public Sub DiagnosePdfConversion()
    Dim pdfPath As Variant
    pdfPath = Application.GetOpenFilename("PDF files (*.pdf),*.pdf", , "Pick the PDF to convert")
    If VarType(pdfPath) = vbBoolean Then Exit Sub         ' user clicked Cancel

    Dim wdApp As Object, doc As Object
    Dim tbl As Object, c As Object, p As Object, shp As Object
    Dim wsSummary As Worksheet, wsTables As Worksheet, wsText As Worksheet
    Dim cellsInRow() As Long
    Dim t As String
    Dim tblIdx As Long, r As Long, nRows As Long, maxCells As Long
    Dim unevenRows As Long, multiLine As Long
    Dim outRow As Long, sumRow As Long, textRow As Long
    Dim widestTable As Long, widestText As Long
    Dim paraCount As Long, boxCount As Long

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.StatusBar = "Opening the PDF in Word (this can take a minute)..."

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = SHOW_WORD
    wdApp.DisplayAlerts = 0                                ' 0 = skip Word's pop-up questions
    Set doc = wdApp.Documents.Open(Filename:=CStr(pdfPath), _
        ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False)

    Set wsSummary = CleanSheet("Summary", False)
    Set wsTables = CleanSheet("Raw Tables", True)
    Set wsText = CleanSheet("Raw Text", True)

    ' ---------- Raw Tables ----------
    wsTables.Columns("A:C").NumberFormat = "General"
    wsTables.Range("A1:C1").Value = Array("Table", "Word row", "Cells in row")
    wsSummary.Range("A6:E6").Value = Array("Table", "Rows", "Most cells in a row", _
        "Rows with fewer cells (pink)", "Cells with line breaks (yellow)")
    outRow = 2
    sumRow = 7

    For tblIdx = 1 To doc.Tables.Count
        Application.StatusBar = "Reading table " & tblIdx & " of " & doc.Tables.Count & "..."
        Set tbl = doc.Tables(tblIdx)
        nRows = tbl.Rows.Count
        ReDim cellsInRow(1 To nRows)
        multiLine = 0

        ' Range.Cells still works when Word has merged cells (Rows(i).Cells does not).
        For Each c In tbl.Range.Cells
            r = c.RowIndex
            cellsInRow(r) = cellsInRow(r) + 1
            t = CleanCellText(c.Range.Text)
            With wsTables.Cells(outRow + r - 1, 3 + c.ColumnIndex)
                .Value = t
                If InStr(t, vbLf) > 0 Then
                    .Interior.Color = YELLOW
                    .WrapText = True
                    multiLine = multiLine + 1
                End If
            End With
            If c.ColumnIndex > widestTable Then widestTable = c.ColumnIndex
        Next c

        maxCells = 0
        For r = 1 To nRows
            If cellsInRow(r) > maxCells Then maxCells = cellsInRow(r)
        Next r
        unevenRows = 0
        For r = 1 To nRows
            wsTables.Cells(outRow + r - 1, 1).Resize(1, 3).Value = Array(tblIdx, r, cellsInRow(r))
            If cellsInRow(r) < maxCells Then
                wsTables.Cells(outRow + r - 1, 3).Interior.Color = PINK
                unevenRows = unevenRows + 1
            End If
        Next r

        wsSummary.Cells(sumRow, 1).Resize(1, 5).Value = _
            Array(tblIdx, nRows, maxCells, unevenRows, multiLine)
        sumRow = sumRow + 1
        outRow = outRow + nRows + 1                         ' blank row between tables
    Next tblIdx

    For r = 1 To widestTable
        wsTables.Cells(1, 3 + r).Value = "Cell " & r
    Next r

    ' ---------- Raw Text ----------
    Application.StatusBar = "Reading text outside tables..."
    wsText.Range("A1").Value = "Source"
    textRow = 2

    For Each p In doc.Paragraphs
        If p.Range.Tables.Count = 0 Then
            t = p.Range.Text
            If HasVisibleText(t) Then
                paraCount = paraCount + 1
                WriteTextLines wsText, textRow, "Paragraph " & paraCount, t, widestText
            End If
        End If
    Next p

    ' Word sometimes puts converted text in floating text boxes instead.
    For Each shp In doc.Shapes
        t = ""
        On Error Resume Next
        t = shp.TextFrame.TextRange.Text                    ' errors for shapes with no text
        Err.Clear
        On Error GoTo Failed
        If HasVisibleText(t) Then
            boxCount = boxCount + 1
            WriteTextLines wsText, textRow, "Text box " & boxCount, t, widestText
        End If
    Next shp

    For r = 1 To widestText
        wsText.Cells(1, 1 + r).Value = "Part " & r
    Next r

    ' ---------- Summary ----------
    With wsSummary
        .Range("A1").Value = "PDF"
        .Range("B1").Value = CStr(pdfPath)
        .Range("A2").Value = "Word tables"
        .Range("B2").Value = doc.Tables.Count
        .Range("A3").Value = "Text paragraphs outside tables"
        .Range("B3").Value = paraCount
        .Range("A4").Value = "Text boxes with text"
        .Range("B4").Value = boxCount
        .Range("A1:A4").Font.Bold = True
        .Rows(6).Font.Bold = True
        If doc.Tables.Count = 0 Then
            .Range("A7").Value = "Word made no tables. Look at the Raw Text sheet instead."
        End If
    End With

    wsTables.Rows(1).Font.Bold = True
    wsText.Rows(1).Font.Bold = True
    FitColumns wsSummary
    FitColumns wsTables
    FitColumns wsText

    Application.ScreenUpdating = True
    wsSummary.Activate
    MsgBox "Done. Look at the Summary, Raw Tables and Raw Text sheets.", vbInformation

Done:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Not SHOW_WORD Then
        If Not doc Is Nothing Then doc.Close 0              ' 0 = don't save
        If Not wdApp Is Nothing Then wdApp.Quit
    End If
    Exit Sub

Failed:
    MsgBox "Stopped with an error: " & Err.Description, vbExclamation
    Resume Done
End Sub

' Returns the named sheet in this workbook, emptied (creates it if missing).
Private Function CleanSheet(ByVal sheetName As String, ByVal asText As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    ' Text format stops Excel turning "007" into 7 or "1-2" into a date.
    If asText Then
        ws.Cells.NumberFormat = "@"
    Else
        ws.Cells.NumberFormat = "General"
    End If
    Set CleanSheet = ws
End Function

' Word ends every cell's text with Chr(13) & Chr(7). Inside a cell, Chr(13) starts
' a new paragraph and Chr(11) is a line break; both become Excel line breaks.
Private Function CleanCellText(ByVal t As String) As String
    If Right$(t, 2) = Chr(13) & Chr(7) Then t = Left$(t, Len(t) - 2)
    t = Replace(t, Chr(7), "")
    t = Replace(t, Chr(13), vbLf)
    t = Replace(t, Chr(11), vbLf)
    CleanCellText = t
End Function

Private Function HasVisibleText(ByVal t As String) As Boolean
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr(13), "")
    t = Replace(t, Chr(11), "")
    t = Replace(t, Chr(7), "")
    HasVisibleText = Len(Trim$(t)) > 0
End Function

' Writes one Excel row per line of text; each tab stop starts a new column.
Private Sub WriteTextLines(ByVal ws As Worksheet, ByRef nextRow As Long, _
                           ByVal label As String, ByVal t As String, ByRef widest As Long)
    Dim lines As Variant, parts As Variant, i As Long, j As Long
    t = Replace(t, Chr(7), "")
    t = Replace(t, Chr(13), Chr(11))
    lines = Split(t, Chr(11))
    For i = LBound(lines) To UBound(lines)
        If HasVisibleText(CStr(lines(i))) Then
            parts = Split(CStr(lines(i)), vbTab)
            ws.Cells(nextRow, 1).Value = label
            For j = 0 To UBound(parts)
                ws.Cells(nextRow, 2 + j).Value = parts(j)
            Next j
            If UBound(parts) + 1 > widest Then widest = UBound(parts) + 1
            nextRow = nextRow + 1
        End If
    Next i
End Sub

Private Sub FitColumns(ByVal ws As Worksheet)
    Dim col As Range
    ws.UsedRange.Columns.AutoFit
    For Each col In ws.UsedRange.Columns
        If col.ColumnWidth > 50 Then col.ColumnWidth = 50
    Next col
End Sub
