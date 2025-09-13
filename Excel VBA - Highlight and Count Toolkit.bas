Attribute VB_Name = "Module1"
Option Explicit

'===============================
' Utility: Find a column by header OR letter
'===============================
Private Function ResolveColumnIndex(ws As Worksheet, headerOrLetter As String) As Long
    Dim rngHeaders As Range
    Dim cell As Range
    Dim colLetter As String
    Dim colIdx As Long

    headerOrLetter = Trim(headerOrLetter)
    If headerOrLetter = "" Then
        ResolveColumnIndex = 0
        Exit Function
    End If

    ' 1) Try header match in the first row of UsedRange (case-insensitive)
    Set rngHeaders = ws.Range(ws.UsedRange.Rows(1).Address)
    For Each cell In rngHeaders.Cells
        If StrComp(CStr(cell.Value), headerOrLetter, vbTextCompare) = 0 Then
            ResolveColumnIndex = cell.Column
            Exit Function
        End If
    Next cell

    ' 2) Try a column letter like "A", "B", "AA", etc.
    On Error Resume Next
    colLetter = UCase(headerOrLetter)
    colIdx = Range(colLetter & "1").Column
    If Err.Number = 0 Then
        ResolveColumnIndex = colIdx
        Exit Function
    End If
    On Error GoTo 0

    ' Not found
    ResolveColumnIndex = 0
End Function

'===============================
' Utility: Get the first empty column to the right of used data
'===============================
Private Function NextEmptyColumn(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        NextEmptyColumn = 1
    Else
        NextEmptyColumn = lastCell.Column + 1
    End If
End Function

'===============================
' 1) Highlight & Count matches in a chosen field
'===============================
Public Sub HighlightAndCountMatches()
    Dim ws As Worksheet
    Dim fld As String
    Dim modeChoice As String
    Dim termsInput As String
    Dim terms() As String
    Dim dataStartRow As Long, r As Long
    Dim colIdx As Long, nextCol As Long
    Dim usedRng As Range
    Dim cell As Range
    Dim matchCount As Long
    Dim i As Long
    Dim rowMatched As Boolean
    Dim clr As Long

    Set ws = ActiveSheet
    If ws.UsedRange Is Nothing Then
        MsgBox "No data found on the active sheet.", vbExclamation
        Exit Sub
    End If
    Set usedRng = ws.UsedRange

    ' Ask for a field (header name OR column letter)
    fld = InputBox("Field to search (header name or column letter, e.g., 'Department' or 'B'):", _
                   "Highlight & Count")
    If fld = "" Then Exit Sub

    colIdx = ResolveColumnIndex(ws, fld)
    If colIdx = 0 Then
        MsgBox "Could not resolve '" & fld & "' to a column. Check header/letter.", vbCritical
        Exit Sub
    End If

    ' Search mode: 1=Exact, 2=Contains (default 2)
    modeChoice = InputBox("Search mode:" & vbCrLf & _
                          "1 = Exact match" & vbCrLf & _
                          "2 = Contains (recommended)", _
                          "Highlight & Count", "2")
    If modeChoice = "" Then Exit Sub
    If modeChoice <> "1" And modeChoice <> "2" Then
        MsgBox "Invalid mode. Please enter 1 or 2.", vbExclamation
        Exit Sub
    End If

    ' Terms (comma-separated)
    termsInput = InputBox("Enter search term(s). For multiple, separate with commas." & vbCrLf & _
                          "Examples: 'Alice'   or   'IT,Operations'", _
                          "Highlight & Count")
    If Trim(termsInput) = "" Then Exit Sub

    terms = Split(termsInput, ",")
    For i = LBound(terms) To UBound(terms)
        terms(i) = Trim(terms(i))
    Next i

    ' Optional: clear prior highlight in the data area (light yellow only)
    ' Comment this block out if you want to keep past highlights.
    usedRng.Interior.ColorIndex = xlNone
    usedRng.Font.Bold = False

    ' Highlight color (light yellow)
    clr = 65535

    dataStartRow = usedRng.Row + 1   ' assume first row of UsedRange is headers
    For r = dataStartRow To usedRng.Row + usedRng.Rows.Count - 1
        rowMatched = False
        Dim v As String
        v = CStr(ws.Cells(r, colIdx).Value)

        If Len(v) > 0 Then
            If modeChoice = "1" Then
                ' Exact match against any provided term (case-insensitive)
                For i = LBound(terms) To UBound(terms)
                    If StrComp(v, terms(i), vbTextCompare) = 0 Then
                        rowMatched = True
                        Exit For
                    End If
                Next i
            Else
                ' Contains match
                For i = LBound(terms) To UBound(terms)
                    If InStr(1, v, terms(i), vbTextCompare) > 0 Then
                        rowMatched = True
                        Exit For
                    End If
                Next i
            End If
        End If

        If rowMatched Then
            ' Highlight the whole row within the used range width
            ws.Range(ws.Cells(r, usedRng.Column), _
                     ws.Cells(r, usedRng.Column + usedRng.Columns.Count - 1)).Interior.Color = clr
            ws.Range(ws.Cells(r, usedRng.Column), _
                     ws.Cells(r, usedRng.Column + usedRng.Columns.Count - 1)).Font.Bold = True
            matchCount = matchCount + 1
        End If
    Next r

    ' Output summary in the first empty column to the right
    nextCol = NextEmptyColumn(ws)
    ws.Cells(usedRng.Row, nextCol).Value = "Match Count"
    ws.Cells(usedRng.Row + 1, nextCol).Value = matchCount
    ws.Cells(usedRng.Row + 2, nextCol).Value = "Field:"
    ws.Cells(usedRng.Row + 2, nextCol + 1).Value = fld
    ws.Cells(usedRng.Row + 3, nextCol).Value = "Mode:"
    ws.Cells(usedRng.Row + 3, nextCol + 1).Value = IIf(modeChoice = "1", "Exact", "Contains")
    ws.Cells(usedRng.Row + 4, nextCol).Value = "Terms:"
    ws.Cells(usedRng.Row + 4, nextCol + 1).Value = termsInput

    ws.Columns(nextCol).EntireColumn.AutoFit
    ws.Columns(nextCol + 1).EntireColumn.AutoFit

    MsgBox "Done. Matches highlighted and counted: " & matchCount, vbInformation
End Sub

'===============================
' 2) Frequency summary by a chosen field (e.g., Department)
'===============================
Public Sub SummarizeCountsByField()
    Dim ws As Worksheet
    Dim fld As String
    Dim colIdx As Long
    Dim usedRng As Range
    Dim dataStartRow As Long, r As Long
    Dim dict As Object
    Dim key As Variant
    Dim nextCol As Long
    Dim valueText As String

    Set ws = ActiveSheet
    If ws.UsedRange Is Nothing Then
        MsgBox "No data found on the active sheet.", vbExclamation
        Exit Sub
    End If
    Set usedRng = ws.UsedRange

    fld = InputBox("Field to summarize (header name or column letter, e.g., 'Department' or 'C'):", _
                   "Summarize Counts")
    If fld = "" Then Exit Sub

    colIdx = ResolveColumnIndex(ws, fld)
    If colIdx = 0 Then
        MsgBox "Could not resolve '" & fld & "' to a column. Check header/letter.", vbCritical
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dataStartRow = usedRng.Row + 1

    For r = dataStartRow To usedRng.Row + usedRng.Rows.Count - 1
        valueText = CStr(ws.Cells(r, colIdx).Value)
        If valueText <> "" Then
            If Not dict.Exists(valueText) Then
                dict.Add valueText, 1
            Else
                dict(valueText) = dict(valueText) + 1
            End If
        End If
    Next r

    ' Output frequency table at the first empty column to the right
    nextCol = NextEmptyColumn(ws)
    ws.Cells(usedRng.Row, nextCol).Value = "Value"
    ws.Cells(usedRng.Row, nextCol + 1).Value = "Count"

    Dim outRow As Long
    outRow = usedRng.Row + 1

    For Each key In dict.Keys
        ws.Cells(outRow, nextCol).Value = CStr(key)
        ws.Cells(outRow, nextCol + 1).Value = CLng(dict(key))
        outRow = outRow + 1
    Next key

    ' Make it neat
    With ws.Range(ws.Cells(usedRng.Row, nextCol), ws.Cells(outRow - 1, nextCol + 1))
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Bold = False
        ws.Cells(usedRng.Row, nextCol).Font.Bold = True
        ws.Cells(usedRng.Row, nextCol + 1).Font.Bold = True
    End With

    MsgBox "Summary created in columns " & Split(Cells(1, nextCol).Address(True, False), "$")(0) & ":" & _
           Split(Cells(1, nextCol + 1).Address(True, False), "$")(0), vbInformation
End Sub

'===============================
' Optional: Clear all highlights (resets formatting in UsedRange)
'===============================
Public Sub ClearHighlights()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.UsedRange Is Nothing Then Exit Sub

    With ws.UsedRange
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
    End With

    MsgBox "Highlights cleared.", vbInformation
End Sub


