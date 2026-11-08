' PAT recap: Add row and Tidy up for the OH PAT Recap workbook.
'
' First time: open OH PAT Recap - BLANK.xlsx, Alt F11, Insert, Module, then Ctrl A and Ctrl V to paste all of this in.
' Save it as a macro-enabled workbook (.xlsm), then Alt F8 and run TidyUp once to put the buttons on.
' To update, open its module (Alt F11, under Modules), press Ctrl A, and paste the new copy over the old one.
'
' Add row: click a line in a section (or its running total), press Add row. One empty row shows up at the
' bottom of that section, right above its running total, and every total counts it.
' Tidy up: takes out the empty rows on all three month tabs, keeps one empty row in a section with
' nothing typed, and says how many it took out on each tab.
' The workbook works the same without this macro. Every total is a formula; the buttons only save time.

Option Explicit

Private Const BOX_TITLE As String = "OH PAT recap"
Private Const FIRST_LINE_ROW As Long = 9
Private Const LAST_LINE_COL As Long = 8            ' column H, the last one a line uses
Private Const TOTAL_TAIL As String = " running total"
Private Const QTR_TAB As String = "QTR RECAP"
Private Const MONTH_TABS As String = "MONTH 1,MONTH 2,MONTH 3"
Private Const ADD_BUTTON As String = "Add row"
Private Const TIDY_BUTTON As String = "Tidy up"
Private Const BUTTON_HEIGHT As Double = 30
Private Const NOT_IN_SECTION As String = "Click a line in the section first, then press Add row. Nothing was changed."

Public Sub AddRow()
    Dim ws As Worksheet, firstRow As Long, totalRow As Long, lastLine As Long

    On Error GoTo Failed
    If PutButtons() Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Or TypeName(Selection) <> "Range" Then
        MsgBox NOT_IN_SECTION, vbInformation, BOX_TITLE
        Exit Sub
    End If
    Set ws = ActiveSheet
    If Not IsMonthTab(ws.Name) Or ActiveCell.Column > LAST_LINE_COL Or _
       Not FindSection(ws, ActiveCell.Row, firstRow, totalRow) Then
        MsgBox NOT_IN_SECTION, vbInformation, BOX_TITLE
        Exit Sub
    End If

    Application.ScreenUpdating = False
    ' Insert inside the section's range, so every SUM and SUMIFS stretches with it. The last line
    ' moves down to the bottom; the typed cells go back up and the bottom row is left empty.
    lastLine = totalRow - 1
    ws.Range(ws.Cells(lastLine, 1), ws.Cells(lastLine, LAST_LINE_COL)).Insert Shift:=xlDown
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, LAST_LINE_COL)).Copy ws.Cells(lastLine, 1)
    ws.Cells(totalRow, "D").ClearContents
    ws.Range(ws.Cells(totalRow, "G"), ws.Cells(totalRow, "H")).ClearContents
    ws.Rows(totalRow).RowHeight = ws.Rows(lastLine).RowHeight
    Application.CutCopyMode = False
    ws.Cells(totalRow, "G").Select
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "Add row stopped: " & Err.Description, vbExclamation, BOX_TITLE
End Sub

Public Sub TidyUp()
    Dim tabName As Variant, ws As Worksheet, said As String, gone As Long

    On Error GoTo Failed
    If PutButtons() Then Exit Sub
    Application.ScreenUpdating = False
    For Each tabName In Split(MONTH_TABS, ",")
        Set ws = Worksheets(CStr(tabName))
        gone = TidyTab(ws)
        said = said & vbLf & ws.Name & ": " & gone & IIf(gone = 1, " row", " rows")
    Next tabName
    Application.ScreenUpdating = True
    MsgBox "Tidy up took out:" & said, vbInformation, BOX_TITLE
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "Tidy up stopped: " & Err.Description, vbExclamation, BOX_TITLE
End Sub

Private Function IsMonthTab(ByVal tabName As String) As Boolean
    IsMonthTab = InStr("," & MONTH_TABS & ",", "," & tabName & ",") > 0
End Function

' A line row has its section's name in E: In price, DIV OWE, Additive owed or Reduced GP.
Private Function IsLine(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Select Case CStr(ws.Cells(r, "E").Text)
        Case "In price", "DIV OWE", "Additive owed", "Reduced GP"
            IsLine = True
    End Select
End Function

' A running total row has its label in D, like "In price running total".
Private Function IsTotal(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim label As String
    label = CStr(ws.Cells(r, "D").Text)
    IsTotal = (Len(label) > Len(TOTAL_TAIL) And Right$(label, Len(TOTAL_TAIL)) = TOTAL_TAIL)
End Function

' Finds the section a row is in: its first line row and its running total row.
Private Function FindSection(ByVal ws As Worksheet, ByVal r As Long, ByRef firstRow As Long, _
                             ByRef totalRow As Long) As Boolean
    Dim i As Long

    If r < FIRST_LINE_ROW Then Exit Function
    If IsTotal(ws, r) Then
        totalRow = r
    ElseIf IsLine(ws, r) Then
        i = r
        Do While IsLine(ws, i)
            i = i + 1
        Loop
        If Not IsTotal(ws, i) Then Exit Function
        totalRow = i
    Else
        Exit Function
    End If
    i = totalRow - 1
    Do While i >= FIRST_LINE_ROW And IsLine(ws, i)
        i = i - 1
    Loop
    firstRow = i + 1
    FindSection = (firstRow < totalRow)
End Function

' Takes the empty rows out of every section of one tab. Returns how many rows went.
Private Function TidyTab(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long, firstRow As Long, endRow As Long, typed As Long, i As Long, gone As Long

    r = FIRST_LINE_ROW
    Do
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        Do While r <= lastRow And Not IsLine(ws, r)
            r = r + 1
        Loop
        If r > lastRow Then Exit Do
        firstRow = r
        endRow = r
        Do While IsLine(ws, endRow + 1)
            endRow = endRow + 1
        Loop
        typed = 0
        For i = firstRow To endRow
            If Not IsEmptyLine(ws, i) Then typed = typed + 1
        Next i
        gone = 0
        For i = endRow To firstRow Step -1
            If IsEmptyLine(ws, i) And (typed > 0 Or i > firstRow) Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, LAST_LINE_COL)).Delete Shift:=xlUp
                gone = gone + 1
            End If
        Next i
        TidyTab = TidyTab + gone
        r = endRow - gone + 1
    Loop
End Function

' A row is empty when nothing is typed in it: no note in D, no gallons in G, no product in H.
Private Function IsEmptyLine(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsEmptyLine = (Len(ws.Cells(r, "D").Formula) = 0 And Len(ws.Cells(r, "G").Formula) = 0 And _
                   Len(ws.Cells(r, "H").Formula) = 0)
End Function

' The first run puts Add row and Tidy up on the month tabs and Tidy up on QTR RECAP, and says so.
' Returns True when it did, so that run does nothing else.
Private Function PutButtons() As Boolean
    Dim tabName As Variant, missing As Boolean

    missing = (Worksheets(QTR_TAB).Buttons.Count = 0)
    For Each tabName In Split(MONTH_TABS, ",")
        If Worksheets(CStr(tabName)).Buttons.Count = 0 Then missing = True
    Next tabName
    If Not missing Then Exit Function

    MakeButton Worksheets(QTR_TAB), "H1:I2", TIDY_BUTTON, "TidyUp"
    For Each tabName In Split(MONTH_TABS, ",")
        MakeButton Worksheets(CStr(tabName)), "J1:K2", ADD_BUTTON, "AddRow"
        MakeButton Worksheets(CStr(tabName)), "L1:M2", TIDY_BUTTON, "TidyUp"
    Next tabName
    MsgBox "The Add row and Tidy up buttons are on. Press one to use it.", vbInformation, BOX_TITLE
    PutButtons = True
End Function

Private Sub MakeButton(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, ByVal macro As String)
    Dim spot As Range, b As Button

    Set spot = ws.Range(address)
    Set b = ws.Buttons.Add(spot.Left + 2, spot.Top + 1, spot.Width - 4, BUTTON_HEIGHT)
    b.Caption = caption
    b.OnAction = macro
    b.Placement = xlFreeFloating
End Sub
