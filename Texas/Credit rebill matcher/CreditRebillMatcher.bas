Option Explicit

'==================================================================
'  CREDIT / REBILL MATCHER
'
'  Pulls net-zero credit + rebill sets out of an order list so the
'  remainder can be uploaded to Texas (which rejects negatives).
'  Credits with no offsetting rebill are quarantined and flagged --
'  those are the ones that may mean amending a prior return.
'
'  Run it with the sheet active. It will ask for:
'     1. any cell in the BOL / order-number column
'     2. any cell in the GALLONS column
'     3. the first row of real data
'==================================================================

'--- Settings -----------------------------------------------------
' Leave True and the macro copies the sheet and works on the copy,
' so the original is untouched (Ctrl+Z can't undo a macro, so the
' copy is your undo). Set False to work on the sheet itself.
Private Const WORK_ON_A_COPY As Boolean = True

' A BOL with more candidate lines than this is left for you to check
' by hand, so one huge BOL can't hang Excel.
Private Const MAX_SUBSET_CANDIDATES As Long = 20
'------------------------------------------------------------------

Private Type RowInfo
    SheetRow As Long
    Key As String
    Amount As Currency
    Matched As Boolean
    GroupNo As Long
    Bad As Boolean
End Type


Public Sub MatchCreditsAndRebills()

    '--- all declarations up front -------------------------------
    Dim src As Worksheet, ws As Worksheet
    Dim rng As Range
    Dim v As Variant, raw As Variant, sorted As Variant, pal As Variant
    Dim amtCell As Variant, k As Variant
    Dim keys() As Variant
    Dim inf() As RowInfo
    Dim dict As Object
    Dim priorCalc As XlCalculation
    Dim madeCopy As Boolean, endOfRun As Boolean, doFill As Boolean
    Dim orderCol As Long, amtCol As Long, startRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, n As Long, r As Long, rr As Long, runStart As Long
    Dim groupCount As Long, tooComplex As Long
    Dim keepCount As Long, orphanCount As Long
    Dim matchedCount As Long, reviewCount As Long
    Dim sec As Long, clr As Long
    Dim h1 As Long, h2 As Long, h3 As Long
    Dim rowSec1 As Long, rowSec2 As Long, rowSec3 As Long
    Dim errNum As Long
    Dim chk As Currency, keepTotal As Currency, orphanTotal As Currency
    Dim msg As String, eDesc As String
    '--------------------------------------------------------------

    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set src = ActiveSheet

    '--- prompts --------------------------------------------------
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Click ANY cell inside the column containing BOLs:", _
        Title:="Select order-number column", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    orderCol = rng.Column

    Set rng = Nothing
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Click ANY cell inside the column containing GALLONS:", _
        Title:="Select amount column", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    amtCol = rng.Column

    If orderCol = amtCol Then
        MsgBox "You picked the same column twice.", vbExclamation
        Exit Sub
    End If

    v = Application.InputBox( _
        Prompt:="Enter the first row number containing actual data:", _
        Title:="Select start row", Default:=2, Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub      ' Cancel
    startRow = CLng(v)
    If startRow < 1 Then Exit Sub

    '--- save Excel state, then arm the handler -------------------
    priorCalc = Application.Calculation
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- work on a copy so the original survives ------------------
    If WORK_ON_A_COPY Then
        src.Copy After:=src
        Set ws = src.Parent.Sheets(src.Index + 1)
        ws.Name = SafeSheetName(src.Parent, src.Name & " MATCHED")
        madeCopy = True
    Else
        Set ws = src
    End If

    '--- trim trailing Grand Total / blank rows -------------------
    r = ws.Cells(ws.Rows.Count, orderCol).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > r Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    Do While r >= startRow
        If IsTotalOrBlankRow(ws, r, orderCol, amtCol) Then
            ws.Rows(r).Delete
            r = r - 1
        Else
            Exit Do
        End If
    Loop

    lastRow = ws.Cells(ws.Rows.Count, orderCol).End(xlUp).Row
    If lastRow < startRow Then
        MsgBox "No data rows found at or below row " & startRow & ".", vbExclamation
        GoTo CleanExit
    End If

    lastCol = LastUsedColumn(ws)
    If lastCol < orderCol Then lastCol = orderCol
    If lastCol < amtCol Then lastCol = amtCol

    '--- read everything once -------------------------------------
    raw = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)).Value

    n = lastRow - startRow + 1
    ReDim inf(1 To n)

    For i = 1 To n
        inf(i).SheetRow = startRow + i - 1
        inf(i).Key = NormalizeKey(raw(i, orderCol))
        amtCell = raw(i, amtCol)

        If Len(inf(i).Key) = 0 Then
            inf(i).Bad = True
        ElseIf IsError(amtCell) Then
            inf(i).Bad = True
        ElseIf IsEmpty(amtCell) Then
            inf(i).Bad = True
        ElseIf Not IsNumeric(amtCell) Then
            inf(i).Bad = True
        Else
            inf(i).Amount = CCur(amtCell)
        End If
    Next i

    '--- bucket the rows by BOL -----------------------------------
    ' Matching only ever looks at the lines on one BOL, never the
    ' whole sheet, which keeps it fast.
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 0                              ' keys already normalized

    For i = 1 To n
        If Not inf(i).Bad Then
            If Not dict.Exists(inf(i).Key) Then dict.Add inf(i).Key, New Collection
            dict(inf(i).Key).Add i
        End If
    Next i

    '--- match ----------------------------------------------------
    For Each k In dict.Keys
        MatchOneOrder dict(k), inf, groupCount, tooComplex
    Next k

    '--- classify each row into an output section -----------------
    '   0 = keep / upload        2 = matched net-zero group
    '   1 = orphan credit        3 = unreadable, needs a human
    ReDim keys(1 To n, 1 To 3)

    For i = 1 To n
        If inf(i).Bad Then
            sec = 3: reviewCount = reviewCount + 1
        ElseIf inf(i).Matched Then
            sec = 2: matchedCount = matchedCount + 1
        ElseIf inf(i).Amount < 0 Then
            sec = 1: orphanCount = orphanCount + 1
        Else
            sec = 0: keepCount = keepCount + 1
        End If
        keys(i, 1) = sec
        keys(i, 2) = inf(i).GroupNo
        keys(i, 3) = i                                ' keeps original order stable
    Next i

    '--- reorder the sheet with ONE sort --------------------------
    ' Sorting (rather than copy / clear / delete row by row) keeps
    ' formulas, number formats and dates intact, and costs the same
    ' whether there are 50 rows or 5,000.
    h1 = lastCol + 2: h2 = h1 + 1: h3 = h2 + 1

    ws.Range(ws.Cells(startRow, h1), ws.Cells(lastRow, h3)).Value = keys
    ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)).Interior.Pattern = xlNone

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(startRow, h1), ws.Cells(lastRow, h1)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(startRow, h2), ws.Cells(lastRow, h2)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(startRow, h3), ws.Cells(lastRow, h3)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, h3))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With

    '--- colour each contiguous block in one shot -----------------
    sorted = ws.Range(ws.Cells(startRow, h1), ws.Cells(lastRow, h2)).Value

    pal = Array(RGB(255, 242, 204), RGB(217, 234, 211), RGB(208, 225, 242), RGB(234, 215, 240), _
                RGB(252, 229, 205), RGB(214, 240, 238), RGB(226, 231, 196), RGB(240, 219, 219))

    runStart = 1
    For rr = 1 To n
        If rr = n Then
            endOfRun = True
        Else
            endOfRun = (sorted(rr, 1) <> sorted(rr + 1, 1)) Or (sorted(rr, 2) <> sorted(rr + 1, 2))
        End If

        If endOfRun Then
            doFill = True
            Select Case CLng(sorted(runStart, 1))
                Case 0
                    doFill = False                               ' clean upload rows stay plain
                Case 1
                    clr = RGB(255, 199, 206)                     ' orphan credit -- red
                Case 2
                    clr = CLng(pal((CLng(sorted(runStart, 2)) - 1) Mod 8))
                Case Else
                    clr = RGB(217, 217, 217)                     ' unreadable -- grey
            End Select
            If doFill Then
                ws.Range(ws.Cells(startRow + runStart - 1, 1), _
                         ws.Cells(startRow + rr - 1, lastCol)).Interior.Color = clr
            End If
            runStart = rr + 1
        End If
    Next rr

    ws.Range(ws.Cells(1, h1), ws.Cells(1, h3)).EntireColumn.Delete

    '--- banners, inserted bottom-up so row numbers hold ----------
    rowSec1 = startRow + keepCount
    rowSec2 = rowSec1 + orphanCount
    rowSec3 = rowSec2 + matchedCount

    If reviewCount > 0 Then
        InsertBanner ws, rowSec3, "--- COULD NOT READ THE BOL OR THE GALLONS - CHECK THESE BY HAND ---"
    End If
    If matchedCount > 0 Then
        InsertBanner ws, rowSec2, "--- MATCHED NET-ZERO GROUPS (VERIFY, THEN DELETE) ---"
    End If
    If orphanCount > 0 Then
        InsertBanner ws, rowSec1, "--- CREDITS WITH NO MATCHING REBILL - LIKELY NEEDS AN AMENDED RETURN ---"
    End If

    '--- self-check: every matched group must sum to exactly zero --
    For i = 1 To n
        If inf(i).Matched Then
            chk = chk + inf(i).Amount
        ElseIf Not inf(i).Bad Then
            If inf(i).Amount < 0 Then
                orphanTotal = orphanTotal + inf(i).Amount
            Else
                keepTotal = keepTotal + inf(i).Amount
            End If
        End If
    Next i

    msg = n & " data rows across " & dict.Count & " BOLs." & vbCrLf & vbCrLf & _
          keepCount & " rows left at the top to upload (" & Format$(keepTotal, "#,##0.00") & " gal)" & vbCrLf & _
          matchedCount & " rows in " & groupCount & " net-zero groups, moved down" & vbCrLf & _
          orphanCount & " orphan credits (" & Format$(orphanTotal, "#,##0.00") & " gal) - RED"
    If reviewCount > 0 Then
        msg = msg & vbCrLf & reviewCount & " rows unreadable - GREY"
    End If
    If tooComplex > 0 Then
        msg = msg & vbCrLf & vbCrLf & "NOTE: " & tooComplex & " BOL(s) had more than " & _
              MAX_SUBSET_CANDIDATES & " candidate lines and were left for you to eyeball."
    End If
    If chk <> 0 Then
        msg = msg & vbCrLf & vbCrLf & "WARNING: matched groups net to " & _
              Format$(chk, "#,##0.00") & " instead of zero. Do not upload."
    End If
    If madeCopy Then
        msg = msg & vbCrLf & vbCrLf & "Original sheet untouched; this is """ & ws.Name & """."
    End If

    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, "Credit / rebill matching"
    Exit Sub

CleanExit:
    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    errNum = Err.Number
    eDesc = Err.Description
    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    msg = "Stopped on error " & errNum & ": " & eDesc & vbCrLf & vbCrLf & _
          "Excel's calculation and screen settings have been put back. "
    If WORK_ON_A_COPY Then
        msg = msg & "Your original sheet was not touched."
    Else
        msg = msg & "The sheet was being edited in place - check it before continuing."
    End If
    MsgBox msg, vbCritical, "Credit / rebill matching"
End Sub


'==================================================================
'  Matching for the lines on ONE BOL
'==================================================================
Private Sub MatchOneOrder(ByVal members As Collection, ByRef inf() As RowInfo, _
                          ByRef groupCount As Long, ByRef tooComplex As Long)

    Dim idx As Variant, jdx As Variant
    Dim a As Long, b As Long
    Dim total As Currency
    Dim hasPos As Boolean, hasNeg As Boolean
    Dim leftovers As Long

    If members.Count < 2 Then Exit Sub

    For Each idx In members
        a = CLng(idx)
        total = total + inf(a).Amount
        If inf(a).Amount > 0 Then hasPos = True
        If inf(a).Amount < 0 Then hasNeg = True
    Next idx

    If Not (hasPos And hasNeg) Then Exit Sub          ' nothing to net against

    ' (a) the whole BOL nets to zero -- take the lot, even when both
    '     sides have more than one line (2 credits against 3 rebills).
    If total = 0 Then
        groupCount = groupCount + 1
        For Each idx In members
            inf(CLng(idx)).Matched = True
            inf(CLng(idx)).GroupNo = groupCount
        Next idx
        Exit Sub
    End If

    ' (b) clean 1:1 pairs of the same size. Most lines match here.
    For Each idx In members
        a = CLng(idx)
        If Not inf(a).Matched Then
            If inf(a).Amount < 0 Then
                For Each jdx In members
                    b = CLng(jdx)
                    If b <> a Then
                        If Not inf(b).Matched Then
                            If inf(b).Amount = -inf(a).Amount Then
                                groupCount = groupCount + 1
                                inf(a).Matched = True: inf(a).GroupNo = groupCount
                                inf(b).Matched = True: inf(b).GroupNo = groupCount
                                Exit For
                            End If
                        End If
                    End If
                Next jdx
            End If
        End If
    Next idx

    ' (c) one credit against a combination of rebills, then (d) the reverse
    SubsetPass members, inf, groupCount, tooComplex, True
    SubsetPass members, inf, groupCount, tooComplex, False

    ' (e) if whatever is still unmatched happens to net to zero, take it
    total = 0
    leftovers = 0
    For Each idx In members
        a = CLng(idx)
        If Not inf(a).Matched Then
            total = total + inf(a).Amount
            leftovers = leftovers + 1
        End If
    Next idx

    If leftovers >= 2 And total = 0 Then
        groupCount = groupCount + 1
        For Each idx In members
            a = CLng(idx)
            If Not inf(a).Matched Then
                inf(a).Matched = True
                inf(a).GroupNo = groupCount
            End If
        Next idx
    End If
End Sub


Private Sub SubsetPass(ByVal members As Collection, ByRef inf() As RowInfo, _
                       ByRef groupCount As Long, ByRef tooComplex As Long, _
                       ByVal creditDriven As Boolean)

    Dim idx As Variant, jdx As Variant
    Dim a As Long, b As Long
    Dim isDriver As Boolean, wanted As Boolean
    Dim cand As Collection, chosen As Collection

    For Each idx In members
        a = CLng(idx)
        If Not inf(a).Matched Then

            If creditDriven Then
                isDriver = (inf(a).Amount < 0)
            Else
                isDriver = (inf(a).Amount > 0)
            End If

            If isDriver Then
                Set cand = New Collection
                For Each jdx In members
                    b = CLng(jdx)
                    If b <> a Then
                        If Not inf(b).Matched Then
                            If creditDriven Then
                                wanted = (inf(b).Amount > 0)
                            Else
                                wanted = (inf(b).Amount < 0)
                            End If
                            If wanted Then cand.Add b
                        End If
                    End If
                Next jdx

                If cand.Count > MAX_SUBSET_CANDIDATES Then
                    tooComplex = tooComplex + 1
                ElseIf cand.Count > 0 Then
                    Set chosen = New Collection
                    If FindSubset(cand, 1, Abs(inf(a).Amount), chosen, inf) Then
                        groupCount = groupCount + 1
                        inf(a).Matched = True
                        inf(a).GroupNo = groupCount
                        For Each jdx In chosen
                            inf(CLng(jdx)).Matched = True
                            inf(CLng(jdx)).GroupNo = groupCount
                        Next jdx
                    End If
                End If
            End If
        End If
    Next idx
End Sub


' Gallons are held as Currency, which adds up exactly, so a group
' has to net to exactly zero. No rounding slack.
Private Function FindSubset(ByVal cand As Collection, ByVal startIndex As Long, _
                            ByVal target As Currency, ByRef chosen As Collection, _
                            ByRef inf() As RowInfo) As Boolean

    Dim thisIdx As Long
    Dim ok As Boolean

    If target = 0 Then
        FindSubset = (chosen.Count > 0)          ' never "match" nothing at all
        Exit Function
    End If
    If target < 0 Or startIndex > cand.Count Then
        FindSubset = False
        Exit Function
    End If

    thisIdx = CLng(cand(startIndex))

    chosen.Add thisIdx
    ok = FindSubset(cand, startIndex + 1, target - Abs(inf(thisIdx).Amount), chosen, inf)
    If ok Then
        FindSubset = True
        Exit Function
    End If
    chosen.Remove chosen.Count

    ok = FindSubset(cand, startIndex + 1, target, chosen, inf)
    FindSubset = ok
End Function


'==================================================================
'  Helpers
'==================================================================

' Cleans up a BOL before comparing, so one stored as text in one row
' and as a number in another (or with leading zeros or spaces) still
' matches instead of landing in the "needs an amended return" pile.
Private Function NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    Dim j As Long
    Dim allDigits As Boolean

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then
        s = Format$(v, "yyyy-mm-dd")
    ElseIf IsNumeric(v) Then
        s = Format$(v, "General Number")
    Else
        s = CStr(v)
    End If

    s = UCase$(Trim$(s))
    s = Replace$(s, Chr$(160), "")               ' non-breaking space
    s = Replace$(s, " ", "")
    If Len(s) = 0 Then Exit Function

    allDigits = True
    For j = 1 To Len(s)
        If Mid$(s, j, 1) < "0" Or Mid$(s, j, 1) > "9" Then
            allDigits = False
            Exit For
        End If
    Next j

    If allDigits Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If

    NormalizeKey = s
End Function


' Reads each cell's shown text, so a stray word in the gallons
' column can't stop the macro with an error.
Private Function IsTotalOrBlankRow(ws As Worksheet, ByVal r As Long, _
                                   ByVal orderCol As Long, ByVal amtCol As Long) As Boolean
    Dim colA As String, colOrder As String, colAmt As String
    colA = UCase$(Trim$(ws.Cells(r, 1).Text))
    colOrder = UCase$(Trim$(ws.Cells(r, orderCol).Text))
    colAmt = UCase$(Trim$(ws.Cells(r, amtCol).Text))

    IsTotalOrBlankRow = (Len(colOrder) = 0) _
                     Or (InStr(colA, "TOTAL") > 0) _
                     Or (InStr(colOrder, "TOTAL") > 0) _
                     Or (InStr(colAmt, "TOTAL") > 0)
End Function


Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = f.Column
    End If
End Function


Private Sub InsertBanner(ws As Worksheet, ByVal atRow As Long, ByVal txt As String)
    ws.Range(ws.Rows(atRow), ws.Rows(atRow + 1)).Insert Shift:=xlDown
    ws.Range(ws.Rows(atRow), ws.Rows(atRow + 1)).Interior.Pattern = xlNone
    With ws.Cells(atRow + 1, 1)
        .Value = txt
        .Font.Bold = True
    End With
End Sub


Private Function SafeSheetName(wb As Workbook, ByVal nm As String) As String
    Dim stem As String, attempt As String
    Dim c As Variant, i As Long

    stem = nm
    For Each c In Array("\", "/", "*", "?", ":", "[", "]")
        stem = Replace$(stem, CStr(c), "-")
    Next c
    stem = Left$(Trim$(stem), 27)

    attempt = stem
    i = 1
    Do While SheetExists(wb, attempt)
        i = i + 1
        attempt = stem & " " & i
    Loop
    SafeSheetName = attempt
End Function


Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not (s Is Nothing)
End Function
