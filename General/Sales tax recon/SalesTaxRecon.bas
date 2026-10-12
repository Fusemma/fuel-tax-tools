Attribute VB_Name = "SalesTaxRecon"
Option Explicit

'==================================================================
'  TERSE  vs  AR SALES  RECONCILER
'
'  Paste the two reports onto their own tabs, run CompareReports,
'  and it shows you only what does NOT agree.
'
'  Columns are found by their HEADER TEXT, not by column letter, so
'  it keeps working if the export gains or loses a column or moves
'  one. The header row is found too, so title lines above it are ok.
'
'  Lines are totalled per Order Number before comparing, because one
'  order is several lines on TERSE (one per order item) and can be
'  several lines on AR SALES (one per tax class).
'
'  The RECON sheet gives you three things:
'     1. a summary and the grand-total tie-out
'     2. one line per order that did not agree, with the difference
'     3. the FULL ORIGINAL ROWS behind those orders, copied from
'        both reports exactly as they look on the report tabs
'
'  It never writes to your report tabs.
'==================================================================

'==================================================================
'  CONFIGURATION -- everything you would ever change is right here
'==================================================================

' The two tabs.
Private Const SHEET_A As String = "TERSE"
Private Const SHEET_B As String = "AR SALES"

' The column whose value identifies the same transaction on both.
Private Const KEY_A As String = "Order Number"
Private Const KEY_B As String = "Order Number"

' What to compare. Left of the ":" is the TERSE header, right is the
' AR SALES header. Use " + " to add columns together on either side.
' Use " | " to compare more than one pair in the same run.
'
'   "Total Product : Taxable Sales"
'   "Total Product : Taxable Sales + Non Taxable Sales"
'   "Total Product : Taxable Sales | Total Sales Taxes : Total Sales Tax"
'
Private Const AMOUNT_PAIRS As String = "Total Product : Taxable Sales"

' How far apart two totals may be and still count as agreeing.
' 0 means to the penny.
Private Const MATCH_TOLERANCE As Currency = 0

' How many rows down to look for the header row.
Private Const HEADER_SEARCH_ROWS As Long = 10

' Name of the sheet the results land on. Rebuilt from scratch each run.
Private Const OUT_SHEET As String = "RECON"

'==================================================================

Private Const ISSUE_ONLY_A As Long = 1
Private Const ISSUE_ONLY_B As Long = 2
Private Const ISSUE_DIFFER As Long = 3


Public Sub CompareReports()

    '--- declarations ---------------------------------------------
    Dim wb As Workbook
    Dim wsA As Worksheet, wsB As Worksheet, outWs As Worksheet
    Dim priorCalc As XlCalculation
    Dim dict As Object, excKeys As Object
    Dim pairSpec As Variant, halves As Variant
    Dim headsA() As String, headsB() As String
    Dim colsA() As Variant, colsB() As Variant
    Dim hdrRowA As Long, hdrRowB As Long
    Dim keyColA As Long, keyColB As Long
    Dim lastColA As Long, lastColB As Long
    Dim lastA As Long, lastB As Long
    Dim nPairs As Long
    Dim maxKeys As Long, keyCount As Long
    Dim i As Long, j As Long, p As Long, idx As Long, rowOut As Long
    Dim errNum As Long
    Dim missing As String, msg As String, eDesc As String, spec As String

    Dim keyRaw() As Variant
    Dim cntA() As Long, cntB() As Long
    Dim seenA() As Boolean, seenB() As Boolean
    Dim sumA() As Currency, sumB() As Currency
    Dim grandA() As Currency, grandB() As Currency
    Dim issue() As Long, biggest() As Currency
    Dim orderIdx() As Long
    Dim okCount As Long, dupCount As Long, exCount As Long
    Dim skipA As Long, skipB As Long
    Dim slot As Long, wanted As Long
    Dim rowsA As Collection, rowsB As Collection
    Dim written As Long
    '--------------------------------------------------------------

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Set wsA = SheetOrNothing(wb, SHEET_A)
    Set wsB = SheetOrNothing(wb, SHEET_B)
    If wsA Is Nothing Or wsB Is Nothing Then
        MsgBox "This workbook needs a tab called """ & SHEET_A & """ and one called """ & _
               SHEET_B & """." & vbCrLf & vbCrLf & _
               "Found: " & SheetList(wb) & vbCrLf & vbCrLf & _
               "Rename the tabs, or change SHEET_A / SHEET_B at the top of this macro.", _
               vbExclamation, "Tabs not found"
        Exit Sub
    End If

    hdrRowA = FindHeaderRow(wsA, KEY_A)
    hdrRowB = FindHeaderRow(wsB, KEY_B)
    If hdrRowA = 0 Or hdrRowB = 0 Then
        MsgBox "Could not find a header row in the first " & HEADER_SEARCH_ROWS & " rows." & vbCrLf & vbCrLf & _
               IIf(hdrRowA = 0, SHEET_A & " has no """ & KEY_A & """ header." & vbCrLf, "") & _
               IIf(hdrRowB = 0, SHEET_B & " has no """ & KEY_B & """ header." & vbCrLf, "") & vbCrLf & _
               "Paste the report WITH its header row, or change KEY_A / KEY_B at the top.", _
               vbExclamation, "Header row not found"
        Exit Sub
    End If
    keyColA = FindCol(wsA, hdrRowA, KEY_A)
    keyColB = FindCol(wsB, hdrRowB, KEY_B)
    lastColA = HeaderWidth(wsA, hdrRowA)
    lastColB = HeaderWidth(wsB, hdrRowB)

    '--- work out which columns each amount pair needs -------------
    pairSpec = Split(AMOUNT_PAIRS, "|")
    nPairs = UBound(pairSpec) - LBound(pairSpec) + 1
    ReDim headsA(1 To nPairs): ReDim headsB(1 To nPairs)
    ReDim colsA(1 To nPairs): ReDim colsB(1 To nPairs)

    For p = 1 To nPairs
        halves = Split(pairSpec(p - 1), ":")
        If UBound(halves) < 1 Then
            MsgBox "AMOUNT_PAIRS is malformed near: " & pairSpec(p - 1) & vbCrLf & vbCrLf & _
                   "Each pair needs a colon, like ""Total Product : Taxable Sales"".", _
                   vbExclamation, "Check the configuration"
            Exit Sub
        End If
        headsA(p) = Trim$(halves(0))
        headsB(p) = Trim$(halves(1))
        colsA(p) = ResolveCols(wsA, hdrRowA, headsA(p), missing, SHEET_A)
        colsB(p) = ResolveCols(wsB, hdrRowB, headsB(p), missing, SHEET_B)
    Next p

    If Len(missing) > 0 Then
        MsgBox "These columns were not found:" & vbCrLf & vbCrLf & missing & vbCrLf & _
               "Check the spelling against the report, or edit AMOUNT_PAIRS at the top of this macro.", _
               vbExclamation, "Columns not found"
        Exit Sub
    End If

    '--- go -------------------------------------------------------
    priorCalc = Application.Calculation
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastA = LastDataRow(wsA, keyColA, hdrRowA)
    lastB = LastDataRow(wsB, keyColB, hdrRowB)
    If lastA <= hdrRowA And lastB <= hdrRowB Then
        MsgBox "Both tabs have headers but no data rows underneath them.", vbExclamation
        GoTo CleanExit
    End If

    maxKeys = (lastA - hdrRowA) + (lastB - hdrRowB) + 2
    ReDim keyRaw(1 To maxKeys)
    ReDim cntA(1 To maxKeys): ReDim cntB(1 To maxKeys)
    ReDim seenA(1 To maxKeys): ReDim seenB(1 To maxKeys)
    ReDim sumA(1 To maxKeys, 1 To nPairs): ReDim sumB(1 To maxKeys, 1 To nPairs)
    ReDim grandA(1 To nPairs): ReDim grandB(1 To nPairs)
    ReDim issue(1 To maxKeys): ReDim biggest(1 To maxKeys)

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 0

    ScanReport wsA, hdrRowA, lastA, lastColA, keyColA, colsA, nPairs, _
               dict, keyCount, keyRaw, cntA, seenA, sumA, grandA, skipA
    ScanReport wsB, hdrRowB, lastB, lastColB, keyColB, colsB, nPairs, _
               dict, keyCount, keyRaw, cntB, seenB, sumB, grandB, skipB

    '--- classify -------------------------------------------------
    Dim gap As Currency
    For idx = 1 To keyCount
        If cntA(idx) > 1 Or cntB(idx) > 1 Then dupCount = dupCount + 1

        If seenA(idx) And Not seenB(idx) Then
            issue(idx) = ISSUE_ONLY_A
        ElseIf seenB(idx) And Not seenA(idx) Then
            issue(idx) = ISSUE_ONLY_B
        Else
            biggest(idx) = 0
            For p = 1 To nPairs
                gap = sumA(idx, p) - sumB(idx, p)
                If gap < 0 Then gap = -gap
                If gap > biggest(idx) Then biggest(idx) = gap
            Next p
            If biggest(idx) > MATCH_TOLERANCE Then
                issue(idx) = ISSUE_DIFFER
            Else
                okCount = okCount + 1
            End If
        End If
        If issue(idx) <> 0 Then exCount = exCount + 1
    Next idx

    ReDim orderIdx(1 To IIf(exCount = 0, 1, exCount))
    For wanted = ISSUE_ONLY_A To ISSUE_DIFFER
        For idx = 1 To keyCount
            If issue(idx) = wanted Then
                slot = slot + 1
                orderIdx(slot) = idx
            End If
        Next idx
    Next wanted
    SortDifferencesDesc orderIdx, issue, biggest, exCount

    '--- which source rows sit behind the exceptions ---------------
    Set excKeys = CreateObject("Scripting.Dictionary")
    excKeys.CompareMode = 0
    For slot = 1 To exCount
        excKeys(NormalizeKey(keyRaw(orderIdx(slot)))) = issue(orderIdx(slot))
    Next slot
    Set rowsA = RowsMatching(wsA, hdrRowA, lastA, keyColA, excKeys)
    Set rowsB = RowsMatching(wsB, hdrRowB, lastB, keyColB, excKeys)

    '--- write RECON ----------------------------------------------
    Set outWs = SheetOrNothing(wb, OUT_SHEET)
    If outWs Is Nothing Then
        Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        outWs.Name = OUT_SHEET
    Else
        outWs.Cells.Clear
        outWs.Cells.Interior.Pattern = xlNone
        On Error Resume Next
        outWs.AutoFilterMode = False
        On Error GoTo CleanFail
    End If

    With outWs
        .Range("A1").Value = SHEET_A & "  vs  " & SHEET_B
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        WriteStat outWs, 3, "Run", Format$(Now, "yyyy-mm-dd hh:mm")
        WriteStat outWs, 4, "Matched on", KEY_A & "  /  " & KEY_B
        For p = 1 To nPairs
            spec = spec & IIf(p > 1, ";   ", "") & headsA(p) & "  vs  " & headsB(p)
        Next p
        WriteStat outWs, 5, "Comparing", spec
        WriteStat outWs, 6, "Tolerance", "must agree within " & Format$(MATCH_TOLERANCE, "0.00##")

        .Range("A8").Value = "SUMMARY"
        .Range("A8").Font.Bold = True
        WriteStat outWs, 9, "Rows read", (lastA - hdrRowA) & " on " & SHEET_A & _
                                          ",   " & (lastB - hdrRowB) & " on " & SHEET_B
        WriteStat outWs, 10, "Distinct order numbers", keyCount
        WriteStat outWs, 11, "Orders that agree", okCount
        WriteStat outWs, 12, "Orders only on " & SHEET_A, CountOf(issue, keyCount, ISSUE_ONLY_A)
        WriteStat outWs, 13, "Orders only on " & SHEET_B, CountOf(issue, keyCount, ISSUE_ONLY_B)
        WriteStat outWs, 14, "Orders where the totals differ", CountOf(issue, keyCount, ISSUE_DIFFER)
        WriteStat outWs, 15, "Orders spread over several lines", dupCount
        WriteStat outWs, 16, "Rows skipped (blank or total lines)", skipA & " on " & SHEET_A & _
                                                                    ",   " & skipB & " on " & SHEET_B

        rowOut = 18
        .Cells(rowOut, 1).Value = "GRAND TOTALS"
        .Cells(rowOut, 1).Font.Bold = True
        rowOut = rowOut + 1
        PaintHeader outWs, rowOut, Array("Comparing", SHEET_A, SHEET_B, "Difference")
        rowOut = rowOut + 1
        For p = 1 To nPairs
            .Cells(rowOut, 1).Value = headsA(p) & "  vs  " & headsB(p)
            .Cells(rowOut, 2).Value = grandA(p)
            .Cells(rowOut, 3).Value = grandB(p)
            .Cells(rowOut, 4).Value = grandA(p) - grandB(p)
            .Range(.Cells(rowOut, 2), .Cells(rowOut, 4)).NumberFormat = MoneyFmt()
            If Abs(grandA(p) - grandB(p)) > MATCH_TOLERANCE Then
                .Cells(rowOut, 4).Interior.Color = RGB(255, 199, 206)
            Else
                .Cells(rowOut, 4).Interior.Color = RGB(198, 239, 206)
            End If
            rowOut = rowOut + 1
        Next p

        rowOut = rowOut + 1
        If exCount = 0 Then
            .Cells(rowOut, 1).Value = "EVERY ORDER TIES OUT. Nothing to chase."
            .Cells(rowOut, 1).Font.Bold = True
            .Cells(rowOut, 1).Font.Size = 12
            .Cells(rowOut, 1).Interior.Color = RGB(198, 239, 206)
        Else
            '--- 1. one line per problem order ---------------------
            .Cells(rowOut, 1).Value = "ORDERS THAT DID NOT AGREE"
            .Cells(rowOut, 1).Font.Bold = True
            .Cells(rowOut, 1).Font.Size = 12
            rowOut = rowOut + 1

            Dim head() As String
            Dim nCols As Long
            nCols = 4 + 3 * nPairs
            ReDim head(0 To nCols - 1)
            head(0) = KEY_A
            head(1) = "Issue"
            head(2) = SHEET_A & " lines"
            head(3) = SHEET_B & " lines"
            For p = 1 To nPairs
                head(4 + (p - 1) * 3) = headsA(p) & " (" & SHEET_A & ")"
                head(4 + (p - 1) * 3 + 1) = headsB(p) & " (" & SHEET_B & ")"
                head(4 + (p - 1) * 3 + 2) = "Difference"
            Next p
            PaintHeader outWs, rowOut, head
            .Rows(rowOut).RowHeight = 30
            rowOut = rowOut + 1

            Dim firstBody As Long
            firstBody = rowOut
            For slot = 1 To exCount
                idx = orderIdx(slot)
                .Cells(rowOut, 1).Value = keyRaw(idx)
                .Cells(rowOut, 2).Value = IssueText(issue(idx))
                If cntA(idx) > 0 Then .Cells(rowOut, 3).Value = cntA(idx)
                If cntB(idx) > 0 Then .Cells(rowOut, 4).Value = cntB(idx)
                For p = 1 To nPairs
                    If seenA(idx) Then .Cells(rowOut, 4 + (p - 1) * 3 + 1).Value = sumA(idx, p)
                    If seenB(idx) Then .Cells(rowOut, 4 + (p - 1) * 3 + 2).Value = sumB(idx, p)
                    If seenA(idx) And seenB(idx) Then
                        .Cells(rowOut, 4 + (p - 1) * 3 + 3).Value = sumA(idx, p) - sumB(idx, p)
                    End If
                Next p
                .Range(.Cells(rowOut, 1), .Cells(rowOut, nCols)).Interior.Color = IssueTint(issue(idx))
                If cntA(idx) > 1 Or cntB(idx) > 1 Then
                    .Cells(rowOut, 3).Font.Bold = True
                    .Cells(rowOut, 4).Font.Bold = True
                End If
                rowOut = rowOut + 1
            Next slot
            .Range(.Cells(firstBody, 5), .Cells(rowOut - 1, nCols)).NumberFormat = MoneyFmt()
            .Range(.Cells(firstBody - 1, 1), .Cells(rowOut - 1, nCols)).Borders.Color = RGB(191, 191, 191)

            '--- 2. the full rows, straight off each report --------
            rowOut = rowOut + 2
            written = DumpRows(outWs, rowOut, wsA, hdrRowA, lastColA, keyColA, rowsA, excKeys, _
                               SHEET_A & " - THE FULL ROWS BEHIND THOSE ORDERS")
            rowOut = rowOut + written + 2
            written = DumpRows(outWs, rowOut, wsB, hdrRowB, lastColB, keyColB, rowsB, excKeys, _
                               SHEET_B & " - THE FULL ROWS BEHIND THOSE ORDERS")
        End If

        .Columns("A:A").ColumnWidth = 18
        .Columns("B:B").ColumnWidth = 22
        .Activate
        .Range("A1").Select
    End With

    msg = keyCount & " orders compared." & vbCrLf & vbCrLf & _
          okCount & " agree" & vbCrLf & _
          CountOf(issue, keyCount, ISSUE_ONLY_A) & " only on " & SHEET_A & vbCrLf & _
          CountOf(issue, keyCount, ISSUE_ONLY_B) & " only on " & SHEET_B & vbCrLf & _
          CountOf(issue, keyCount, ISSUE_DIFFER) & " with totals that differ" & vbCrLf & vbCrLf & _
          "Grand total gap: " & Format$(grandA(1) - grandB(1), "#,##0.00")
    If exCount = 0 Then
        msg = msg & vbCrLf & vbCrLf & "Nothing to chase."
    Else
        msg = msg & vbCrLf & vbCrLf & rowsA.Count & " " & SHEET_A & " rows and " & _
              rowsB.Count & " " & SHEET_B & " rows copied to " & OUT_SHEET & "."
    End If

    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, "Report comparison"
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
    MsgBox "Stopped on error " & errNum & ": " & eDesc & vbCrLf & vbCrLf & _
           "Your report tabs were not changed - this macro only reads them.", _
           vbCritical, "Report comparison"
End Sub


'==================================================================
'  Read one report into the shared tallies
'==================================================================
Private Sub ScanReport(ws As Worksheet, ByVal hdrRow As Long, ByVal lastRow As Long, _
                       ByVal lastCol As Long, ByVal keyCol As Long, ByRef amtCols As Variant, _
                       ByVal nPairs As Long, ByRef dict As Object, ByRef keyCount As Long, _
                       ByRef keyRaw() As Variant, ByRef cnt() As Long, ByRef seen() As Boolean, _
                       ByRef total() As Currency, ByRef grand() As Currency, ByRef skipped As Long)

    Dim block As Variant
    Dim i As Long, p As Long, idx As Long, c As Long
    Dim kt As String
    Dim rowSum As Currency

    If lastRow <= hdrRow Then Exit Sub
    block = ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(lastRow, lastCol)).Value

    For i = 1 To UBound(block, 1)
        kt = NormalizeKey(block(i, keyCol))
        If Len(kt) = 0 Then
            skipped = skipped + 1
        ElseIf InStr(1, kt, "TOTAL", vbTextCompare) > 0 Then
            skipped = skipped + 1              ' a Grand Total line pasted with the report
        Else
            If Not dict.Exists(kt) Then
                keyCount = keyCount + 1
                dict.Add kt, keyCount
                keyRaw(keyCount) = block(i, keyCol)
            End If
            idx = dict(kt)
            seen(idx) = True
            cnt(idx) = cnt(idx) + 1

            For p = 1 To nPairs
                rowSum = 0
                For c = LBound(amtCols(p)) To UBound(amtCols(p))
                    If amtCols(p)(c) > 0 Then rowSum = rowSum + SafeCur(block(i, amtCols(p)(c)))
                Next c
                total(idx, p) = total(idx, p) + rowSum
                grand(p) = grand(p) + rowSum
            Next p
        End If
    Next i
End Sub


' Row numbers on ws whose key is one of the exception keys, ascending.
Private Function RowsMatching(ws As Worksheet, ByVal hdrRow As Long, ByVal lastRow As Long, _
                              ByVal keyCol As Long, ByRef excKeys As Object) As Collection
    Dim hits As New Collection
    Dim col As Variant
    Dim i As Long
    Set RowsMatching = hits
    If lastRow <= hdrRow Then Exit Function
    If excKeys.Count = 0 Then Exit Function

    col = ws.Range(ws.Cells(hdrRow + 1, keyCol), ws.Cells(lastRow, keyCol)).Value
    For i = 1 To UBound(col, 1)
        If excKeys.Exists(NormalizeKey(col(i, 1))) Then hits.Add hdrRow + i
    Next i
End Function


' Copies the header row and the listed rows onto the output sheet,
' keeping the report's own formatting. Column A carries the issue.
' Returns how many rows were written, header included.
Private Function DumpRows(outWs As Worksheet, ByVal atRow As Long, src As Worksheet, _
                          ByVal hdrRow As Long, ByVal lastCol As Long, ByVal keyCol As Long, _
                          ByRef hits As Collection, ByRef excKeys As Object, _
                          ByVal caption As String) As Long

    Dim i As Long, batchStart As Long, batchEnd As Long, n As Long
    Dim chunk As Range, oneRow As Range
    Dim destRow As Long
    Dim v As Variant

    outWs.Cells(atRow, 1).Value = caption
    outWs.Cells(atRow, 1).Font.Bold = True
    outWs.Cells(atRow, 1).Font.Size = 12

    outWs.Cells(atRow + 1, 1).Value = "Issue"
    outWs.Cells(atRow + 1, 1).Font.Bold = True
    outWs.Cells(atRow + 1, 1).Font.Color = RGB(255, 255, 255)
    outWs.Cells(atRow + 1, 1).Interior.Color = RGB(68, 84, 106)
    src.Range(src.Cells(hdrRow, 1), src.Cells(hdrRow, lastCol)).Copy _
        Destination:=outWs.Cells(atRow + 1, 2)

    destRow = atRow + 2
    n = hits.Count
    If n = 0 Then
        outWs.Cells(destRow, 1).Value = "(none)"
        DumpRows = 3
        Exit Function
    End If

    ' Copy in batches so a very long exception list does not build one
    ' enormous multi-area range.
    batchStart = 1
    Do While batchStart <= n
        batchEnd = batchStart + 99
        If batchEnd > n Then batchEnd = n
        Set chunk = Nothing
        For i = batchStart To batchEnd
            Set oneRow = src.Range(src.Cells(hits(i), 1), src.Cells(hits(i), lastCol))
            If chunk Is Nothing Then
                Set chunk = oneRow
            Else
                Set chunk = Application.Union(chunk, oneRow)
            End If
        Next i
        On Error Resume Next
        chunk.Copy Destination:=outWs.Cells(destRow + batchStart - 1, 2)
        If Err.Number <> 0 Then
            ' Excel refused the multi-area copy; fall back to one row at a time.
            Err.Clear
            On Error GoTo 0
            For i = batchStart To batchEnd
                src.Range(src.Cells(hits(i), 1), src.Cells(hits(i), lastCol)).Copy _
                    Destination:=outWs.Cells(destRow + i - 1, 2)
            Next i
        End If
        On Error GoTo 0
        batchStart = batchEnd + 1
    Loop
    Application.CutCopyMode = False

    ' Label each copied row with what is wrong with its order.
    For i = 1 To n
        v = outWs.Cells(destRow + i - 1, 1 + keyCol).Value
        outWs.Cells(destRow + i - 1, 1).Value = IssueText(CLng(excKeys(NormalizeKey(v))))
        outWs.Cells(destRow + i - 1, 1).Interior.Color = IssueTint(CLng(excKeys(NormalizeKey(v))))
    Next i

    outWs.Range(outWs.Cells(atRow + 1, 1), outWs.Cells(destRow + n - 1, 1 + lastCol)) _
         .Borders.Color = RGB(191, 191, 191)

    DumpRows = 2 + n
End Function


'==================================================================
'  Helpers
'==================================================================

Private Function IssueText(ByVal code As Long) As String
    Select Case code
        Case ISSUE_ONLY_A: IssueText = "Only on " & SHEET_A
        Case ISSUE_ONLY_B: IssueText = "Only on " & SHEET_B
        Case ISSUE_DIFFER: IssueText = "Totals differ"
    End Select
End Function


Private Function IssueTint(ByVal code As Long) As Long
    If code = ISSUE_DIFFER Then
        IssueTint = RGB(255, 235, 156)
    Else
        IssueTint = RGB(255, 199, 206)
    End If
End Function


Private Function MoneyFmt() As String
    MoneyFmt = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_)"
End Function


Private Sub PaintHeader(ws As Worksheet, ByVal r As Long, ByRef labels As Variant)
    Dim j As Long
    For j = LBound(labels) To UBound(labels)
        With ws.Cells(r, j - LBound(labels) + 1)
            .Value = labels(j)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 84, 106)
            .WrapText = True
            .VerticalAlignment = xlBottom
        End With
    Next j
End Sub


Private Function SheetOrNothing(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), nm, vbTextCompare) = 0 Then
            Set SheetOrNothing = ws
            Exit Function
        End If
    Next ws
End Function


Private Function SheetList(wb As Workbook) As String
    Dim ws As Worksheet, s As String
    For Each ws In wb.Worksheets
        s = s & IIf(Len(s) > 0, ", ", "") & ws.Name
    Next ws
    SheetList = s
End Function


Private Function HeaderWidth(ws As Worksheet, ByVal hdrRow As Long) As Long
    Dim c As Long
    c = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    If c < 1 Then c = 1
    HeaderWidth = c
End Function


Private Function FindHeaderRow(ws As Worksheet, ByVal keyHeader As String) As Long
    Dim r As Long
    For r = 1 To HEADER_SEARCH_ROWS
        If FindCol(ws, r, keyHeader) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function


Private Function FindCol(ws As Worksheet, ByVal hdrRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    Dim want As String
    want = SquashSpaces(headerText)
    If Len(want) = 0 Then Exit Function

    lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function
    For c = 1 To lastCol
        If StrComp(SquashSpaces(ws.Cells(hdrRow, c).Text), want, vbTextCompare) = 0 Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function


' "Taxable Sales + Non Taxable Sales" -> array of column numbers.
' Anything not found is appended to notFound instead of stopping.
Private Function ResolveCols(ws As Worksheet, ByVal hdrRow As Long, ByVal spec As String, _
                             ByRef notFound As String, ByVal sheetNm As String) As Variant
    Dim parts As Variant
    Dim out() As Long
    Dim i As Long, c As Long, keep As Long
    parts = Split(spec, "+")
    ReDim out(0 To UBound(parts))
    keep = -1
    For i = 0 To UBound(parts)
        c = FindCol(ws, hdrRow, Trim$(parts(i)))
        If c = 0 Then
            notFound = notFound & "   " & sheetNm & ":  """ & Trim$(parts(i)) & """" & vbCrLf
        Else
            keep = keep + 1
            out(keep) = c
        End If
    Next i
    If keep < 0 Then
        ReDim out(0 To 0)
        out(0) = 0
    Else
        ReDim Preserve out(0 To keep)
    End If
    ResolveCols = out
End Function


Private Function LastDataRow(ws As Worksheet, ByVal keyCol As Long, ByVal hdrRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If r < hdrRow Then r = hdrRow
    LastDataRow = r
End Function


' A key stored as text on one report and as a number on the other
' has to tie out, so both go through the same grinder.
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
    s = Replace$(s, Chr$(160), "")
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


Private Function SquashSpaces(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, Chr$(160), " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, vbCr, " ")
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    SquashSpaces = Trim$(t)
End Function


' Blank cells and blank-looking strings are zero; anything else that
' is not a number is zero too, and the row still gets counted so the
' line counts stay honest.
Private Function SafeCur(ByVal v As Variant) As Currency
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If Not IsNumeric(v) Then Exit Function
    SafeCur = CCur(v)
End Function


Private Function CountOf(ByRef issue() As Long, ByVal upTo As Long, ByVal wanted As Long) As Long
    Dim i As Long, hits As Long
    For i = 1 To upTo
        If issue(i) = wanted Then hits = hits + 1
    Next i
    CountOf = hits
End Function


' Biggest disagreements first inside the "totals differ" group.
Private Sub SortDifferencesDesc(ByRef orderIdx() As Long, ByRef issue() As Long, _
                                ByRef biggest() As Currency, ByVal total As Long)
    Dim lo As Long, hi As Long, i As Long, j As Long, hold As Long
    If total < 2 Then Exit Sub

    lo = 0
    For i = 1 To total
        If issue(orderIdx(i)) = ISSUE_DIFFER Then
            lo = i
            Exit For
        End If
    Next i
    If lo = 0 Then Exit Sub
    hi = total

    For i = lo + 1 To hi
        hold = orderIdx(i)
        j = i - 1
        Do While j >= lo
            If biggest(orderIdx(j)) >= biggest(hold) Then Exit Do
            orderIdx(j + 1) = orderIdx(j)
            j = j - 1
        Loop
        orderIdx(j + 1) = hold
    Next i
End Sub


Private Sub WriteStat(ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal v As Variant)
    ws.Cells(r, 1).Value = lbl
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = v
    ws.Cells(r, 2).HorizontalAlignment = xlLeft
End Sub
