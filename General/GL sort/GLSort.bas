' Two macros for the month end GL.
'
' SortGL is for the AP GL. It sorts the sheet into groups by keyword.
'   Rows that match nothing stay on top in their original order.
'   Matched rows move below a gap, one block per group, also in original order.
'   Then it sums the block on top, with the net (debit minus credit) under the debit column.
'   Leaves no extra columns behind. Safe to run again on a sheet it already sorted.
'
' ARGLTotals is for the AR GL. No sorting, just the sums,
'   with the net (credit minus debit) under the credit column.
'
' The sums always run from the first row down to the row right above them,
' so cutting rows out of the block can't stretch them.
'
' Both work on whatever sheet is active.
' Paste into an empty module and run one of them with Alt F8.

Private Const MAKE_BACKUP As Boolean = True    ' copy the sheet first, since a macro can't be undone
Private Const SHOW_SUMMARY As Boolean = True   ' popup with the counts at the end
Private Const GAP_ROWS As Long = 10            ' blank rows between the rows that stay and the groups that move
Private Const GROUP_GAP_ROWS As Long = 0       ' blank rows between each moved group, 0 for none
Private Const ADD_TOTALS As Boolean = True     ' after sorting, total the block left on top, AP style
Private Const HEADER_ROW As Long = 0           ' row with the column names, 0 to search the top 20 rows

' Column names the amount work looks for. Spellings separated by |
Private Const DEBIT_HEADER As String = "debit amount|debit"
Private Const CREDIT_HEADER As String = "credit amount|credit"
Private Const VENDOR_HEADER As String = "vendor|vendor name"

' ---- Pairing ----
' An ORD is a prior month invoice reported as paid but not paid yet, and it sits on the credit side.
' When it really gets paid, a normal looking entry shows up on the debit side for the same amount.
' Those payments are the exceptions: they stay in the top block but drop to the bottom of it,
' right above the sums, shaded, in the same order as the ORDs they clear.
' A payment matches an ORD on the vendor and the amount. The bill of lading (BL and the number
' after it, in the description column) then sorts them out:
'   green:  vendor, amount and bill of lading all match
'   yellow: vendor and amount match, and one of the two rows has no bill of lading to compare
'   orange: your call. The amount fits more than one ORD, or the bills of lading don't match
' They sit in that order, green first. Down in the ORD group, the ORDs still open sort first
' and the cleared ones sit shaded at the end.
Private Const PAIR_ORD As Boolean = True       ' False turns all of this off
Private Const PAIR_GROUP As String = "ORD"     ' the group name from GetRules that gets paired
Private Const VENDOR_LETTERS As Long = 4       ' letters of the vendor name that have to match, since names vary by state, 0 for amount alone
Private Const MATCH_BOL As Boolean = True      ' False ignores the bill of lading, the way it used to
Private Const BOL_HEADER As String = "document description|description"
Private Const BOL_DIGITS As Long = 0           ' digits after BL that have to match, 0 for the whole number
Private Const COLOR_ROWS As Boolean = True     ' shade the exceptions and the ORDs they clear
Private Const SUM_SKIPS_SHADED As Boolean = False  ' True leaves the shaded exceptions out of the sums

' Sort ranks. 0 stays on top, groups are 1, 2, 3. These sit between, so exceptions land at the bottom of the top block.
Private Const SURE_RANK As Double = 0.4
Private Const EXCEPTION_RANK As Double = 0.5
Private Const MAYBE_RANK As Double = 0.6

' Pastels, written back to front the way Excel wants them (blue, green, red)
Private Const SURE_COLOR As Long = &HCEEFC6        ' soft green: bill of lading matches too
Private Const EXCEPTION_COLOR As Long = &HCCF2FF   ' soft yellow: an exception, no bill of lading to check it by
Private Const MAYBE_COLOR As Long = &HD6E4FC       ' soft orange: your call
Private Const ORD_MATCHED_COLOR As Long = &HF7EBDD ' soft blue: the ORD an exception clears

' ---- Rules: column header | keyword | group name ----
' First matching rule wins. Groups land in this order.
' To add one, copy a line and change it.
' Header names ignore caps, spaces, line breaks and punctuation, so "AP Document" + new line + "Number" matches "ap document number".
' List a few spellings with | between them, or skip names and use "column H".
Private Function GetRules() As Variant
    GetRules = Array( _
        Array("document description|description", "MISC", "MISC"), _
        Array("ap document number|document number", "ORD", "ORD"), _
        Array("document description|description", "BANK", "BANK") _
    )
End Function

Public Sub SortGL()
    Dim ws As Worksheet, ruleList As Variant, r As Long, i As Long, c As Long, k As Long
    Dim headerRow As Long, dataRow As Long, lastRow As Long, lastCol As Long, nRows As Long
    Dim ruleCol() As Long, allData As Variant, v As Variant, cellText As String
    Dim rankArr() As Variant, origArr() As Variant, pairArr() As Variant, isBlank As Boolean
    Dim counts() As Long, unmatched As Long, blanks As Long, moved As Long, above As Long
    Dim ordRank As Long, paired As Long, sure As Long, ambiguous As Long, topEnd As Long, ordStart As Long
    Dim debCol As Long, credCol As Long, vendCol As Long, bolCol As Long
    Dim msg As String, totalsMsg As String, oldCalc As XlCalculation

    Set ws = ActiveSheet
    ruleList = GetRules()

    ' Header row is HEADER_ROW, or the first row in the top 20 with every column the rules need
    headerRow = FindHeaderRow(ws, ruleList)
    If headerRow = 0 Then
        MsgBox HeaderHelp(ws, ruleList), vbExclamation
        Exit Sub
    End If

    If ws.FilterMode Then ws.ShowAllData

    ' Last row and column with anything in them, so no stray column gets left behind
    lastRow = ws.Cells.Find("*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    dataRow = FirstDataRow(ws, headerRow, lastCol)

    nRows = lastRow - dataRow + 1
    If nRows < 2 Then
        MsgBox "Not enough rows under the header to sort.", vbExclamation
        Exit Sub
    End If

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    If MAKE_BACKUP Then
        ws.Copy After:=ws
        On Error Resume Next
        ActiveSheet.Name = Left$(ws.Name, 19) & " before sort"
        On Error GoTo CleanUp
        ws.Activate
    End If

    ReDim ruleCol(0 To UBound(ruleList))
    For r = 0 To UBound(ruleList)
        ruleCol(r) = FindHeaderCol(ws, headerRow, lastCol, CStr(ruleList(r)(0)))
    Next r

    ' Read the whole block at once
    allData = ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, lastCol)).Value

    ReDim rankArr(1 To nRows, 1 To 1)
    ReDim origArr(1 To nRows, 1 To 1)
    ReDim pairArr(1 To nRows, 1 To 1)
    ReDim counts(0 To UBound(ruleList))

    For i = 1 To nRows
        origArr(i, 1) = i
        rankArr(i, 1) = 0
        pairArr(i, 1) = 0

        ' Empty rows (like an old gap) sink to the very bottom
        isBlank = True
        For c = 1 To lastCol
            v = allData(i, c)
            If IsError(v) Then
                isBlank = False
            ElseIf Len(CStr(v)) > 0 Then
                isBlank = False
            End If
            If Not isBlank Then Exit For
        Next c

        If isBlank Then
            rankArr(i, 1) = UBound(ruleList) + 2
            blanks = blanks + 1
        Else
            For r = 0 To UBound(ruleList)
                v = allData(i, ruleCol(r))
                If IsError(v) Then cellText = "" Else cellText = CStr(v)
                If InStr(1, cellText, ruleList(r)(1), vbTextCompare) > 0 Then
                    rankArr(i, 1) = r + 1
                    counts(r) = counts(r) + 1
                    Exit For
                End If
            Next r
            If rankArr(i, 1) = 0 Then unmatched = unmatched + 1
        End If
    Next i

    ' Pull each paid ORD's payment off the top and park it with its ORD
    ordRank = 0
    If PAIR_ORD Then
        For r = 0 To UBound(ruleList)
            If StrComp(CStr(ruleList(r)(2)), PAIR_GROUP, vbTextCompare) = 0 Then ordRank = r + 1
        Next r
        debCol = FindHeaderCol(ws, headerRow, lastCol, DEBIT_HEADER)
        credCol = FindHeaderCol(ws, headerRow, lastCol, CREDIT_HEADER)
        vendCol = FindHeaderCol(ws, headerRow, lastCol, VENDOR_HEADER)
        If MATCH_BOL Then bolCol = FindHeaderCol(ws, headerRow, lastCol, BOL_HEADER)
        If ordRank > 0 And debCol > 0 And credCol > 0 Then
            paired = PairOrdRows(allData, rankArr, pairArr, nRows, ordRank, debCol, credCol, vendCol, bolCol, sure, ambiguous)
        End If
    End If

    moved = nRows - unmatched - blanks

    ' Three temporary columns to sort by: group, then pair, then original row
    ws.Range(ws.Cells(dataRow, lastCol + 1), ws.Cells(lastRow, lastCol + 1)).Value = rankArr
    ws.Range(ws.Cells(dataRow, lastCol + 2), ws.Cells(lastRow, lastCol + 2)).Value = pairArr
    ws.Range(ws.Cells(dataRow, lastCol + 3), ws.Cells(lastRow, lastCol + 3)).Value = origArr

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(dataRow, lastCol + 1), ws.Cells(lastRow, lastCol + 1)), Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(dataRow, lastCol + 2), ws.Cells(lastRow, lastCol + 2)), Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(dataRow, lastCol + 3), ws.Cells(lastRow, lastCol + 3)), Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, lastCol + 3))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With

    ' Temporary columns gone
    ws.Range(ws.Cells(dataRow, lastCol + 1), ws.Cells(lastRow, lastCol + 3)).Clear

    ' Old shading off, so running it twice doesn't leave color on the wrong rows
    If COLOR_ROWS Then ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, lastCol)).Interior.Pattern = xlNone

    ' Gaps between groups first, from the bottom up so row numbers above don't move
    If GROUP_GAP_ROWS > 0 Then
        For r = UBound(ruleList) To 1 Step -1
            above = 0
            For k = 0 To r - 1
                above = above + counts(k)
            Next k
            If counts(r) > 0 And above > 0 Then
                InsertBlankRows ws, dataRow + unmatched + above, GROUP_GAP_ROWS, lastCol
            End If
        Next r
    End If

    ' Then the big gap between the rows that stayed and everything that moved
    If GAP_ROWS > 0 And unmatched > 0 And moved > 0 Then
        InsertBlankRows ws, dataRow + unmatched, GAP_ROWS, lastCol
    End If

    ' Shade the exceptions sitting at the bottom of the top block, and the ORDs they clear
    If COLOR_ROWS And paired + ambiguous > 0 Then
        topEnd = dataRow + unmatched - 1
        ShadeRows ws, topEnd - paired - ambiguous + 1, sure, lastCol, SURE_COLOR
        ShadeRows ws, topEnd - paired - ambiguous + 1 + sure, paired - sure, lastCol, EXCEPTION_COLOR
        ShadeRows ws, topEnd - ambiguous + 1, ambiguous, lastCol, MAYBE_COLOR
        If paired > 0 Then
            ordStart = GroupStartRow(counts, ordRank, dataRow + unmatched + GAP_ROWS)
            ShadeRows ws, ordStart + counts(ordRank - 1) - paired, paired, lastCol, ORD_MATCHED_COLOR
        End If
    End If

    If ADD_TOTALS Then totalsMsg = DoTotals(ws, True)

    msg = "Done. Left on top: " & Format$(unmatched - paired - ambiguous, "#,##0") & vbCrLf
    For r = 0 To UBound(ruleList)
        msg = msg & ruleList(r)(2) & ": " & Format$(counts(r), "#,##0")
        If r + 1 = ordRank And paired > 0 Then msg = msg & " (" & Format$(paired, "#,##0") & " of them cleared, shaded blue at the end of the group)"
        msg = msg & vbCrLf
    Next r
    If ordRank > 0 Then
        msg = msg & vbCrLf & "Exceptions found: " & Format$(paired, "#,##0") & _
              " (bottom of the top block, in the same order as the ORDs they clear)"
        If sure > 0 Then msg = msg & vbCrLf & "    Green: " & Format$(sure, "#,##0") & _
              " (vendor, amount and bill of lading all match)"
        If paired > sure Then msg = msg & vbCrLf & "    Yellow: " & Format$(paired - sure, "#,##0") & _
              " (vendor and amount match, no bill of lading to check)"
        If ambiguous > 0 Then msg = msg & vbCrLf & "Your call: " & Format$(ambiguous, "#,##0") & _
              " (under them in orange: the amount fit more than one ORD, or the bills of lading don't match)"
        msg = msg & vbCrLf & "ORDs still open: " & Format$(counts(ordRank - 1) - paired, "#,##0")
    End If
    If GAP_ROWS > 0 And unmatched > 0 And moved > 0 Then
        msg = msg & vbCrLf & "Moved rows start at row " & Format$(dataRow + unmatched + GAP_ROWS, "#,##0") & "."
    End If
    If Len(totalsMsg) > 0 Then msg = msg & vbCrLf & vbCrLf & totalsMsg
    If MAKE_BACKUP Then msg = msg & vbCrLf & vbCrLf & "Original is on the sheet next to this one."

CleanUp:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Stopped with an error: " & Err.Description, vbCritical
    ElseIf SHOW_SUMMARY Then
        MsgBox msg, vbInformation
    End If
End Sub

' ---- Pairing ----
' Looks for the payment that cleared an ORD: same amount, other side, vendor name starting the same.
' If both rows have a bill of lading, it has to match too, and it picks between look-alike ORDs.
' A payment that pairs sits at the bottom of the top block, and its ORD goes blue.
' Anything it can't settle goes orange for a human, and its ORD stays open.
' Hands back how many paired, plus how many of those were green and how many went orange.
Private Function PairOrdRows(allData As Variant, rankArr As Variant, pairArr As Variant, ByVal nRows As Long, _
                             ByVal ordRank As Long, ByVal debCol As Long, ByVal credCol As Long, ByVal vendCol As Long, _
                             ByVal bolCol As Long, ByRef sure As Long, ByRef ambiguous As Long) As Long
    Dim i As Long, n As Long, amt As Double, k As String, hit As String, rest As String
    Dim cands As Variant, bol As String, pick As Long, bolHits As Long, pairNo As Long
    Dim ordByKey As Collection
    Set ordByKey = New Collection

    ' Every ORD credit filed under vendor letters plus amount. Look-alikes share one entry, like "12|40|57"
    For i = 1 To nRows
        If rankArr(i, 1) = ordRank Then
            amt = RowNet(allData, i, debCol, credCol)
            If amt < 0 Then
                k = MatchKey(allData, i, vendCol, -amt)
                hit = KeyEntry(ordByKey, k)
                If Len(hit) > 0 Then
                    ordByKey.Remove k
                    hit = hit & "|"
                End If
                ordByKey.Add hit & CStr(i), k
            End If
        End If
    Next i

    ' Then the debits still sitting on top
    For i = 1 To nRows
        If rankArr(i, 1) = 0 Then
            amt = RowNet(allData, i, debCol, credCol)
            If amt > 0 Then
                k = MatchKey(allData, i, vendCol, amt)
                hit = KeyEntry(ordByKey, k)
                If Len(hit) > 0 Then
                    cands = Split(hit, "|")
                    pick = 0
                    bolHits = 0

                    ' The bill of lading picks the ORD, when the payment has one
                    bol = BolKey(allData, i, bolCol)
                    If Len(bol) > 0 Then
                        For n = 0 To UBound(cands)
                            If BolKey(allData, CLng(cands(n)), bolCol) = bol Then
                                bolHits = bolHits + 1
                                pick = CLng(cands(n))
                            End If
                        Next n
                        If bolHits <> 1 Then pick = 0
                    End If

                    ' No bill of lading to go by: fine if only one ORD fits and nothing disagrees
                    If pick = 0 And bolHits = 0 And UBound(cands) = 0 Then
                        If Len(bol) = 0 Or Len(BolKey(allData, CLng(cands(0)), bolCol)) = 0 Then pick = CLng(cands(0))
                    End If

                    If pick > 0 Then
                        ' Green pairs number first and yellow after, so the ORDs line up in the same order
                        pairNo = pairNo + 1
                        If bolHits = 1 Then
                            sure = sure + 1
                            rankArr(i, 1) = SURE_RANK
                            pairArr(i, 1) = pairNo
                        Else
                            rankArr(i, 1) = EXCEPTION_RANK
                            pairArr(i, 1) = nRows + pairNo
                        End If
                        pairArr(pick, 1) = pairArr(i, 1)
                        PairOrdRows = PairOrdRows + 1

                        ' That ORD is spoken for
                        rest = ""
                        For n = 0 To UBound(cands)
                            If CLng(cands(n)) <> pick Then rest = rest & "|" & cands(n)
                        Next n
                        ordByKey.Remove k
                        If Len(rest) > 0 Then ordByKey.Add Mid$(rest, 2), k
                    Else
                        ' More than one ORD it could be, or the bills of lading disagree
                        rankArr(i, 1) = MAYBE_RANK
                        ambiguous = ambiguous + 1
                    End If
                End If
            End If
        End If
    Next i
End Function

' What the list holds under a key, "" when nothing does
Private Function KeyEntry(col As Collection, ByVal k As String) As String
    On Error Resume Next
    KeyEntry = col.Item(k)
    On Error GoTo 0
End Function

' "BL" plus the number after it, found anywhere in the cell, so "BL7831631 FREIGHT" gives BL7831631.
' A space, dash, colon or # between BL and the number is fine. "" when there isn't one.
Private Function BolKey(allData As Variant, ByVal i As Long, ByVal bolCol As Long) As String
    Dim s As String, p As Long, q As Long, d As String
    If bolCol = 0 Then Exit Function
    If IsError(allData(i, bolCol)) Then Exit Function
    s = UCase$(CStr(allData(i, bolCol)))

    p = InStr(1, s, "BL")
    Do While p > 0
        q = p + 2
        Do While q <= Len(s)
            If InStr(" -:#", Mid$(s, q, 1)) = 0 Then Exit Do
            q = q + 1
        Loop
        If Mid$(s, q, 1) Like "#" Then Exit Do
        p = InStr(p + 1, s, "BL")
    Loop
    If p = 0 Then Exit Function

    Do While Mid$(s, q, 1) Like "#"
        d = d & Mid$(s, q, 1)
        q = q + 1
    Loop
    If BOL_DIGITS > 0 Then d = Left$(d, BOL_DIGITS)
    BolKey = "BL" & d
End Function

' Debit minus credit, so a credit comes back negative
Private Function RowNet(allData As Variant, ByVal i As Long, ByVal debCol As Long, ByVal credCol As Long) As Double
    Dim d As Double, c As Double
    If Not IsError(allData(i, debCol)) Then
        If IsNumeric(allData(i, debCol)) Then d = CDbl(allData(i, debCol))
    End If
    If Not IsError(allData(i, credCol)) Then
        If IsNumeric(allData(i, credCol)) Then c = CDbl(allData(i, credCol))
    End If
    RowNet = d - c
End Function

' First few letters of the vendor name plus the amount to the cent
Private Function MatchKey(allData As Variant, ByVal i As Long, ByVal vendCol As Long, ByVal amt As Double) As String
    Dim v As String
    If vendCol > 0 Then
        If Not IsError(allData(i, vendCol)) Then v = Squash(CStr(allData(i, vendCol)))
    End If
    If VENDOR_LETTERS > 0 Then v = Left$(v, VENDOR_LETTERS)
    MatchKey = v & "~" & Format$(amt, "0.00")
End Function

' ---- Totals ----
' Sum under the debit and credit columns, then the net two rows under the sum.
' AR nets credit minus debit and sits in the credit column.
' AP nets debit minus credit and sits in the debit column.
' On a sheet SortGL has already split, only the block on top gets summed.
Public Sub ARGLTotals()
    Dim summary As String
    summary = DoTotals(ActiveSheet, False)
    If SHOW_SUMMARY And Len(summary) > 0 Then MsgBox summary, vbInformation
End Sub

' Does the work and hands back what it did, so SortGL can put it in its own popup
Private Function DoTotals(ws As Worksheet, ByVal isAP As Boolean) As String
    Dim headerRow As Long, dataRow As Long, lastRow As Long, lastCol As Long
    Dim debCol As Long, credCol As Long, blockEnd As Long, sumEnd As Long, sumRow As Long, diffRow As Long
    Dim kind As String, netCell As String, netVal As Double, c As Variant

    lastRow = ws.Cells.Find("*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    headerRow = FindAmountHeaderRow(ws, lastCol)
    If headerRow = 0 Then
        MsgBox "Couldn't find a Debit and a Credit column in the top 20 rows.", vbExclamation
        Exit Function
    End If
    debCol = FindHeaderCol(ws, headerRow, lastCol, DEBIT_HEADER)
    credCol = FindHeaderCol(ws, headerRow, lastCol, CREDIT_HEADER)

    dataRow = FirstDataRow(ws, headerRow, lastCol)
    blockEnd = BlockEndRow(ws, dataRow, lastRow, lastCol, debCol, credCol, sumEnd)
    If blockEnd < dataRow Then
        MsgBox "No rows to total.", vbExclamation
        Exit Function
    End If

    kind = IIf(isAP, "AP", "AR")
    sumRow = blockEnd + 1
    diffRow = sumRow + 2

    For Each c In Array(debCol, credCol)
        If sumEnd < blockEnd Then
            ' Skipping the shaded rows means a fixed end, on the last row that isn't shaded
            ws.Cells(sumRow, c).Formula = "=SUM(" & ColLetter(c) & dataRow & ":" & ColLetter(c) & sumEnd & ")"
        Else
            ' Ends on the row right above the sum, wherever that is, so cutting rows out can't stretch it
            ws.Cells(sumRow, c).Formula = "=SUM(" & ColLetter(c) & dataRow & ":INDEX(" & ColLetter(c) & ":" & ColLetter(c) & ",ROW()-1))"
        End If
        ws.Cells(sumRow, c).NumberFormat = ws.Cells(blockEnd, c).NumberFormat
    Next c

    If isAP Then
        ws.Cells(diffRow, debCol).Formula = "=" & ColLetter(debCol) & sumRow & "-" & ColLetter(credCol) & sumRow
        ws.Cells(diffRow, debCol).NumberFormat = ws.Cells(blockEnd, debCol).NumberFormat
        netCell = ColLetter(debCol) & diffRow
    Else
        ws.Cells(diffRow, credCol).Formula = "=" & ColLetter(credCol) & sumRow & "-" & ColLetter(debCol) & sumRow
        ws.Cells(diffRow, credCol).NumberFormat = ws.Cells(blockEnd, credCol).NumberFormat
        netCell = ColLetter(credCol) & diffRow
    End If

    ws.Calculate
    netVal = ws.Range(netCell).Value

    DoTotals = kind & " ledger." & vbCrLf & _
               "Summed rows " & Format$(dataRow, "#,##0") & " to " & Format$(sumEnd, "#,##0") & _
               " (" & Format$(sumEnd - dataRow + 1, "#,##0") & " rows) on row " & Format$(sumRow, "#,##0") & "." & vbCrLf
    If sumEnd < blockEnd Then DoTotals = DoTotals & "The " & Format$(blockEnd - sumEnd, "#,##0") & " shaded rows above the sums are left out of it." & vbCrLf
    DoTotals = DoTotals & "Net is in cell " & netCell & ": " & Format$(netVal, "#,##0.00")
End Function

' First row in the top 20 that has both amount columns
Private Function FindAmountHeaderRow(ws As Worksheet, ByVal lastCol As Long) As Long
    Dim rowNum As Long, wide As Long
    For rowNum = 1 To 20
        wide = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
        If wide < lastCol Then wide = lastCol
        If FindHeaderCol(ws, rowNum, wide, DEBIT_HEADER) > 0 And FindHeaderCol(ws, rowNum, wide, CREDIT_HEADER) > 0 Then
            FindAmountHeaderRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

' Data starts under the header, even when the header cells are merged two rows tall
Private Function FirstDataRow(ws As Worksheet, ByVal headerRow As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    FirstDataRow = headerRow + 1
    For c = 1 To lastCol
        With ws.Cells(headerRow, c).MergeArea
            If .Row + .Rows.Count > FirstDataRow Then FirstDataRow = .Row + .Rows.Count
        End With
    Next c
End Function

' Last row of the first solid block: stops at a blank row, at a gap, or at totals from an earlier run
Private Function BlockEndRow(ws As Worksheet, ByVal dataRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByVal debCol As Long, ByVal credCol As Long, ByRef cleanEnd As Long) As Long
    Dim r As Long, c As Long, filled As Boolean, vals As Variant, forms As Variant, shade As Long
    vals = ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, lastCol)).Value
    forms = ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, lastCol)).Formula
    If Not IsArray(vals) Then
        BlockEndRow = dataRow
        cleanEnd = dataRow
        Exit Function
    End If
    cleanEnd = 0
    For r = 1 To lastRow - dataRow + 1
        ' A sum from an earlier run ends the block
        If Left$(CStr(forms(r, debCol)), 1) = "=" Or Left$(CStr(forms(r, credCol)), 1) = "=" Then Exit For
        filled = False
        For c = 1 To lastCol
            If IsError(vals(r, c)) Then
                filled = True
            ElseIf Len(CStr(vals(r, c))) > 0 Then
                filled = True
            End If
            If filled Then Exit For
        Next c
        If Not filled Then Exit For
        BlockEndRow = dataRow + r - 1
        ' Only if SUM_SKIPS_SHADED: stop the sum where the shaded exceptions start
        If SUM_SKIPS_SHADED And COLOR_ROWS And cleanEnd = 0 Then
            shade = ws.Cells(dataRow + r - 1, 1).Interior.Color
            If shade = SURE_COLOR Or shade = EXCEPTION_COLOR Or shade = MAYBE_COLOR Then cleanEnd = dataRow + r - 2
        End If
    Next r
    If cleanEnd < dataRow Then cleanEnd = BlockEndRow
End Function

' Column number to letter, so the formulas read like ones typed by hand
Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

' Paints a block of rows, only across the data columns
Private Sub ShadeRows(ByVal ws As Worksheet, ByVal atRow As Long, ByVal howMany As Long, ByVal lastCol As Long, ByVal fillColor As Long)
    If howMany < 1 Then Exit Sub
    With ws.Range(ws.Cells(atRow, 1), ws.Cells(atRow + howMany - 1, lastCol)).Interior
        .Pattern = xlSolid
        .Color = fillColor
    End With
End Sub

' Where a group starts once everything has moved, counting the groups above it and any gaps between them
Private Function GroupStartRow(counts() As Long, ByVal groupRank As Long, ByVal firstMovedRow As Long) As Long
    Dim r As Long, k As Long, above As Long, gaps As Long, before As Long
    For r = 0 To groupRank - 2
        above = above + counts(r)
    Next r
    If GROUP_GAP_ROWS > 0 Then
        For r = 1 To groupRank - 1
            before = 0
            For k = 0 To r - 1
                before = before + counts(k)
            Next k
            If counts(r) > 0 And before > 0 Then gaps = gaps + GROUP_GAP_ROWS
        Next r
    End If
    GroupStartRow = firstMovedRow + above + gaps
End Function

' Pushes the data down to open blank rows, only across the data columns
Private Sub InsertBlankRows(ByVal ws As Worksheet, ByVal atRow As Long, ByVal howMany As Long, ByVal lastCol As Long)
    ws.Range(ws.Cells(atRow, 1), ws.Cells(atRow + howMany - 1, lastCol)).Insert Shift:=xlDown
    ws.Range(ws.Cells(atRow, 1), ws.Cells(atRow + howMany - 1, lastCol)).ClearFormats
End Sub

Private Function FindHeaderRow(ws As Worksheet, ruleList As Variant) As Long
    Dim rowNum As Long
    If HEADER_ROW > 0 Then
        If RowHasAllColumns(ws, HEADER_ROW, ruleList) Then FindHeaderRow = HEADER_ROW
        Exit Function
    End If
    For rowNum = 1 To 20
        If RowHasAllColumns(ws, rowNum, ruleList) Then
            FindHeaderRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function RowHasAllColumns(ws As Worksheet, ByVal rowNum As Long, ruleList As Variant) As Boolean
    Dim r As Long, lastCol As Long
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    For r = 0 To UBound(ruleList)
        If FindHeaderCol(ws, rowNum, lastCol, CStr(ruleList(r)(0))) = 0 Then Exit Function
    Next r
    RowHasAllColumns = True
End Function

' Finds a column by any of its spellings, or takes a straight "column H"
Private Function FindHeaderCol(ws As Worksheet, ByVal headerRow As Long, ByVal lastCol As Long, ByVal headerSpec As String) As Long
    Dim c As Long, n As Long, names As Variant, cellName As String
    headerSpec = Trim$(headerSpec)
    If LCase$(Left$(headerSpec, 7)) = "column " Then
        On Error Resume Next
        FindHeaderCol = ws.Range(Trim$(Mid$(headerSpec, 8)) & "1").Column
        On Error GoTo 0
        Exit Function
    End If
    names = Split(headerSpec, "|")
    For c = 1 To lastCol
        cellName = Squash(CStr(ws.Cells(headerRow, c).Text))
        If Len(cellName) > 0 Then
            For n = 0 To UBound(names)
                If cellName = Squash(CStr(names(n))) Then
                    FindHeaderCol = c
                    Exit Function
                End If
            Next n
        End If
    Next c
End Function

' Lowercase letters and digits only, so caps, spaces and punctuation don't matter
Private Function Squash(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then out = out & ch
    Next i
    Squash = out
End Function

' When a column can't be found, say which, and list what the header row really has
Private Function HeaderHelp(ws As Worksheet, ruleList As Variant) As String
    Dim rowNum As Long, lastCol As Long, c As Long, r As Long, best As Long, filled As Long
    Dim missing As String, found As String, txt As String
    If HEADER_ROW > 0 Then
        rowNum = HEADER_ROW
    Else
        For r = 1 To 20
            filled = Application.WorksheetFunction.CountA(ws.Rows(r))
            If filled > best Then best = filled: rowNum = r
        Next r
        If rowNum = 0 Then rowNum = 1
    End If
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    For r = 0 To UBound(ruleList)
        If FindHeaderCol(ws, rowNum, lastCol, CStr(ruleList(r)(0))) = 0 Then
            If InStr(1, missing, "[" & ruleList(r)(0) & "]", vbTextCompare) = 0 Then missing = missing & "[" & ruleList(r)(0) & "] "
        End If
    Next r
    For c = 1 To lastCol
        txt = Trim$(CStr(ws.Cells(rowNum, c).Text))
        If Len(txt) > 0 Then found = found & Split(ws.Cells(1, c).Address(True, False), "$")(0) & " = " & txt & vbCrLf
    Next c
    If Len(found) = 0 Then found = "(nothing)" & vbCrLf
    HeaderHelp = "Couldn't find " & Trim$(missing) & " in row " & rowNum & "." & vbCrLf & vbCrLf & _
                 "Row " & rowNum & " has:" & vbCrLf & found & vbCrLf & _
                 "Fix the header names in GetRules at the top, or set HEADER_ROW."
End Function
