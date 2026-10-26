' New Mexico QFR rows: turns the NON-EXCISE purchase report for New Mexico into rows for
' NMED's quarterly fuel report upload template (CTFP-ARCS, Liquids tab).
'
' Setting it up, once:
'   1. Copy all of this text (click Raw, then Ctrl A, Ctrl C).
'   2. Open a new blank workbook in Excel. Alt F11, then Insert > Module. If the new window
'      already has a line in it, Ctrl A and Delete first. Paste, then close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm). Name it NM QFR Builder.
'   4. Alt F8, pick BuildQFRRows, Run. The first run only makes the Report sheet.
'
' Each quarter:
'   1. Run the purchase report for New Mexico, tax type NON-EXCISE, for the quarter.
'   2. Paste the whole report onto the Report sheet at A1 (clear the old one first).
'   3. Alt F8, BuildQFRRows, Run. It makes two sheets:
'      Upload: the rows for NMED's template, in the template's column order.
'      CHECK: every report row and what happened to it, left-out rows first, with why.
'   4. Read CHECK. Then in CTFP-ARCS, Liquids tab, Download Template. Copy Upload from A2
'      down to the last row, column P, and in the template click the first blank row under
'      the headers in column A and paste as Values (Home, Paste, Values), so the template
'      keeps its own dropdowns and checks. Row Status should say each row is fine.
'   5. Save the template and Bulk Upload it in CTFP-ARCS.
'
' The settings below are the only lines to change when a code or a rule changes.

Option Explicit

' ---- settings -------------------------------------------------------------------------

' Ethanol share of an E10 load, in percent. The rest is reported as gasoline.
Private Const ETH_PCT As Double = 10

' Products by code, as the report spells them, between bars.
Private Const DIESEL_PRODUCTS As String = "|NO2ULS|NO2ULCNRT|"
Private Const E10_PRODUCTS As String = "|86E10|"
Private Const DYED_PRODUCTS As String = "|NO2ULR|"

' False leaves dyed loads out (listed on CHECK). True reports each dyed import as two rows,
' the import and NMED's "Exempt fuel - Dyed fuel" line that takes it back off.
Private Const REPORT_DYED As Boolean = False

' Fuel name, pathway code and vehicle application code, as NMED's lists spell them.
Private Const DIESEL_NAME As String = "Diesel"
Private Const DIESEL_CODE As String = "NMULSD001"
Private Const DIESEL_VEH As String = "ULS-MHDV/Off-Road-ICEV"
Private Const GAS_NAME As String = "Clear Gasoline"
Private Const GAS_CODE As String = "NMGAS001"
Private Const GAS_VEH As String = "GAS-LMDV-ICEV"
Private Const ETH_NAME As String = "Ethanol"
Private Const ETH_CODE As String = "NMETOH001"
Private Const ETH_VEH As String = "ETH-LMDV-ICEV"

' Transaction types. The " - " in the dyed one is written as NMED's long dash.
Private Const TT_IMPORT As String = "Imported outside the bulk system"
Private Const TT_DYED As String = "Exempt fuel - Dyed fuel"

' Quarter as written in the Reporting Quarter column: "Q" makes Q2. Use "" if the
' template's dropdown shows just 2.
Private Const QTR_PREFIX As String = "Q"

' Description: this text, then the BOL number.
Private Const DESC_PREFIX As String = "BOL "

' True puts the report's Supplier Terminal Name in Business Partner Name. False leaves it blank.
Private Const PARTNER_IS_TERMINAL As Boolean = True

' Blended Fuel? for the two halves of an E10 load, and for everything else.
Private Const BLEND_E10 As String = "Yes"
Private Const BLEND_OTHER As String = "No"

' Aggregator Indicator and Aggregation Indicator.
Private Const IND_NO As String = "No"

Private Const HOME_STATE As String = "NM"

' ---- the rest -------------------------------------------------------------------------

Private Const SH_REPORT As String = "Report"
Private Const SH_UP As String = "Upload"
Private Const SH_CHECK As String = "CHECK"
Private Const BOX_TITLE As String = "NM QFR rows"
Private Const UP_HEADS As String = "Reporting Year *|Reporting Quarter *|Transaction Date *|" & _
    "Fuel Name *|Fuel Pathway Code (FPC) *|Fuel Vehicle Application Code *|Transaction Type *|" & _
    "Volume/Quantity *|Business Partner Name|Business Partner FEIN|Aggregator Indicator *|" & _
    "Designator FEIN|Blended Fuel? *|EPA Facility ID|Aggregation Indicator *|Transaction Description"
Private Const UP_COLS As Long = 16
Private Const CHECK_HEADS As String = "Report row|BOL|Loading date|Terminal state|Destination|" & _
    "Product|Gallons|Dollars|Kept?|What happened"
Private Const CHECK_COLS As Long = 10
Private Const CHECK_TOP As Long = 6     ' header row of the table on CHECK

' The report's columns, found by name.
Private Const COL_NAMES As String = "BOL|Loading Date|Supplier Terminal Name|Terminal State|" & _
    "Destination State|Product|Purchased Gallons|Dollars"
Private Const K_BOL As Long = 0, K_DATE As Long = 1, K_TERM As Long = 2, K_TSTATE As Long = 3
Private Const K_DEST As Long = 4, K_PROD As Long = 5, K_GAL As Long = 6, K_DOL As Long = 7

Private mQuiet As Boolean       ' test runs only: no boxes
Private mSaid As String         ' test runs only: what the boxes would have said

' One report row as read.
Private rRow() As Long, rBol() As String, rDate() As Variant, rTerm() As String
Private rTState() As String, rDest() As String, rProd() As String, rGal() As Variant
Private rDol() As Variant, rKept() As Boolean, rWhy() As String
Private rN As Long


Public Sub BuildQFRRows()
    ' Reads the Report sheet and writes the Upload and CHECK sheets.
    On Error GoTo Oops
    Build
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Something went wrong: " & Err.Description & vbLf & vbLf & _
        "Nothing was sent anywhere. Fix the Report sheet or the settings and run it again."
End Sub

Public Function QfrTest() As String
    ' For test runs: runs the build with no boxes and returns what they would say.
    mQuiet = True
    mSaid = ""
    BuildQFRRows
    mQuiet = False
    QfrTest = mSaid
End Function

Private Sub Say(ByVal msg As String)
    If mQuiet Then
        mSaid = mSaid & msg & vbLf
    Else
        MsgBox msg, vbOKOnly, BOX_TITLE
    End If
End Sub

Private Sub Build()
    Dim wb As Workbook, ws As Worksheet
    Dim hdr As Long, cols(7) As Long, missing As String, lastRow As Long, blanks As Long

    Erase rRow, rBol, rDate, rTerm, rTState, rDest, rProd, rGal, rDol, rKept, rWhy
    rN = 0
    Set wb = ThisWorkbook

    Set ws = FindSheet(wb, SH_REPORT)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = SH_REPORT
        Say "Made the Report sheet. Paste the NON-EXCISE purchase report for New Mexico " & _
            "there at A1, then run BuildQFRRows again."
        Exit Sub
    End If

    hdr = FindHeaderRow(ws, cols, missing)
    If hdr = 0 Then
        Say "Couldn't find the report's header row on the Report sheet (looked in the first " & _
            "20 rows for BOL and Destination State)." & vbLf & vbLf & _
            "Paste the whole report at A1, headers included."
        Exit Sub
    End If
    If missing <> "" Then
        Say "The report on the Report sheet is missing these columns: " & missing & vbLf & vbLf & _
            "They're found by name, so a renamed column needs its name fixed on the sheet."
        Exit Sub
    End If

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    blanks = ReadRows(ws, hdr, lastRow, cols)
    If rN = 0 Then
        Say "No report rows under the headers on the Report sheet."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    NetReversals
    ApplyFilters
    WriteSheets wb, blanks
    Application.ScreenUpdating = True
End Sub

Private Function FindSheet(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then Set FindSheet = ws: Exit Function
    Next ws
End Function

Private Function FindHeaderRow(ws As Worksheet, cols() As Long, missing As String) As Long
    ' The header row is the first row with both BOL and Destination State in it.
    Dim r As Long, c As Long, lastCol As Long, names() As String, i As Long, t As String
    names = Split(COL_NAMES, "|")
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For r = 1 To 20
        If HasCell(ws, r, lastCol, "BOL") And HasCell(ws, r, lastCol, "Destination State") Then
            For i = 0 To UBound(names)
                cols(i) = 0
                For c = 1 To lastCol
                    t = Trim(CStr(ws.Cells(r, c).Value))
                    If StrComp(t, names(i), vbTextCompare) = 0 Then cols(i) = c: Exit For
                Next c
                If cols(i) = 0 Then missing = missing & IIf(missing = "", "", ", ") & names(i)
            Next i
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function HasCell(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal txt As String) As Boolean
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim(CStr(ws.Cells(r, c).Value)), txt, vbTextCompare) = 0 Then HasCell = True: Exit Function
    Next c
End Function

Private Function ReadRows(ws As Worksheet, ByVal hdr As Long, ByVal lastRow As Long, cols() As Long) As Long
    ' Fills the r arrays. Returns how many blank rows were skipped.
    Dim r As Long, n As Long, blanks As Long, bol As String, prod As String, gal As Variant
    n = lastRow - hdr
    If n < 1 Then Exit Function
    ReDim rRow(1 To n), rBol(1 To n), rDate(1 To n), rTerm(1 To n), rTState(1 To n), rDest(1 To n)
    ReDim rProd(1 To n), rGal(1 To n), rDol(1 To n), rKept(1 To n), rWhy(1 To n)
    For r = hdr + 1 To lastRow
        bol = CellText(ws.Cells(r, cols(K_BOL)).Value)
        prod = UCase(CellText(ws.Cells(r, cols(K_PROD)).Value))
        gal = ws.Cells(r, cols(K_GAL)).Value
        If bol = "" And prod = "" And IsEmpty(gal) Then
            blanks = blanks + 1
        Else
            rN = rN + 1
            rRow(rN) = r
            rBol(rN) = bol
            rProd(rN) = prod
            rGal(rN) = gal
            rDate(rN) = ws.Cells(r, cols(K_DATE)).Value
            rTerm(rN) = CellText(ws.Cells(r, cols(K_TERM)).Value)
            rTState(rN) = UCase(CellText(ws.Cells(r, cols(K_TSTATE)).Value))
            rDest(rN) = UCase(CellText(ws.Cells(r, cols(K_DEST)).Value))
            rDol(rN) = ws.Cells(r, cols(K_DOL)).Value
            rKept(rN) = True
            rWhy(rN) = ""
        End If
    Next r
    ReadRows = blanks
End Function

Private Function CellText(ByVal v As Variant) As String
    ' Numbers come back without a trailing .0, so a BOL typed as a number reads as typed.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Then
        If v = Int(v) Then CellText = Format(v, "0"): Exit Function
    End If
    CellText = Trim(CStr(v))
End Function

Private Function GoodNumber(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    GoodNumber = IsNumeric(v)
End Function

Private Function DayOf(ByVal v As Variant) As Variant
    ' The day part of a loading date, or Empty if it isn't a date.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsDate(v) Then DayOf = CLng(Int(CDate(v)))
End Function

Private Function DollarsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then DollarsBlank = True: Exit Function
    DollarsBlank = (Trim(CStr(v)) = "")
End Function

Private Sub NetReversals()
    ' A row with negative gallons cancels a row with the same BOL, day, product and the same
    ' gallons positive. Both come out; the real load (the rebill) stays. When more than one
    ' row could match, the one whose dollars look like the reversal's (both blank, or the
    ' same amount) goes first, then the nearest one above it.
    Dim i As Long, j As Long, best As Long, pass As Long
    For i = 1 To rN
        If GoodNumber(rGal(i)) And Not IsEmpty(DayOf(rDate(i))) Then
            If CDbl(rGal(i)) < 0 And rKept(i) Then
                best = 0
                For pass = 1 To 2
                    For j = 1 To rN
                        If j <> i And rKept(j) And CancelsOut(i, j) Then
                            If pass = 2 Or SameDollars(rDol(i), rDol(j)) Then
                                If best = 0 Then
                                    best = j
                                ElseIf Abs(j - i) < Abs(best - i) Or (Abs(j - i) = Abs(best - i) And j < i) Then
                                    best = j
                                End If
                            End If
                        End If
                    Next j
                    If best > 0 Then Exit For
                Next pass
                If best > 0 Then
                    rKept(i) = False
                    rWhy(i) = "reversal, cancels report row " & rRow(best)
                    rKept(best) = False
                    rWhy(best) = "reversed by report row " & rRow(i)
                Else
                    rKept(i) = False
                    rWhy(i) = "negative gallons with no load to cancel"
                End If
            End If
        End If
    Next i
End Sub

Private Function CancelsOut(ByVal i As Long, ByVal j As Long) As Boolean
    If Not GoodNumber(rGal(j)) Then Exit Function
    If IsEmpty(DayOf(rDate(j))) Then Exit Function
    If CDbl(rGal(j)) <= 0 Then Exit Function
    If Abs(CDbl(rGal(i)) + CDbl(rGal(j))) > 0.0001 Then Exit Function
    If rBol(i) <> rBol(j) Or rProd(i) <> rProd(j) Then Exit Function
    If DayOf(rDate(i)) <> DayOf(rDate(j)) Then Exit Function
    CancelsOut = True
End Function

Private Function SameDollars(ByVal a As Variant, ByVal b As Variant) As Boolean
    If DollarsBlank(a) Or DollarsBlank(b) Then
        SameDollars = DollarsBlank(a) And DollarsBlank(b)
    ElseIf GoodNumber(a) And GoodNumber(b) Then
        SameDollars = (Abs(Abs(CDbl(a)) - Abs(CDbl(b))) < 0.005)
    End If
End Function

Private Sub ApplyFilters()
    ' Keeps import loads only. Every reason that applies is listed, not just the first.
    Dim i As Long, why As String, p As String
    For i = 1 To rN
        If rKept(i) Then
            why = ""
            p = "|" & rProd(i) & "|"
            If IsEmpty(DayOf(rDate(i))) Then why = AddWhy(why, "loading date isn't a date")
            If Not GoodNumber(rGal(i)) Then
                why = AddWhy(why, "gallons aren't a number")
            ElseIf CDbl(rGal(i)) = 0 Then
                why = AddWhy(why, "zero gallons")
            End If
            If rTState(i) = HOME_STATE Then why = AddWhy(why, "loaded at a New Mexico terminal, not an import")
            If rDest(i) <> HOME_STATE Then why = AddWhy(why, "destination isn't New Mexico")
            If Not DollarsBlank(rDol(i)) Then
                If Not GoodNumber(rDol(i)) Then
                    why = AddWhy(why, "dollars aren't a number")
                ElseIf CDbl(rDol(i)) <> 0 Then
                    why = AddWhy(why, "dollars not zero, so not the CTFP line")
                End If
            End If
            If InStr(DYED_PRODUCTS, p) > 0 Then
                If Not REPORT_DYED Then why = AddWhy(why, "dyed, left out (REPORT_DYED setting)")
            ElseIf InStr(DIESEL_PRODUCTS, p) = 0 And InStr(E10_PRODUCTS, p) = 0 Then
                why = AddWhy(why, "product " & rProd(i) & " isn't in the settings")
            End If
            If why <> "" Then
                rKept(i) = False
                rWhy(i) = why
            End If
        End If
    Next i
End Sub

Private Function AddWhy(ByVal why As String, ByVal more As String) As String
    AddWhy = why & IIf(why = "", "", "; ") & more
End Function

Private Sub WriteSheets(wb As Workbook, ByVal blanks As Long)
    Dim order() As Long, nk As Long, i As Long, j As Long, k As Long, t As Long
    Dim up() As Variant, nUp As Long, gal As Double, eth As Double, p As String
    Dim keptGal As Double, upGal As Double, quarters As String, q As String, dyedType As String
    Dim ws As Worksheet, ck() As Variant, nDrop As Long, ln As Long, dupes As String

    dyedType = Replace(TT_DYED, " - ", " " & ChrW(8211) & " ")

    ' kept rows by loading date, then report order
    ReDim order(1 To rN)
    For i = 1 To rN
        If rKept(i) Then nk = nk + 1: order(nk) = i
    Next i
    For i = 2 To nk
        t = order(i)
        j = i - 1
        Do While j >= 1
            If CDate(rDate(order(j))) <= CDate(rDate(t)) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = t
    Next i

    ' same BOL, day, product and gallons twice, both kept: a look, not a drop
    For i = 1 To nk
        For j = i + 1 To nk
            If rBol(order(i)) = rBol(order(j)) And rProd(order(i)) = rProd(order(j)) And _
               DayOf(rDate(order(i))) = DayOf(rDate(order(j))) And _
               CDbl(rGal(order(i))) = CDbl(rGal(order(j))) Then
                dupes = dupes & IIf(dupes = "", "", ", ") & rRow(order(i)) & " and " & rRow(order(j))
            End If
        Next j
    Next i

    ReDim up(1 To IIf(nk = 0, 1, nk * 2), 1 To UP_COLS)
    For k = 1 To nk
        i = order(k)
        gal = Int(CDbl(rGal(i)) + 0.5)
        keptGal = keptGal + gal
        p = "|" & rProd(i) & "|"
        If InStr(E10_PRODUCTS, p) > 0 Then
            eth = Int(gal * ETH_PCT / 100 + 0.5)
            AddUp up, nUp, i, GAS_NAME, GAS_CODE, GAS_VEH, TT_IMPORT, gal - eth, BLEND_E10
            AddUp up, nUp, i, ETH_NAME, ETH_CODE, ETH_VEH, TT_IMPORT, eth, BLEND_E10
            upGal = upGal + gal
        ElseIf InStr(DYED_PRODUCTS, p) > 0 Then
            AddUp up, nUp, i, DIESEL_NAME, DIESEL_CODE, DIESEL_VEH, TT_IMPORT, gal, BLEND_OTHER
            AddUp up, nUp, i, DIESEL_NAME, DIESEL_CODE, DIESEL_VEH, dyedType, gal, BLEND_OTHER
            upGal = upGal + gal
        Else
            AddUp up, nUp, i, DIESEL_NAME, DIESEL_CODE, DIESEL_VEH, TT_IMPORT, gal, BLEND_OTHER
            upGal = upGal + gal
        End If
        q = Year(CDate(rDate(i))) & " " & QTR_PREFIX & ((Month(CDate(rDate(i))) - 1) \ 3 + 1)
        If InStr("|" & quarters & "|", "|" & q & "|") = 0 Then quarters = quarters & IIf(quarters = "", "", "|") & q
    Next k

    ' Upload
    Set ws = FreshSheet(wb, SH_UP)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UP_COLS)).Value = Split(UP_HEADS, "|")
    ws.Rows(1).Font.Bold = True
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(5).NumberFormat = "@"
    ws.Columns(16).NumberFormat = "@"
    If nUp > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(nUp + 1, UP_COLS)).Value = up
        ws.Range(ws.Cells(2, 3), ws.Cells(nUp + 1, 3)).NumberFormat = "m/d/yyyy h:mm"
        ws.Range(ws.Cells(2, 8), ws.Cells(nUp + 1, 8)).NumberFormat = "#,##0"
    End If
    ws.Columns("A:P").AutoFit

    ' CHECK: left-out rows first, then kept, each in report order
    ReDim ck(1 To rN, 1 To CHECK_COLS)
    ln = 0
    For t = 1 To 2
        For i = 1 To rN
            If (t = 1 And Not rKept(i)) Or (t = 2 And rKept(i)) Then
                ln = ln + 1
                ck(ln, 1) = rRow(i)
                ck(ln, 2) = rBol(i)
                ck(ln, 3) = rDate(i)
                ck(ln, 4) = rTState(i)
                ck(ln, 5) = rDest(i)
                ck(ln, 6) = rProd(i)
                ck(ln, 7) = rGal(i)
                ck(ln, 8) = rDol(i)
                If rKept(i) Then
                    ck(ln, 9) = "Yes"
                    p = "|" & rProd(i) & "|"
                    If InStr(E10_PRODUCTS, p) > 0 Then
                        ck(ln, 10) = "kept, split into gasoline and ethanol"
                    ElseIf InStr(DYED_PRODUCTS, p) > 0 Then
                        ck(ln, 10) = "kept, import plus the dyed exempt line"
                    Else
                        ck(ln, 10) = "kept"
                    End If
                Else
                    ck(ln, 9) = "No"
                    ck(ln, 10) = rWhy(i)
                    nDrop = nDrop + 1
                End If
            End If
        Next i
    Next t

    Set ws = FreshSheet(wb, SH_CHECK)
    ws.Range("A1").Value = "NM QFR rows, built " & Format(Now, "mmm d, yyyy h:mm am/pm")
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Report rows read: " & rN & ". Left out: " & nDrop & ". Loads kept: " & nk & _
        ". Upload rows: " & nUp & "." & IIf(blanks > 0, " Blank rows skipped: " & blanks & ".", "")
    ws.Range("A3").Value = "Gallons kept: " & Format(keptGal, "#,##0") & ". Gallons on Upload, " & _
        "E10 counted once: " & Format(upGal, "#,##0") & "." & _
        IIf(keptGal = upGal, " They match.", " THEY DON'T MATCH, look before uploading.")
    If InStr(quarters, "|") > 0 Then
        ws.Range("A4").Value = "LOOK: kept loads fall in more than one quarter (" & _
            Replace(quarters, "|", ", ") & "). One upload is one quarter."
        ws.Range("A4").Font.Color = RGB(192, 0, 0)
    ElseIf quarters <> "" Then
        ws.Range("A4").Value = "Quarter: " & quarters & "."
    End If
    If dupes <> "" Then
        ws.Range("A5").Value = "LOOK: same BOL, day, product and gallons twice, both kept (report rows " & _
            dupes & "). Delete one from Upload if it's a double pull."
        ws.Range("A5").Font.Color = RGB(192, 0, 0)
    End If
    ws.Range(ws.Cells(CHECK_TOP, 1), ws.Cells(CHECK_TOP, CHECK_COLS)).Value = Split(CHECK_HEADS, "|")
    ws.Rows(CHECK_TOP).Font.Bold = True
    ws.Columns(2).NumberFormat = "@"
    ws.Range(ws.Cells(CHECK_TOP + 1, 1), ws.Cells(CHECK_TOP + rN, CHECK_COLS)).Value = ck
    ws.Range(ws.Cells(CHECK_TOP + 1, 3), ws.Cells(CHECK_TOP + rN, 3)).NumberFormat = "m/d/yyyy h:mm"
    If nDrop > 0 Then
        ws.Range(ws.Cells(CHECK_TOP + 1, 1), ws.Cells(CHECK_TOP + nDrop, CHECK_COLS)).Interior.Color = RGB(255, 242, 204)
    End If
    ws.Columns("A:I").AutoFit
    ws.Columns(10).ColumnWidth = 60

    wb.Worksheets(SH_UP).Activate
    Say nUp & " upload rows from " & nk & " loads. " & nDrop & " report rows left out; CHECK says why." & _
        IIf(keptGal = upGal, "", vbLf & vbLf & "The gallons don't add up. Look at CHECK before uploading.") & _
        IIf(InStr(quarters, "|") > 0, vbLf & vbLf & "The loads fall in more than one quarter. See CHECK.", "") & _
        IIf(dupes <> "", vbLf & vbLf & "Some loads look doubled. See CHECK.", "")
End Sub

Private Sub AddUp(up() As Variant, nUp As Long, ByVal i As Long, ByVal fuel As String, ByVal code As String, _
                  ByVal veh As String, ByVal ttype As String, ByVal vol As Double, ByVal blend As String)
    Dim d As Date
    d = CDate(rDate(i))
    nUp = nUp + 1
    up(nUp, 1) = Year(d)
    up(nUp, 2) = QTR_PREFIX & ((Month(d) - 1) \ 3 + 1)
    up(nUp, 3) = d
    up(nUp, 4) = fuel
    up(nUp, 5) = code
    up(nUp, 6) = veh
    up(nUp, 7) = ttype
    up(nUp, 8) = vol
    up(nUp, 9) = IIf(PARTNER_IS_TERMINAL, rTerm(i), "")
    up(nUp, 10) = ""
    up(nUp, 11) = IND_NO
    up(nUp, 12) = ""
    up(nUp, 13) = blend
    up(nUp, 14) = ""
    up(nUp, 15) = IND_NO
    up(nUp, 16) = DESC_PREFIX & rBol(i)
End Sub

Private Function FreshSheet(wb As Workbook, ByVal nm As String) As Worksheet
    ' The sheet emptied, or made after the Report sheet if it isn't there.
    Dim ws As Worksheet
    Set ws = FindSheet(wb, nm)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set FreshSheet = ws
End Function
