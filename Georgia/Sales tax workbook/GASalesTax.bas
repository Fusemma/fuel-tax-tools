' Georgia sales tax: the rates come from the state's import template, and the two files to
' import get made from it.
'
' Get rates: pick the quarter's ST-3 Excel import template, downloaded from
' dor.georgia.gov/sales-use-tax-import. The rate table hidden inside it goes onto the RATES
' tab, and both return tabs take every county's rate from there for the month on RECAP.
'
' Make the state's files: fills two copies of that template, one per return, and saves them
' beside this workbook, like GA ST-3 off-road 2026-09.xls. Then a box compares each copy's
' Main tab with RECAP. Import each copy on its return in GTC.
'
' If the .xlsm won't download, set it up from this text once:
'   1. Open GA Sales Tax Workbook - BLANK.xlsx.
'   2. Alt F11, then Insert > Module. Paste all of this text, then close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm).
'   4. Alt F8 runs GetRates and MakeStateFiles.

Option Explicit

Private Const BOX_TITLE As String = "Georgia sales tax"
Private Const RATE_SHEET As String = "tax rate table"
Private Const MACROS_OFF As Long = 3        ' msoAutomationSecurityForceDisable

Public Sub GetRates()
    Dim f As Variant, why As String
    f = PickTemplate()
    If VarType(f) = vbBoolean Then Exit Sub
    why = LoadRates(CStr(f))
    If why <> "" Then
        MsgBox why, vbExclamation, BOX_TITLE
    Else
        MsgBox "Rates from " & ThisWorkbook.Worksheets("RATES").Range("H1").Value & ", for " & _
            Format$(ThisWorkbook.Worksheets("RATES").Range("H2").Value, "m/d/yyyy") & " to " & _
            Format$(ThisWorkbook.Worksheets("RATES").Range("I2").Value, "m/d/yyyy") & "." & vbLf & vbLf & _
            ThisWorkbook.Worksheets("RATES").Range("H3").Value & " rate rows. RECAP and the return tabs use them now.", _
            vbInformation, BOX_TITLE
    End If
End Sub

Public Sub MakeStateFiles()
    Dim looks As Variant
    Application.CalculateFull
    looks = ThisWorkbook.Worksheets("REVIEW").Range("E2").Value
    If IsNumeric(looks) Then
        If looks > 0 Then
            If MsgBox("REVIEW has " & looks & IIf(looks = 1, " thing", " things") & " to look at." & vbLf & vbLf & _
                "Make the state's files anyway?", vbYesNo + vbQuestion + vbDefaultButton2, BOX_TITLE) = vbNo Then Exit Sub
        End If
    End If
    If TemplatePath() = "" Then
        Dim f As Variant, why As String
        f = PickTemplate()
        If VarType(f) = vbBoolean Then Exit Sub
        why = LoadRates(CStr(f))
        If why <> "" Then
            MsgBox why, vbExclamation, BOX_TITLE
            Exit Sub
        End If
    End If
    Dim msg As String
    msg = MakeFilesFrom(TemplatePath())
    MsgBox msg, IIf(InStr(msg, "DIFFERS") > 0 Or InStr(msg, "Nothing") > 0, vbExclamation, vbInformation), BOX_TITLE
End Sub

' The template RATES came from, if it's still where it was.
Private Function TemplatePath() As String
    Dim p As String
    p = CStr(ThisWorkbook.Worksheets("RATES").Range("H4").Value)
    If p <> "" Then
        If Dir(p) <> "" Then TemplatePath = p
    End If
End Function

Private Function PickTemplate() As Variant
    PickTemplate = Application.GetOpenFilename("Excel files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , _
        "Pick the quarter's ST-3 import template")
End Function

' Copies the template's hidden rate table onto RATES: code, from, to, state part, local part,
' total. Finds the columns by their names, so a moved column doesn't matter. Returns why it
' couldn't, or "" when it worked.
Public Function LoadRates(ByVal path As String) As String
    Dim tpl As Workbook, ws As Worksheet, v As Variant, outRows() As Variant
    Dim hr As Long, r As Long, c As Long, n As Long
    Dim cCode As Long, cFrom As Long, cTo As Long, cState As Long, cLocal As Long
    Dim covFrom As Variant, covTo As Variant, why As String
    Application.ScreenUpdating = False
    On Error GoTo Failed
    Set tpl = OpenQuiet(path, True)
    On Error Resume Next
    Set ws = tpl.Worksheets(RATE_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        why = "That file has no rate table in it:" & vbLf & path & vbLf & vbLf & _
            "Pick the state's ST-3 Excel import template, from dor.georgia.gov/sales-use-tax-import."
        GoTo Done
    End If
    With ws.UsedRange
        v = ws.Range("A1", .Cells(.Rows.Count, .Columns.Count)).Value2
    End With
    For r = 1 To Application.Min(60, UBound(v, 1))
        For c = 1 To UBound(v, 2)
            If LCase$(Txt(v(r, c))) = "date range" And r + 2 <= UBound(v, 1) Then
                covTo = AsDate(v(r + 1, c))
                covFrom = AsDate(v(r + 2, c))
            End If
            If Txt(v(r, c)) = "fstrDecode2" Then hr = r
        Next c
        If hr > 0 Then Exit For
    Next r
    If hr > 0 Then
        For c = 1 To UBound(v, 2)
            Select Case Txt(v(hr, c))
                Case "fstrDecode2": cCode = c
                Case "fdtmPeriodBegin": cFrom = c
                Case "fdtmPeriodEnd": cTo = c
                Case "fcurStateOverallRate": cState = c
                Case "fcurOverallRate": cLocal = c
            End Select
        Next c
    End If
    If cCode * cFrom * cTo * cState * cLocal = 0 Or IsEmpty(covFrom) Or IsEmpty(covTo) Then
        why = "The rate table in this template isn't laid out the way it used to be:" & vbLf & path & vbLf & vbLf & _
            "The state changed its template. The rates on RATES stay as they were."
        GoTo Done
    End If
    n = 0
    For r = hr + 1 To UBound(v, 1)
        If Txt(v(r, cCode)) <> "" Then n = n + 1
    Next r
    If n = 0 Then
        why = "The rate table in this template is empty:" & vbLf & path
        GoTo Done
    End If
    ReDim outRows(1 To n, 1 To 6)
    n = 0
    For r = hr + 1 To UBound(v, 1)
        If Txt(v(r, cCode)) <> "" Then
            n = n + 1
            outRows(n, 1) = Txt(v(r, cCode))
            outRows(n, 2) = v(r, cFrom)
            outRows(n, 3) = v(r, cTo)
            outRows(n, 4) = v(r, cState)
            outRows(n, 5) = v(r, cLocal)
            outRows(n, 6) = Val(Txt(v(r, cState))) + Val(Txt(v(r, cLocal)))
        End If
    Next r
    With ThisWorkbook.Worksheets("RATES")
        .Range("A2:F" & .Rows.Count).ClearContents
        .Range("A2").Resize(n, 6).Value = outRows
        .Range("H1").Value = Mid$(path, InStrRev(path, "\") + 1)
        .Range("H2").Value = CDate(covFrom)
        .Range("I2").Value = CDate(covTo)
        .Range("H3").Value = n
        .Range("H4").Value = path
    End With
Done:
    If Not tpl Is Nothing Then tpl.Close False
    Application.ScreenUpdating = True
    Application.CalculateFull
    LoadRates = why
    Exit Function
Failed:
    why = "Couldn't read the rates: " & Err.Description
    Resume Done
End Function

' Fills one copy of the template per return and says how each one's Main tab compares with RECAP.
Public Function MakeFilesFrom(ByVal tplPath As String) As String
    Dim mon As Variant, monEnd As Date, covFrom As Variant, covTo As Variant
    On Error GoTo Failed
    mon = Nm("ReturnMonth").Value
    If Not IsDate(mon) Or IsEmpty(mon) Then
        MakeFilesFrom = "Nothing made. Type the month on RECAP first, like 9/1/2026."
        Exit Function
    End If
    monEnd = DateSerial(Year(mon), Month(mon) + 1, 0)
    covFrom = ThisWorkbook.Worksheets("RATES").Range("H2").Value
    covTo = ThisWorkbook.Worksheets("RATES").Range("I2").Value
    If Not IsDate(covFrom) Or Not IsDate(covTo) Then
        MakeFilesFrom = "Nothing made. Press Get rates and pick this quarter's template first."
        Exit Function
    End If
    If DateSerial(Year(mon), Month(mon), 1) < covFrom Or monEnd > covTo Then
        MakeFilesFrom = "Nothing made. The template on RATES is for " & Format$(covFrom, "m/d/yyyy") & " to " & _
            Format$(covTo, "m/d/yyyy") & ", and RECAP's month is " & Format$(mon, "mmmm yyyy") & "." & vbLf & vbLf & _
            "Download that quarter's template, press Get rates and pick it."
        Exit Function
    End If
    Application.ScreenUpdating = False
    Application.CalculateFull
    MakeFilesFrom = FillOne(True, tplPath, monEnd) & vbLf & vbLf & FillOne(False, tplPath, monEnd)
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Function
Failed:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MakeFilesFrom = "Nothing made. " & Stopped("either return", "reading RECAP and RATES")
End Function

' Shows the step on Excel's status bar, at the bottom of the window, and hands it back.
Private Function Progress(ByVal road As String, ByVal stepAt As String) As String
    Application.StatusBar = "Georgia sales tax, " & road & ": " & stepAt & "..."
    Progress = stepAt
End Function

' Says what stopped it and where, instead of dropping into the code.
Private Function Stopped(ByVal road As String, ByVal stepAt As String) As String
    Stopped = "Nothing saved for " & road & ". It stopped while " & stepAt & ": " & Err.Description & _
        " (error " & Err.Number & ")"
End Function

Private Function FillOne(ByVal offRoad As Boolean, ByVal tplPath As String, ByVal monEnd As Date) As String
    Dim road As String, dest As String, wb As Workbook, m As Worksheet, dist As Worksheet, mfd As Worksheet
    Dim lines As Range, acct As Variant, i As Long, k As Long, j As Long, msg As String
    Dim codes As Variant, a As Variant, b As Variant, cc As Variant, d As Variant, e As Variant
    Dim stepAt As String
    road = IIf(offRoad, "off-road", "on-road")
    On Error GoTo Failed
    stepAt = Progress(road, "reading RECAP")
    Set lines = Nm(IIf(offRoad, "OffLines", "OnLines"))
    dest = FreeName(ThisWorkbook.Path & "\GA ST-3 " & road & " " & Format$(monEnd, "yyyy-mm"), _
        Mid$(tplPath, InStrRev(tplPath, ".")))
    stepAt = Progress(road, "copying the template")
    FileCopy tplPath, dest
    stepAt = Progress(road, "opening the copy")
    ' Events off, so none of the template's own code runs. Not "macros off": Excel quietly ends
    ' the macro that saves a workbook it opened that way.
    Application.EnableEvents = False
    Set wb = Workbooks.Open(Filename:=dest, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    stepAt = Progress(road, "filling its Main tab")
    Set m = wb.Worksheets("Main")
    acct = Nm(IIf(offRoad, "OffAccount", "OnAccount")).Value
    If Txt(acct) <> "" Then m.Range("D4").Value = acct
    m.Range("D5").Value = monEnd
    m.Range("D6").Value = "Monthly"
    m.Range("K6").Value = "Yes"
    m.Range("D10").Value = Round(lines.Cells(1, 1).Value, 2)
    m.Range("D11").Value = Round(lines.Cells(2, 1).Value, 2)
    m.Range("D16").Value = Round(lines.Cells(7, 1).Value, 2)
    stepAt = Progress(road, "filling the county rows")
    Set dist = wb.Worksheets("Sales Tax Distribution")
    k = 0
    If offRoad Then
        ' the 4% on PART B; each county's local tax from its row on the distribution tab
        wb.Worksheets("PART B").Range("D6").Value = Round(lines.Cells(3, 1).Value, 2)
        codes = Nm("OffCodes").Value
        a = Nm("OffTaxable").Value
        For i = 1 To UBound(codes, 1)
            If Txt(codes(i, 1)) <> "" And Round(Val(Txt(a(i, 1))), 2) <> 0 Then
                k = k + 1
                dist.Cells(2 + k, 1).Value = codes(i, 1)
                dist.Cells(2 + k, 4).Value = Round(a(i, 1), 2)
            End If
        Next i
        msg = k & IIf(k = 1, " county", " counties") & " on the distribution tab."
    Else
        ' College Park on the distribution tab (line 4), every other county on the motor fuel tab (line 6)
        Set mfd = wb.Worksheets("MFD On-Road Worksheet")
        codes = Nm("OnCodes").Value
        a = Nm("OnFuelType").Value
        b = Nm("OnGallons").Value
        cc = Nm("OnExempt").Value
        d = Nm("OnTaxableSales").Value
        e = Nm("OnGoesOn").Value
        For i = 1 To UBound(codes, 1)
            If Txt(codes(i, 1)) <> "" And Val(Txt(b(i, 1))) <> 0 Then
                If Txt(e(i, 1)) = "line 4" Then
                    k = k + 1
                    dist.Cells(2 + k, 1).Value = codes(i, 1)
                    dist.Cells(2 + k, 4).Value = Round(d(i, 1), 2)
                Else
                    j = j + 1
                    mfd.Cells(2 + j, 1).Value = codes(i, 1)
                    mfd.Cells(2 + j, 3).Value = a(i, 1)
                    mfd.Cells(2 + j, 9).Value = b(i, 1)
                    mfd.Cells(2 + j, 10).Value = cc(i, 1)
                End If
            End If
        Next i
        msg = j & IIf(j = 1, " row", " rows") & " on the motor fuel tab, " & k & " on the distribution tab."
    End If
    stepAt = Progress(road, "checking its Main tab")
    Application.Calculate
    msg = "Saved " & Mid$(dest, InStrRev(dest, "\") + 1) & ". " & msg & vbLf & "Its Main tab, and RECAP:" & vbLf
    If offRoad Then msg = msg & LineCheck(m, "Line 3", "D12", lines.Cells(3, 1).Value)
    msg = msg & LineCheck(m, "Line 4", "D13", lines.Cells(4, 1).Value)
    If Not offRoad Then msg = msg & LineCheck(m, "Line 6", "D15", lines.Cells(6, 1).Value)
    msg = msg & LineCheck(m, "Line 8", "D17", lines.Cells(8, 1).Value)
    msg = msg & LineCheck(m, "Line 9", "D18", lines.Cells(9, 1).Value)
    msg = msg & LineCheck(m, "Line 10", "D19", lines.Cells(10, 1).Value)
    msg = msg & LineCheck(m, "Total due", "D20", _
        lines.Cells(8, 1).Value + lines.Cells(9, 1).Value - lines.Cells(10, 1).Value)
    ' an .xls asks about features older Excel lacks; the state's own template is fine as it is
    stepAt = Progress(road, "saving the copy")
    Application.DisplayAlerts = False
    wb.Save
    Application.DisplayAlerts = True
    wb.Close False
    Application.EnableEvents = True
    FillOne = msg
    Exit Function
Failed:
    Application.DisplayAlerts = True
    FillOne = Stopped(road, stepAt)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True
End Function

Private Function LineCheck(ByVal m As Worksheet, ByVal label As String, ByVal cell As String, ByVal want As Variant) As String
    Dim got As Double
    got = Round(Val(Txt(m.Range(cell).Value)), 2)
    LineCheck = "  " & label & "  " & Format$(got, "#,##0.00")
    If Abs(got - Round(want, 2)) >= 0.005 Then LineCheck = LineCheck & "  DIFFERS, RECAP says " & Format$(want, "#,##0.00")
    LineCheck = LineCheck & vbLf
End Function

' A named range in this workbook, never in the template that's open beside it.
Private Function Nm(ByVal name As String) As Range
    Set Nm = ThisWorkbook.Names(name).RefersToRange
End Function

' A date from a cell that holds one as a date, a date number, or text like 7/1/2026; Empty if not.
Private Function AsDate(ByVal x As Variant) As Variant
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If IsNumeric(x) Then
        AsDate = CDate(CDbl(x))
    ElseIf IsDate(x) Then
        AsDate = CDate(x)
    End If
End Function

' The cell as text; an error value, like #N/A, comes back blank.
Private Function Txt(ByVal x As Variant) As String
    If Not IsError(x) Then Txt = CStr(x)
End Function

' name.xls, or name (2).xls if that's taken, and so on.
Private Function FreeName(ByVal stem As String, ByVal ext As String) As String
    Dim n As Long
    FreeName = stem & ext
    n = 1
    Do While Dir(FreeName) <> ""
        n = n + 1
        FreeName = stem & " (" & n & ")" & ext
    Loop
End Function

' Opens a workbook with its macros off, the way the state's template works fine.
Private Function OpenQuiet(ByVal path As String, ByVal readOnly As Boolean) As Workbook
    Dim was As Long, num As Long, desc As String
    was = Application.AutomationSecurity
    Application.AutomationSecurity = MACROS_OFF
    On Error Resume Next
    Set OpenQuiet = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=readOnly, AddToMru:=False)
    num = Err.Number
    desc = Err.Description
    On Error GoTo 0
    Application.AutomationSecurity = was
    If num <> 0 Then Err.Raise num, , desc
End Function
