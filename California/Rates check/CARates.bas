' California rates check: shows what changed in the state's city and county sales and use
' tax rates each quarter, from the CDTFA's rates spreadsheet.
'
' Setting it up, once:
'   1. Copy all of this text (click Raw, then Ctrl A, Ctrl C).
'   2. Open a new blank workbook in Excel. Alt F11, then Insert > Module. If the new window
'      already has a line in it, Ctrl A and Delete first. Paste, then close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm). Name it CA Rates Check and keep
'      it with your other work files.
'   4. Alt F8, pick CheckCaliforniaRates, Run. It makes the sheets and the big button, then
'      asks for the rates file.
' The How to use sheet it makes has the steps for each quarter.

Option Explicit

Private Const SH_MAIN As String = "What changed"
Private Const SH_HOWTO As String = "How to use"
Private Const SH_LOG As String = "Log"
Private Const SH_SAVED As String = "Saved rates"
Private Const BTN_NAME As String = "CheckRatesButton"
Private Const BTN_TEXT As String = "Check new rates file"
Private Const BOX_TITLE As String = "California rates"
Private Const TITLE_TEXT As String = "California rates: what changed"
Private Const WHERE_TEXT As String = "cdtfa.ca.gov, See Tax Rates, Sales and Use Tax Rates, then " & _
    "the Excel (xlsx) file under Download"

' The state's columns, in the order Saved rates keeps them. The file's columns are found by
' name, so a list with its columns moved around still reads right.
Private Const HEADS As String = "Location|Rate|County|Type|Notes"
Private Const NC As Long = 5
Private Const C_PLACE As Long = 1, C_RATE As Long = 2, C_COUNTY As Long = 3, C_TYPE As Long = 4
Private Const C_NOTES As Long = 5

Private Const LIST_HEADS As String = "Place|County|What changed|Old rate|New rate|Starts|Notes"
Private Const LIST_COLS As Long = 7
Private Const LOG_HEADS As String = "Checked|File|" & LIST_HEADS & "|Change ID"
Private Const LOG_COLS As Long = 10
Private Const SAVED_ROW As Long = 4         ' header row on Saved rates
Private Const ALERT_ROW As Long = 4         ' the get-the-new-file line on What changed
Private Const RATE_LOOK As String = "0.000%"

' One change, as kept while sorting and drawing. 1 to 7 are the list's columns in order.
Private Const F_SORT As Long = 0, F_PLACE As Long = 1, F_COUNTY As Long = 2, F_WHAT As Long = 3
Private Const F_OLD As Long = 4, F_NEW As Long = 5, F_START As Long = 6, F_NOTES As Long = 7
Private Const F_LOOK As Long = 8, F_ID As Long = 9

Private mFeed As Collection     ' test runs only: answers for the boxes, in order
Private mSaid As String         ' test runs only: what the boxes would have said
Private mToday As Date          ' test runs only: the day to act as today


' ---- the button, and opening the workbook ------------------------------------------

Public Sub CheckCaliforniaRates()
    ' Asks for the new rates list, then shows what changed since the last one.
    On Error GoTo Oops
    LoadFile
    GoTo Tidy
Oops:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Say "Something went wrong: " & Err.Description & vbLf & vbLf & _
        "Try it again. If it keeps happening, download the file again and pick that."
Tidy:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
End Sub

Public Sub Auto_Open()
    ' Opening the workbook redraws only the get-the-new-file line. That isn't anything to
    ' save, so the workbook is marked unchanged and closing won't ask.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = SheetByName(SH_MAIN)
    If ws Is Nothing Then Exit Sub
    DrawAlert ws
    ThisWorkbook.Saved = True
End Sub

Private Sub LoadFile()
    Dim picked As Variant, fname As String, problem As String, tabName As String
    Dim allRecs As Variant, nAll As Long, newRecs As Variant, nNew As Long, newIdx As Collection
    Dim savedRecs As Variant, nSaved As Long, oldRecs As Variant, nOld As Long, oldIdx As Collection
    Dim savedFile As String, savedWhen As Variant, savedStarts As Date, dups As Long, x As Long
    Dim starts As Date, typed As Boolean, today As Date, changes As Collection

    today = CheckDate()
    EnsureSheets
    picked = PickFile()
    If VarType(picked) = vbBoolean Then Exit Sub
    fname = FileNameOf(CStr(picked))

    Application.Cursor = xlWait
    Application.StatusBar = "Reading " & fname & "..."
    If Not ReadList(CStr(picked), allRecs, nAll, tabName, problem) Then
        Application.Cursor = xlDefault
        Say problem
        Exit Sub
    End If

    ' The day the rates start: from the file's name, else its tab's, else ask
    starts = DateInText(fname)
    If starts = 0 Then starts = DateInText(tabName)
    If starts = 0 Then
        Application.Cursor = xlDefault
        starts = AskDate(QuarterGuess(today))
        If starts = 0 Then Exit Sub
        typed = True
        Application.Cursor = xlWait
    End If

    Set newIdx = New Collection
    nNew = OnePerPlace(allRecs, nAll, newRecs, newIdx, dups)

    ' The list loaded last time, if there was one
    nSaved = ReadSaved(savedRecs, savedFile, savedWhen, savedStarts)
    If nSaved > 0 Then
        Set oldIdx = New Collection
        nOld = OnePerPlace(savedRecs, nSaved, oldRecs, oldIdx, x)
        If starts < savedStarts Then
            Application.Cursor = xlDefault
            If Not Confirm("This list looks older than the one loaded last time." & vbLf & vbLf & _
                    "Its rates start " & ShowDate(starts) & ". The saved ones start " & _
                    ShowDate(savedStarts) & "." & vbLf & vbLf & _
                    "Load it anyway? It will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
        If nNew < nOld * 0.9 Then
            Application.Cursor = xlDefault
            If Not Confirm("This list has only " & Format$(nNew, "#,##0") & " places. The saved one has " & _
                    Format$(nOld, "#,##0") & ", so the download may have been cut short." & vbLf & vbLf & _
                    "Load it anyway? It will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
    End If

    Application.StatusBar = "Comparing..."
    Set changes = New Collection
    If nOld > 0 Then Compare oldRecs, nOld, oldIdx, newRecs, nNew, newIdx, starts, changes

    Application.ScreenUpdating = False
    Application.StatusBar = "Drawing..."
    SaveCopy newRecs, nNew, fname, today, starts
    DrawFront changes, (nOld > 0), fname, today, starts, typed, savedFile, savedWhen, savedStarts, nNew, dups
    AddToLog changes, (nOld > 0), fname, today, nNew
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
    SheetByName(SH_MAIN).Range("A1").Select
End Sub


' ---- reading the state's list ------------------------------------------------------------

Private Function ReadList(ByVal fullPath As String, ByRef recs As Variant, ByRef n As Long, _
                          ByRef tabName As String, ByRef problem As String) As Boolean
    ' Reads the list into recs, one row per line, columns in HEADS order: the place, its rate
    ' as a fraction (Empty when the list gives none), county, type and notes. False, with
    ' problem saying why, when it isn't the rates list.
    Dim wb As Workbook, ws As Worksheet, openedHere As Boolean, headAt As Long, isMap As Boolean
    Dim colOf As Variant, v As Variant, lastRow As Long, i As Long, c As Long, maxCol As Long
    Dim place As String, county As String, rate As Variant, anything As Boolean, rates As Long

    n = 0
    ReDim colOf(1 To NC)
    Set wb = OpenList(fullPath, openedHere, problem)
    If wb Is Nothing Then Exit Function
    On Error GoTo Failed
    For Each ws In wb.Worksheets
        If HasAnything(ws) Then anything = True
        headAt = FindHeader(ws, colOf, isMap)
        If headAt > 0 Then Exit For
    Next
    If headAt = 0 Then
        If isMap Then
            problem = "That's the state's map file from data.ca.gov. This check reads the CDTFA's " & _
                "rates list instead: " & WHERE_TEXT & "." & vbLf & vbLf & "Nothing was changed."
        ElseIf Not anything Then
            problem = "That file is empty. Download the rates list again and pick that." & _
                vbLf & vbLf & "Nothing was changed."
        Else
            problem = "That doesn't look like the California rates list. " & FileNameOf(fullPath) & _
                " has no Location and Rate columns." & vbLf & vbLf & "Get it from " & WHERE_TEXT & _
                ". Nothing was changed."
        End If
        GoTo Done
    End If
    tabName = ws.Name
    For c = 1 To NC
        If colOf(c) > maxCol Then maxCol = colOf(c)
    Next
    lastRow = ws.Cells(ws.Rows.Count, colOf(C_PLACE)).End(xlUp).Row
    If lastRow > headAt Then
        v = ws.Range(ws.Cells(headAt + 1, 1), ws.Cells(lastRow, maxCol)).Value2
        ReDim recs(1 To UBound(v, 1), 1 To NC)
        For i = 1 To UBound(v, 1)
            place = CellText(v(i, colOf(C_PLACE)))
            rate = RateOf(v(i, colOf(C_RATE)))
            county = ""
            If colOf(C_COUNTY) > 0 Then county = CellText(v(i, colOf(C_COUNTY)))
            ' A line with no rate and no county is a title or a footnote, not a place
            If place <> "" And Not (IsEmpty(rate) And county = "") Then
                n = n + 1
                recs(n, C_PLACE) = place
                recs(n, C_RATE) = rate
                recs(n, C_COUNTY) = county
                recs(n, C_TYPE) = ""
                recs(n, C_NOTES) = ""
                If colOf(C_TYPE) > 0 Then recs(n, C_TYPE) = CellText(v(i, colOf(C_TYPE)))
                If colOf(C_NOTES) > 0 Then recs(n, C_NOTES) = CellText(v(i, colOf(C_NOTES)))
                If Not IsEmpty(rate) Then rates = rates + 1
            End If
        Next
    End If
    If rates = 0 Then
        n = 0
        problem = "That file has no rates in it. Download the rates list again and pick that." & _
            vbLf & vbLf & "Nothing was changed."
    Else
        ReadList = True
    End If
Done:
    On Error Resume Next
    If openedHere Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Exit Function
Failed:
    n = 0
    problem = "Couldn't read " & FileNameOf(fullPath) & ": " & Err.Description & vbLf & vbLf & _
        "Nothing was changed."
    Resume Done
End Function

Private Function OpenList(ByVal fullPath As String, ByRef openedHere As Boolean, _
                          ByRef problem As String) As Workbook
    ' The picked file as a workbook, opened read-only with any macros in it off. A file that's
    ' already open in Excel is used as it is and left open, Protected View included (that's
    ' where a download opened from the browser sits).
    Dim wb As Workbook, pv As Object, security As Long, events As Boolean, alerts As Boolean
    Dim there As Boolean, fname As String, clash As Boolean
    openedHere = False
    fname = FileNameOf(fullPath)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then Set OpenList = wb: Exit Function
        If StrComp(wb.Name, fname, vbTextCompare) = 0 Then clash = True
    Next
    On Error Resume Next
    For Each pv In Application.ProtectedViewWindows
        Set wb = Nothing
        Set wb = pv.Workbook
        If Not wb Is Nothing Then
            If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then Set OpenList = wb: Exit Function
            If StrComp(wb.Name, fname, vbTextCompare) = 0 Then clash = True
        End If
    Next
    Err.Clear
    there = (Len(Dir$(fullPath)) > 0)
    On Error GoTo 0
    If clash Then
        ' Excel can't have two files with one name open, and trying can hang it
        problem = "Another file named " & fname & " is open in Excel. Close it, then click the " & _
            "button again." & vbLf & vbLf & "Nothing was changed."
        Exit Function
    End If
    If Not there Then
        problem = "Couldn't find " & fullPath & "." & vbLf & vbLf & "Nothing was changed."
        Exit Function
    End If
    security = Application.AutomationSecurity
    events = Application.EnableEvents
    alerts = Application.DisplayAlerts
    Application.AutomationSecurity = 3          ' macros off in the file being read
    Application.EnableEvents = False
    Application.DisplayAlerts = False           ' a broken file gets this macro's message, not Excel's
    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    If Err.Number <> 0 Then Set wb = Nothing
    On Error GoTo 0
    Application.AutomationSecurity = security
    Application.EnableEvents = events
    Application.DisplayAlerts = alerts
    If wb Is Nothing Then
        problem = "Couldn't open " & fname & "." & vbLf & vbLf & _
            "If it's open in Excel, close it and click the button again. If it isn't, download it " & _
            "again. Nothing was changed."
        Exit Function
    End If
    openedHere = True
    Set OpenList = wb
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByRef colOf As Variant, ByRef isMap As Boolean) As Long
    ' The row, in the first 20, with the Location and Rate headers on it. colOf gets the column
    ' of each of HEADS. isMap is set when it's the state's map file instead of the list.
    Dim v As Variant, r As Long, c As Long, k As Long, h As String
    v = ws.Range("A1:AN20").Value2
    For r = 1 To 20
        For k = 1 To NC
            colOf(k) = 0
        Next
        For c = 1 To 40
            h = SqueezeHead(CellText(v(r, c)))
            If h = "JURISNAME" Then isMap = True
            k = HeadNumber(h)
            If k > 0 Then If colOf(k) = 0 Then colOf(k) = c
        Next
        If colOf(C_PLACE) > 0 And colOf(C_RATE) > 0 Then FindHeader = r: Exit Function
    Next
    For k = 1 To NC
        colOf(k) = 0
    Next
End Function

Private Function HeadNumber(ByVal squeezed As String) As Long
    ' Which of the state's columns a header names, 0 for none
    Dim h As Variant, i As Long
    If squeezed = "" Then Exit Function
    h = Split(HEADS, "|")
    For i = 0 To UBound(h)
        If SqueezeHead(CStr(h(i))) = squeezed Then HeadNumber = i + 1: Exit Function
    Next
End Function

Private Function SqueezeHead(ByVal s As String) As String
    s = UCase$(s)
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, """", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr$(160), "")
    SqueezeHead = s
End Function

Private Function RateOf(ByVal v As Variant) As Variant
    ' A rate as a fraction, 0.0775 for 7.75%, however the cell has it: 0.0775, 7.75, "7.750%".
    ' Empty when there's no rate.
    Dim s As String, pct As Boolean, x As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(Replace(v, Chr$(160), " "))
        pct = (InStr(s, "%") > 0)
        s = Trim$(Replace(Replace(s, "%", ""), ",", ""))
        If Not IsPlainNumber(s) Then Exit Function
        x = Val(s)
    ElseIf IsNumeric(v) Then
        x = CDbl(v)
    Else
        Exit Function
    End If
    If pct Or x > 1 Then x = x / 100
    RateOf = Round(x, 7)
End Function

Private Function DateInText(ByVal s As String) As Date
    ' A date in a file or tab name, the ways the state writes them: 10-1-26, 04-01-26,
    ' 2026-10-01, Oct2025, October 2026, Jan2026. 0 when there isn't one.
    Dim i As Long, j As Long, k As Long, p(1 To 3) As String, ok As Boolean, fresh As Boolean
    Dim m As Long, d As Long, y As Long, months As Variant, u As String, at As Long, nm As Variant

    ' Three numbers with dashes between them
    s = Replace(Replace(s, "_", "-"), ".", "-")
    For i = 1 To Len(s)
        fresh = IsDigit(Mid$(s, i, 1))
        If fresh And i > 1 Then fresh = Not IsDigit(Mid$(s, i - 1, 1))
        If fresh Then
            p(1) = "": p(2) = "": p(3) = ""
            j = i
            k = 1
            Do
                Do While IsDigit(Mid$(s, j, 1))
                    p(k) = p(k) & Mid$(s, j, 1)
                    j = j + 1
                Loop
                If k = 3 Then Exit Do
                If Mid$(s, j, 1) <> "-" Or Not IsDigit(Mid$(s, j + 1, 1)) Then Exit Do
                j = j + 1
                k = k + 1
            Loop
            If k = 3 And Not IsDigit(Mid$(s, j, 1)) Then
                If Len(p(1)) = 4 Then
                    ok = (Len(p(2)) <= 2 And Len(p(3)) <= 2)
                    y = Val(p(1)): m = Val(p(2)): d = Val(p(3))
                Else
                    ok = (Len(p(1)) <= 2 And Len(p(2)) <= 2 And (Len(p(3)) = 2 Or Len(p(3)) = 4))
                    m = Val(p(1)): d = Val(p(2)): y = Val(p(3))
                    If Len(p(3)) = 2 Then y = 2000 + y
                End If
                If ok And m >= 1 And m <= 12 And d >= 1 And d <= 31 And y >= 2000 And y <= 2099 Then
                    DateInText = DateSerial(y, m, d)
                    Exit Function
                End If
            End If
        End If
    Next

    ' A month's name with a year after it
    u = UCase$(s)
    months = Array("JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", _
        "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER")
    For m = 1 To 12
        For Each nm In Array(months(m - 1), "SEPT", Left$(months(m - 1), 3))
            If nm <> "SEPT" Or m = 9 Then
                at = InStr(u, nm)
                Do While at > 0
                    j = at + Len(nm)
                    Do While Mid$(u, j, 1) = " " Or Mid$(u, j, 1) = "-"
                        j = j + 1
                    Loop
                    If Mid$(u, j, 4) Like "####" And Not IsDigit(Mid$(u, j + 4, 1)) Then
                        y = Val(Mid$(u, j, 4))
                        If y >= 2000 And y <= 2099 Then DateInText = DateSerial(y, m, 1): Exit Function
                    End If
                    at = InStr(at + 1, u, nm)
                Loop
            End If
        Next
    Next
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (ch Like "#")
End Function


' ---- comparing -------------------------------------------------------------------------

Private Function OnePerPlace(ByRef recs As Variant, ByVal n As Long, ByRef kept As Variant, _
                             ByVal idx As Object, ByRef dups As Long) As Long
    ' The list with each place once, the first time it shows. idx gets each place's key.
    Dim i As Long, c As Long, k As String, cnt As Long
    dups = 0
    If n = 0 Then Exit Function
    ReDim kept(1 To n, 1 To NC)
    For i = 1 To n
        k = PlaceKey(CStr(recs(i, C_PLACE)))
        If k <> "" Then
            If Found(idx, k) > 0 Then
                dups = dups + 1
            Else
                cnt = cnt + 1
                idx.Add cnt, k
                For c = 1 To NC
                    kept(cnt, c) = recs(i, c)
                Next
            End If
        End If
    Next
    OnePerPlace = cnt
End Function

Private Function PlaceKey(ByVal s As String) As String
    ' A place's name with only its letters and numbers: St. Helena and ST HELENA match, and so
    ' does a La Canada whose n came through as other letters.
    Dim i As Long, ch As String, out As String
    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next
    PlaceKey = out
End Function

Private Function CountyKey(ByVal s As String) As String
    ' Los Angeles and LOS ANGELES COUNTY are the same county
    s = PlaceKey(s)
    If Len(s) > 6 And Right$(s, 6) = "COUNTY" Then s = Left$(s, Len(s) - 6)
    CountyKey = s
End Function

Private Function Found(ByVal idx As Object, ByVal k As String) As Long
    On Error Resume Next
    Found = idx(k)
    If Err.Number <> 0 Then Found = 0
    On Error GoTo 0
End Function

Private Sub Compare(ByRef oldRecs As Variant, ByVal nOld As Long, ByVal oldIdx As Object, _
                    ByRef newRecs As Variant, ByVal nNew As Long, ByVal newIdx As Object, _
                    ByVal starts As Date, ByVal changes As Object)
    ' Every place whose rate is different in the new list, or that's new or gone. A place whose
    ' county changed shows too.
    Dim i As Long, j As Long, what As String, look As String, notes As String
    Dim o As Variant, nw As Variant

    For i = 1 To nNew
        j = Found(oldIdx, PlaceKey(CStr(newRecs(i, C_PLACE))))
        If j = 0 Then
            changes.Add MakeChange("New", "new", CStr(newRecs(i, C_NOTES)), oldRecs, 0, newRecs, i, starts)
        Else
            what = "": look = "": notes = ""
            o = oldRecs(j, C_RATE)
            nw = newRecs(i, C_RATE)
            If IsEmpty(o) And Not IsEmpty(nw) Then
                what = "Has a rate now": look = "new"
            ElseIf IsEmpty(nw) And Not IsEmpty(o) Then
                what = "No rate now": look = "gone": notes = CStr(newRecs(i, C_NOTES))
            ElseIf Not IsEmpty(o) Then
                If Abs(CDbl(nw) - CDbl(o)) > 0.0000001 Then
                    If CDbl(nw) > CDbl(o) Then what = "Went up": look = "up" Else what = "Went down": look = "down"
                End If
            End If
            If CountyKey(CStr(oldRecs(j, C_COUNTY))) <> CountyKey(CStr(newRecs(i, C_COUNTY))) Then
                notes = Trim$(notes & " County was " & OrBlank(CStr(oldRecs(j, C_COUNTY))) & ".")
                If what = "" Then what = "County changed": look = "other"
            End If
            If what <> "" Then changes.Add MakeChange(what, look, notes, oldRecs, j, newRecs, i, starts)
        End If
    Next

    ' Gone: in the old list but not the new
    For j = 1 To nOld
        If Found(newIdx, PlaceKey(CStr(oldRecs(j, C_PLACE)))) = 0 Then
            changes.Add MakeChange("Gone", "gone", "", oldRecs, j, newRecs, 0, starts)
        End If
    Next
End Sub

Private Function MakeChange(ByVal what As String, ByVal look As String, ByVal notes As String, _
                            ByRef oldRecs As Variant, ByVal j As Long, ByRef newRecs As Variant, _
                            ByVal i As Long, ByVal starts As Date) As Variant
    ' One line of the list. i is the place's row in the new list and j in the old, 0 for none.
    Dim f(0 To 9) As Variant, place As String, county As String
    If i > 0 Then
        place = newRecs(i, C_PLACE): county = newRecs(i, C_COUNTY)
    Else
        place = oldRecs(j, C_PLACE): county = oldRecs(j, C_COUNTY)
    End If
    f(F_PLACE) = place
    f(F_COUNTY) = county
    f(F_WHAT) = what
    If j > 0 Then f(F_OLD) = oldRecs(j, C_RATE)
    If i > 0 Then f(F_NEW) = newRecs(i, C_RATE)
    f(F_START) = starts
    f(F_NOTES) = notes
    f(F_LOOK) = look
    ' The same change seen again, by a later file, has the same ID
    f(F_ID) = PlaceKey(place) & "|" & what & "|" & Format$(starts, "yyyymmdd") & "|" & RateText(f(F_NEW))
    f(F_SORT) = UCase$(place) & Chr$(1) & UCase$(county)
    MakeChange = f
End Function

Private Function SortedOrder(ByVal changes As Object) As Variant
    ' The changes by place name, then county
    Dim n As Long, idx As Variant, keys As Variant, ch As Variant
    Dim gap As Long, i As Long, j As Long, t As Long
    n = changes.Count
    If n = 0 Then Exit Function
    ReDim idx(1 To n)
    ReDim keys(1 To n)
    For i = 1 To n
        idx(i) = i
        ch = changes(i)
        keys(i) = ch(F_SORT)
    Next
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            t = idx(i)
            j = i
            Do While j > gap
                If StrComp(keys(idx(j - gap)), keys(t), vbBinaryCompare) <= 0 Then Exit Do
                idx(j) = idx(j - gap)
                j = j - gap
            Loop
            idx(j) = t
        Next
        gap = gap \ 2
    Loop
    SortedOrder = idx
End Function

Private Function Tally(ByVal changes As Object) As String
    ' "3 went up, 1 went down, 2 new" and so on
    Dim looks As Variant, words As Variant, cnt(0 To 4) As Long, ch As Variant, i As Long, s As String
    looks = Array("up", "down", "new", "gone", "other")
    words = Array("went up", "went down", "new", "gone", "other")
    For Each ch In changes
        For i = 0 To 4
            If ch(F_LOOK) = looks(i) Then cnt(i) = cnt(i) + 1
        Next
    Next
    For i = 0 To 4
        If cnt(i) > 0 Then
            If s <> "" Then s = s & ", "
            s = s & Format$(cnt(i), "#,##0") & " " & words(i)
        End If
    Next
    Tally = s
End Function


' ---- the sheets ------------------------------------------------------------------------

Private Sub EnsureSheets()
    ' First run: makes the three sheets. Later runs only put back the button.
    Dim wb As Workbook, ws As Worksheet, main As Worksheet, i As Long
    Set wb = ThisWorkbook
    Set main = SheetByName(SH_MAIN)
    If main Is Nothing Then
        For Each ws In wb.Worksheets
            If ws.Name Like "Sheet#*" Then
                If IsBlankSheet(ws) Then Set main = ws: Exit For
            End If
        Next
        If main Is Nothing Then Set main = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        main.Name = SH_MAIN
        If main.Index <> 1 Then main.Move Before:=wb.Worksheets(1)
        DrawWelcome main
    End If
    If SheetByName(SH_HOWTO) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_MAIN))
        ws.Name = SH_HOWTO
        ThisWorkbook.Activate
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    End If
    If SheetByName(SH_LOG) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_HOWTO))
        ws.Name = SH_LOG
        LogHeader ws
    End If
    If SheetByName(SH_SAVED) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SH_SAVED
        ws.Range("A1").Value = "Nothing saved yet. Loading a file puts its rates here."
        ws.Range("A1").Font.Size = 14
    End If
    ' Empty starter sheets from a new workbook go
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If ws.Name Like "Sheet#*" And wb.Worksheets.Count > 1 Then
            If IsBlankSheet(ws) Then ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
    DrawHowTo
    DrawButton SheetByName(SH_MAIN)
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
End Sub

Private Sub DrawWelcome(ByVal ws As Worksheet)
    ClearSheet ws
    ListWidths ws
    ws.Cells.Font.Size = 14
    With ws.Range("A1")
        .Value = TITLE_TEXT
        .Font.Size = 26
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 50
    With ws.Range("A3")
        .Value = "Click the button and pick the rates list."
        .Font.Size = 20
    End With
    ws.Rows(3).RowHeight = 34
    ws.Range("A4").Value = "Get it from " & WHERE_TEXT & "."
    ws.Rows(4).RowHeight = 24
    ws.Range("A5").Value = "The How to use tab has the steps."
    ws.Rows(5).RowHeight = 24
    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
    DrawButton ws
End Sub

Private Sub DrawHowTo()
    ' How to use: the steps, kept in the macro so they travel with it. Redrawn on every click,
    ' so a newer version of the macro brings its own steps.
    Dim ws As Worksheet, r As Long
    Set ws = SheetByName(SH_HOWTO)
    If ws Is Nothing Then Exit Sub
    ClearSheet ws
    ws.Columns(1).ColumnWidth = 120
    ws.Columns(1).NumberFormat = "@"
    ws.Cells.Font.Size = 16
    r = 1
    HowToLine ws, r, "T", "How to use this"
    HowToLine ws, r, "P", "This shows what changed in California's city and county sales and use tax " & _
        "rates, each time the state posts a new list. Each rate is the whole rate, state and local together."
    HowToLine ws, r, "H", "Each quarter"
    HowToLine ws, r, "P", "California's rates change only on January 1, April 1, July 1 and October 1. " & _
        "The state posts the new list a month before, on the 1st of March, June, September and December."
    HowToLine ws, r, "N", "1.  Go to cdtfa.ca.gov, See Tax Rates, Sales and Use Tax Rates. Under Download, " & _
        "click Excel Version 2012 (xlsx) for the new date. It saves as something like " & _
        "SalesTaxRates1-1-27.xlsx. No need to open it."
    HowToLine ws, r, "N", "2.  Come back to this workbook and click Check new rates file on the What " & _
        "changed sheet. The box opens in Downloads. Pick the file you just downloaded."
    HowToLine ws, r, "N", "3.  What changed shows each place whose rate changes, with the old rate, the " & _
        "new rate and the day it starts. Red went up, green went down, blue is new, gray is gone. " & _
        "NO CHANGES means every place has the same rate as the last list."
    HowToLine ws, r, "N", "4.  Save the workbook. It never saves itself."
    HowToLine ws, r, "H", "Good to know"
    HowToLine ws, r, "B", "The first list loaded only saves the rates. Every list after that is checked " & _
        "against the one before."
    HowToLine ws, r, "B", "From the 1st of the month before a quarter, a yellow line at the top of What " & _
        "changed says the new list should be out. Once the quarter starts, it turns red: GET THE NEW FILE."
    HowToLine ws, r, "B", "The Log keeps every change once, plus a line for each check."
    HowToLine ws, r, "B", "Saved rates has every place's rate from the last list, with filter buttons, so " & _
        "it's a quick way to look one up. Loading a file replaces it, so leave it alone otherwise."
    HowToLine ws, r, "H", "If something comes up"
    HowToLine ws, r, "B", "If a yellow bar says macros are disabled, click Enable Content."
    HowToLine ws, r, "B", "If it says the list looks older than last time, or looks cut short, answer No " & _
        "and download the file again."
    HowToLine ws, r, "B", "If it asks what day the rates start, the file's name didn't say. Type the " & _
        "quarter's first day, like 1/1/2027."
    HowToLine ws, r, "B", "If it says another file with the same name is open, close that one and click " & _
        "the button again."
    HowToLine ws, r, "B", "For a newer version of the macro: Alt F11, click into the module, Ctrl A, paste " & _
        "the new one over it, close that window and save. Saved rates and the Log stay."
    HowToLine ws, r, "S", "This sheet is redrawn each time you click the button, so notes typed here won't stay."
End Sub

Private Sub HowToLine(ByVal ws As Worksheet, ByRef r As Long, ByVal kind As String, ByVal words As String)
    ' One line of How to use: T title, H heading, P plain, N numbered step, B bullet, S small print
    If kind = "H" Or kind = "S" Then
        ws.Rows(r).RowHeight = 14
        r = r + 1
    End If
    If kind = "B" Then words = ChrW$(8226) & "  " & words
    With ws.Cells(r, 1)
        .Value = words
        .WrapText = True
        .VerticalAlignment = xlTop
        Select Case kind
            Case "T"
                .Font.Size = 26
                .Font.Bold = True
            Case "H"
                .Font.Size = 20
                .Font.Bold = True
                .Font.Color = RGB(31, 78, 121)
            Case "N", "B"
                .IndentLevel = 1
            Case "S"
                .Font.Size = 14
                .Font.Color = RGB(89, 89, 89)
        End Select
    End With
    ws.Rows(r).AutoFit
    If kind = "T" Then ws.Rows(r).RowHeight = 50
    r = r + 1
End Sub

Private Sub DrawFront(ByVal changes As Object, ByVal compared As Boolean, ByVal fname As String, _
                      ByVal loadedOn As Date, ByVal starts As Date, ByVal typed As Boolean, _
                      ByVal savedFile As String, ByVal savedWhen As Variant, ByVal savedStarts As Date, _
                      ByVal places As Long, ByVal dups As Long)
    ' What changed: which files, the count in big type, then the list.
    Dim ws As Worksheet, r As Long, n As Long, s As String, skipped As String, lines As Collection
    Dim ln As Variant

    Set ws = SheetByName(SH_MAIN)
    ClearSheet ws
    ListWidths ws
    ws.Cells.Font.Size = 14
    ws.Cells.VerticalAlignment = xlCenter
    With ws.Range("A1")
        .Value = TITLE_TEXT
        .Font.Size = 26
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 50
    s = "From " & fname & ", loaded " & ShowDate(loadedOn) & ". Its rates start " & ShowDate(starts)
    If typed Then s = s & " (the day you typed; the file doesn't say)"
    ws.Range("A2").Value = s & "."
    If compared Then
        s = "Checked against " & savedFile & ", loaded " & ShowWhen(savedWhen) & _
            ", with rates starting " & ShowDate(savedStarts) & "."
    Else
        s = "First file, so there's nothing to check it against yet."
    End If
    ws.Range("A3").Value = s
    ws.Range("A3").Font.Color = RGB(89, 89, 89)
    ws.Rows(2).RowHeight = 24
    ws.Rows(3).RowHeight = 24
    ' Row 4 is the get-the-new-file line, drawn by DrawAlert below
    r = ALERT_ROW + 1
    ws.Rows(r).RowHeight = 10
    r = r + 1

    n = changes.Count
    Set lines = New Collection
    If Not compared Then
        Band ws, r, "NOTHING TO COMPARE YET", RGB(221, 235, 247), RGB(31, 78, 121), 28
        s = "The rates for " & Format$(places, "#,##0") & " places are saved. The next file gets checked against them."
    ElseIf n = 0 Then
        Band ws, r, "NO CHANGES", RGB(198, 239, 206), RGB(0, 97, 0), 28
        s = "Every place has the same rate as in the last file."
    Else
        If n = 1 Then s = "1 CHANGE" Else s = Format$(n, "#,##0") & " CHANGES"
        Band ws, r, s & " ON " & UCase$(Format$(starts, "mmmm d, yyyy")), RGB(255, 217, 102), RGB(0, 0, 0), 28
        s = Tally(changes) & "."
        s = UCase$(Left$(s, 1)) & Mid$(s, 2)
    End If
    r = r + 1
    ws.Cells(r, 1).Value = s
    ws.Cells(r, 1).Font.Size = 16
    ws.Rows(r).RowHeight = 28
    r = r + 1

    ' Anything odd about this pair of files, one line each
    If compared Then
        If starts = savedStarts And n > 0 Then
            lines.Add "This list starts the same day as the last one, so these are fixes the state made to it."
        ElseIf starts < savedStarts Then
            lines.Add "This list is older than the last one, so these go back to its rates."
        End If
        skipped = QuartersBetween(savedStarts, starts)
        If skipped <> "" Then
            lines.Add "The last list loaded was for " & ShowDate(savedStarts) & ", so this one covers more " & _
                "than a quarter. Some of these may have started " & skipped & "."
        End If
    End If
    If dups = 1 Then lines.Add "The file lists 1 place twice; the first one was used."
    If dups > 1 Then lines.Add "The file lists " & dups & " places twice; the first of each was used."
    For Each ln In lines
        ws.Cells(r, 1).Value = ln
        ws.Cells(r, 1).Font.Color = RGB(89, 89, 89)
        ws.Rows(r).RowHeight = 24
        r = r + 1
    Next

    If n > 0 Then
        ws.Rows(r).RowHeight = 10
        r = r + 1
        DrawList ws, r, changes
    End If
    DrawAlert ws

    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
    DrawButton ws
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo 0
End Sub

Private Sub DrawAlert(ByVal ws As Worksheet)
    ' The line under the file names. The state posts each quarter's list on the 1st of the
    ' month before it starts, so from then on this says to get it, in red once it's started.
    Dim starts As Date, today As Date, due As Date, s As String
    starts = SavedStartDay()
    If starts = 0 Then Exit Sub                 ' nothing loaded yet
    With ws.Range(ws.Cells(ALERT_ROW, 1), ws.Cells(ALERT_ROW, LIST_COLS))
        .UnMerge
        .Clear
    End With
    ws.Rows(ALERT_ROW).RowHeight = 8
    today = CheckDate()
    due = QuarterGuess(today)
    If today >= NextQuarter(starts) Then
        s = "GET THE NEW FILE. THE " & UCase$(Format$(due, "mmmm d, yyyy")) & " LIST IS OUT."
        Band ws, ALERT_ROW, s, RGB(255, 199, 206), RGB(156, 0, 6), 24
    ElseIf due > starts Then
        s = "THE " & UCase$(Format$(due, "mmmm d, yyyy")) & " LIST SHOULD BE OUT. GET THE NEW FILE."
        Band ws, ALERT_ROW, s, LookFill("other"), LookText("other"), 20
    End If
End Sub

Private Sub Band(ByVal ws As Worksheet, ByVal r As Long, ByVal words As String, _
                 ByVal fillColor As Long, ByVal textColor As Long, ByVal pts As Long)
    ' A colored line across the list, for the things that need to be seen
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, LIST_COLS))
        .Merge
        .Value = words
        .Font.Size = pts
        .Font.Bold = True
        .Font.Color = textColor
        .Interior.Color = fillColor
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(r).RowHeight = pts * 1.8
End Sub

Private Sub DrawList(ByVal ws As Worksheet, ByRef r As Long, ByVal changes As Object)
    ' The header row at r, then one row per change. r ends on the row after the list.
    Dim n As Long, order As Variant, out As Variant, ch As Variant, titles As Variant
    Dim i As Long, c As Long, rr As Long
    n = changes.Count
    titles = Split(LIST_HEADS, "|")
    For c = 0 To UBound(titles)
        ws.Cells(r, c + 1).Value = titles(c)
    Next
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, LIST_COLS))
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .WrapText = True
    End With
    ws.Rows(r).RowHeight = 44
    order = SortedOrder(changes)
    ReDim out(1 To n, 1 To LIST_COLS)
    For i = 1 To n
        ch = changes(order(i))
        For c = 1 To LIST_COLS
            out(i, c) = ch(c)
        Next
    Next
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 2)).NumberFormat = "@"
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, LIST_COLS))
        .Value = out
        .Font.Size = 16
    End With
    ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + n, 5)).NumberFormat = RATE_LOOK
    ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + n, 5)).Font.Bold = True
    ws.Range(ws.Cells(r + 1, 6), ws.Cells(r + n, 6)).NumberFormat = "m/d/yyyy"
    For i = 1 To n
        ch = changes(order(i))
        With ws.Range(ws.Cells(r + i, 1), ws.Cells(r + i, LIST_COLS)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
        With ws.Cells(r + i, 3)
            .Interior.Color = LookFill(CStr(ch(F_LOOK)))
            .Font.Color = LookText(CStr(ch(F_LOOK)))
            .Font.Bold = True
        End With
    Next
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).WrapText = True
    ws.Range(ws.Cells(r + 1, LIST_COLS), ws.Cells(r + n, LIST_COLS)).WrapText = True
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).EntireRow.AutoFit
    For rr = r + 1 To r + n
        If ws.Rows(rr).RowHeight < 30 Then ws.Rows(rr).RowHeight = 30
    Next
    r = r + n + 1
End Sub

Private Sub ListWidths(ByVal ws As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(34, 18, 20, 12, 12, 14, 60)
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next
End Sub

Private Sub AddToLog(ByVal changes As Object, ByVal compared As Boolean, ByVal fname As String, _
                     ByVal loadedOn As Date, ByVal places As Long)
    ' Each change goes on the Log once, the first time a file shows it. Every check gets a line.
    Dim ws As Worksheet, r As Long, last As Long, ids As Collection, v As Variant, i As Long
    Dim out As Variant, n As Long, order As Variant, ch As Variant, c As Long, looks() As String

    Set ws = SheetByName(SH_LOG)
    If ws.FilterMode Then ws.ShowAllData
    If CStr(ws.Cells(1, 1).Value) = "" Then LogHeader ws
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set ids = New Collection
    If last >= 2 Then
        v = ws.Range(ws.Cells(2, LOG_COLS), ws.Cells(last, LOG_COLS)).Value
        On Error Resume Next
        If IsArray(v) Then
            For i = 1 To UBound(v, 1)
                If CStr(v(i, 1)) <> "" Then ids.Add 1, CStr(v(i, 1))
            Next
        ElseIf CStr(v) <> "" Then
            ids.Add 1, CStr(v)
        End If
        On Error GoTo 0
    End If

    ReDim out(1 To changes.Count + 1, 1 To LOG_COLS)
    ReDim looks(0 To changes.Count + 1)
    If changes.Count > 0 Then
        order = SortedOrder(changes)
        For i = 1 To changes.Count
            ch = changes(order(i))
            If Found(ids, CStr(ch(F_ID))) = 0 Then
                ids.Add 1, CStr(ch(F_ID))
                n = n + 1
                out(n, 1) = loadedOn
                out(n, 2) = fname
                For c = 1 To LIST_COLS
                    out(n, c + 2) = ch(c)
                Next
                out(n, LOG_COLS) = ch(F_ID)
                looks(n) = CStr(ch(F_LOOK))
            End If
        Next
    End If
    If n = 0 Then
        n = 1
        out(1, 1) = loadedOn
        out(1, 2) = fname
        If compared Then
            out(1, 5) = "Nothing new"
            out(1, 9) = "Everything this file shows was already on the Log."
        Else
            out(1, 5) = "First file"
            out(1, 9) = Format$(places, "#,##0") & " places saved. Nothing to compare yet."
        End If
        looks(1) = ""
    End If

    r = last + 1
    If r < 2 Then r = 2
    ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 4)).NumberFormat = "@"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, LOG_COLS))
        .Value = out
        .Font.Size = 14
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 1)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 6), ws.Cells(r + n - 1, 7)).NumberFormat = RATE_LOOK
    ws.Range(ws.Cells(r, 8), ws.Cells(r + n - 1, 8)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 9), ws.Cells(r + n - 1, 9)).WrapText = True
    For i = 1 To n
        If looks(i) <> "" Then
            With ws.Cells(r + i - 1, 5)
                .Interior.Color = LookFill(looks(i))
                .Font.Color = LookText(looks(i))
            End With
        End If
    Next
    ' A line across between checks
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, LOG_COLS - 1)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With
    ws.Columns(LOG_COLS).Hidden = True
    ' Filter buttons on the header, covering every line
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(r + n - 1, LOG_COLS)).AutoFilter
End Sub

Private Sub LogHeader(ByVal ws As Worksheet)
    Dim titles As Variant, widths As Variant, c As Long
    titles = Split(LOG_HEADS, "|")
    widths = Array(13, 30, 30, 16, 18, 11, 11, 12, 50, 10)
    ws.Cells.Font.Size = 14
    For c = 0 To UBound(titles)
        ws.Cells(1, c + 1).Value = titles(c)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(titles) + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 40
    ws.Columns(LOG_COLS).Hidden = True
    FreezeAt ws, 1
End Sub

Private Sub SaveCopy(ByRef recs As Variant, ByVal n As Long, ByVal fname As String, _
                     ByVal loadedOn As Date, ByVal starts As Date)
    ' Saved rates: the list just loaded, for the next file to be checked against
    Dim ws As Worksheet, out As Variant, i As Long, c As Long, titles As Variant
    Set ws = SheetByName(SH_SAVED)
    ws.Unprotect
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.Font.Size = 14
    With ws.Range("A1")
        .Value = "The rates from the last list loaded. The next file gets checked against them. " & _
            "Loading a file replaces them, so leave this sheet alone."
        .Font.Bold = True
    End With
    ws.Range("A2").Value = "File"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = fname
    ws.Range("C2").Value = "Loaded"
    ws.Range("D2").Value = loadedOn
    ws.Range("D2").NumberFormat = "m/d/yyyy"
    ws.Range("E2").Value = "Rates start"
    ws.Range("F2").Value = starts
    ws.Range("F2").NumberFormat = "m/d/yyyy"
    ws.Range("G2").Value = "Places"
    ws.Range("H2").Value = n
    ws.Range("A2,C2,E2,G2").Font.Bold = True
    titles = Split(HEADS, "|")
    For c = 1 To NC
        ws.Cells(SAVED_ROW, c).Value = titles(c - 1)
    Next
    ws.Range(ws.Cells(SAVED_ROW, 1), ws.Cells(SAVED_ROW, NC)).Font.Bold = True
    ReDim out(1 To n, 1 To NC)
    For i = 1 To n
        For c = 1 To NC
            out(i, c) = recs(i, c)
            If c <> C_RATE And CStr(out(i, c)) = "" Then out(i, c) = Empty
        Next
    Next
    ' Words stay words, whatever they look like
    ws.Range(ws.Cells(SAVED_ROW + 1, C_PLACE), ws.Cells(SAVED_ROW + n, C_PLACE)).NumberFormat = "@"
    ws.Range(ws.Cells(SAVED_ROW + 1, C_COUNTY), ws.Cells(SAVED_ROW + n, C_NOTES)).NumberFormat = "@"
    ws.Range(ws.Cells(SAVED_ROW + 1, C_RATE), ws.Cells(SAVED_ROW + n, C_RATE)).NumberFormat = RATE_LOOK
    ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + n, NC)).Value = out
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 60
    ws.Range(ws.Cells(SAVED_ROW, 1), ws.Cells(SAVED_ROW + n, NC)).AutoFilter
    ws.Protect AllowFiltering:=True
End Sub

Private Function ReadSaved(ByRef recs As Variant, ByRef savedFile As String, ByRef savedWhen As Variant, _
                           ByRef savedStarts As Date) As Long
    ' Every place on Saved rates, the way ReadList gives them. 0 when nothing's saved yet.
    Dim ws As Worksheet, n As Long, i As Long, c As Long, v As Variant
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Cells(SAVED_ROW, 1).Value) <> "Location" Then Exit Function
    ' H2 has the count, so a filter left on the sheet doesn't cut the list short
    n = Val(CStr(ws.Range("H2").Value))
    If n <= 0 Then n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - SAVED_ROW
    If n <= 0 Then Exit Function
    v = ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + n, NC)).Value2
    ReDim recs(1 To n, 1 To NC)
    For i = 1 To n
        For c = 1 To NC
            If c = C_RATE Then recs(i, c) = RateOf(v(i, c)) Else recs(i, c) = CellText(v(i, c))
        Next
    Next
    savedFile = CStr(ws.Range("B2").Value)
    savedWhen = ws.Range("D2").Value
    savedStarts = SavedStartDay()
    ReadSaved = n
End Function

Private Function SavedStartDay() As Date
    ' The day the saved rates start, 0 when nothing's saved
    Dim ws As Worksheet, v As Variant
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    v = ws.Range("F2").Value
    If IsDate(v) Then SavedStartDay = CDate(v)
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.RowHeight = ws.StandardHeight
End Sub

Private Sub DrawButton(ByVal ws As Worksheet)
    ' The big button beside the title. Made once, then only moved back into place.
    Dim b As Object
    On Error Resume Next
    Set b = ws.Buttons(BTN_NAME)
    On Error GoTo 0
    If b Is Nothing Then
        Set b = ws.Buttons.Add(0, 0, 100, 30)
        b.Name = BTN_NAME
    End If
    b.Placement = xlFreeFloating
    b.Left = ws.Range("F1").Left + 10
    b.Top = ws.Range("A1").Top + 4
    b.Width = 290
    b.Height = ws.Range("A1").Height - 8
    b.Caption = BTN_TEXT
    b.OnAction = "CheckCaliforniaRates"
    On Error Resume Next
    b.Font.Size = 18
    b.Font.Bold = True
    On Error GoTo 0
End Sub

Private Sub FreezeAt(ByVal ws As Worksheet, ByVal rowsAbove As Long)
    ThisWorkbook.Activate
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        If rowsAbove > 0 Then
            .SplitColumn = 0
            .SplitRow = rowsAbove
            .FreezePanes = True
        End If
    End With
End Sub


' ---- small helpers -----------------------------------------------------------------------

Private Function SheetByName(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then Set SheetByName = ws: Exit Function
    Next
End Function

Private Function HasAnything(ByVal ws As Worksheet) As Boolean
    ' Anything at all on a sheet of the picked file. (CountA can't look into a file that's
    ' open in Protected View, so this asks the sheet itself.)
    Dim u As Range
    Set u = ws.UsedRange
    If u.Cells.CountLarge > 1 Then HasAnything = True Else HasAnything = Not IsEmpty(u.Value2)
End Function

Private Function IsBlankSheet(ByVal ws As Worksheet) As Boolean
    IsBlankSheet = (Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0)
End Function

Private Function QuarterGuess(ByVal d As Date) As Date
    ' The newest list the state has posted by day d: the quarter d is in, or the next one from
    ' the 1st of the month before it starts.
    Dim q As Date
    q = DateSerial(Year(d), ((Month(d) - 1) \ 3) * 3 + 1, 1)
    If Month(d) Mod 3 = 0 Then q = DateSerial(Year(q), Month(q) + 3, 1)
    QuarterGuess = q
End Function

Private Function NextQuarter(ByVal d As Date) As Date
    ' The first quarter start after day d
    NextQuarter = DateSerial(Year(d), ((Month(d) - 1) \ 3) * 3 + 4, 1)
End Function

Private Function QuartersBetween(ByVal fromDay As Date, ByVal toDay As Date) As String
    ' The quarter starts after fromDay and before toDay, in words: "July 1, 2026 or October 1, 2026"
    Dim q As Date, s As String, n As Long
    q = NextQuarter(fromDay)
    Do While q < toDay
        n = n + 1
        If n > 4 Then QuartersBetween = "in between": Exit Function
        If s <> "" Then s = s & " or "
        s = s & Format$(q, "mmmm d, yyyy")
        q = NextQuarter(q)
    Loop
    QuartersBetween = s
End Function

Private Function LookFill(ByVal look As String) As Long
    Select Case look
        Case "up": LookFill = RGB(255, 199, 206)
        Case "down": LookFill = RGB(198, 239, 206)
        Case "new": LookFill = RGB(189, 215, 238)
        Case "gone": LookFill = RGB(217, 217, 217)
        Case Else: LookFill = RGB(255, 235, 156)
    End Select
End Function

Private Function LookText(ByVal look As String) As Long
    Select Case look
        Case "up": LookText = RGB(156, 0, 6)
        Case "down": LookText = RGB(0, 97, 0)
        Case "new": LookText = RGB(31, 78, 121)
        Case "gone": LookText = RGB(64, 64, 64)
        Case Else: LookText = RGB(156, 87, 0)
    End Select
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

Private Function RateText(ByVal v As Variant) As String
    ' A rate the same way however it was read: 0.0775 and 0.077500 both come out 0.0775
    If IsEmpty(v) Then Exit Function
    RateText = Trim$(Str$(Round(CDbl(v), 7)))
End Function

Private Function OrBlank(ByVal s As String) As String
    If Trim$(s) = "" Then OrBlank = "blank" Else OrBlank = Trim$(s)
End Function

Private Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long, digits As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
        ElseIf ch = "-" And i = 1 Then
            ' a minus sign up front is fine
        Else
            Exit Function
        End If
    Next
    IsPlainNumber = (digits > 0 And dots <= 1)
End Function

Private Function ShowDate(ByVal d As Date) As String
    ShowDate = Format$(d, "m/d/yyyy")
End Function

Private Function ShowWhen(ByVal v As Variant) As String
    If IsDate(v) Then ShowWhen = Format$(CDate(v), "m/d/yyyy") Else ShowWhen = "before"
End Function

Private Function FileNameOf(ByVal fullPath As String) As String
    FileNameOf = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Private Function CheckDate() As Date
    If mToday <> 0 Then CheckDate = mToday Else CheckDate = Date
End Function


' ---- the boxes -------------------------------------------------------------------------

Private Function PickFile() As Variant
    ' The file picked, or False if the box was cancelled. Starts in Downloads when there is one.
    Dim dl As String
    If Not mFeed Is Nothing Then
        If mFeed.Count = 0 Then
            PickFile = False
        Else
            PickFile = mFeed(1)
            mFeed.Remove 1
            If PickFile = "<cancel>" Then PickFile = False
        End If
        Exit Function
    End If
    dl = Environ$("USERPROFILE") & "\Downloads"
    On Error Resume Next
    If Len(Dir$(dl, vbDirectory)) > 0 Then
        ChDrive Left$(dl, 1)
        ChDir dl
    End If
    On Error GoTo 0
    PickFile = Application.GetOpenFilename("Rates list (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", 1, _
        "Pick the new California rates list (SalesTaxRates...xlsx)")
End Function

Private Function AskDate(ByVal guess As Date) As Date
    ' The day the rates start, when the file's name doesn't say. 0 when the box is cancelled.
    Dim a As String, q As String
    q = "What day do these rates start? The file's name doesn't say." & vbLf & vbLf & _
        "The state's lists start on the 1st of January, April, July or October. Type it like " & _
        ShowDate(guess) & "."
    If Not mFeed Is Nothing Then
        mSaid = mSaid & q & vbLf
        If mFeed.Count > 0 Then
            a = mFeed(1)
            mFeed.Remove 1
        End If
    Else
        a = InputBox(q, BOX_TITLE, ShowDate(guess))
    End If
    a = Trim$(a)
    If a = "" Or a = "<cancel>" Then Exit Function
    If Not IsDate(a) Then
        Say "That isn't a date, so nothing was changed. Click the button to try again."
        Exit Function
    End If
    AskDate = DateValue(a)
End Function

Private Sub Say(ByVal msg As String)
    If Not mFeed Is Nothing Then
        mSaid = mSaid & msg & vbLf
    Else
        MsgBox msg, vbInformation, BOX_TITLE
    End If
End Sub

Private Function Confirm(ByVal question As String) As Boolean
    Dim a As Variant
    If Not mFeed Is Nothing Then
        mSaid = mSaid & question & vbLf
        If mFeed.Count > 0 Then
            a = mFeed(1)
            mFeed.Remove 1
        End If
        Confirm = (LCase$(CStr(a)) = "yes")
        Exit Function
    End If
    Confirm = (MsgBox(question, vbYesNo + vbQuestion + vbDefaultButton2, BOX_TITLE) = vbYes)
End Function

Public Function CaRatesTest(ByVal answers As String, ByVal todayText As String) As String
    ' For test runs: feeds answers to the boxes (split on |), acts as if today were todayText
    ' (yyyy-mm-dd, blank for the real day), and returns what the boxes would have said.
    ' "<open>" as the answers does what opening the workbook does instead of the button.
    ' Not needed day to day.
    Dim a As Variant
    Set mFeed = New Collection
    mSaid = ""
    If todayText <> "" Then
        mToday = DateSerial(Val(Left$(todayText, 4)), Val(Mid$(todayText, 6, 2)), Val(Mid$(todayText, 9, 2)))
    Else
        mToday = 0
    End If
    If answers = "<open>" Then
        Auto_Open
    Else
        If answers <> "" Then
            For Each a In Split(answers, "|")
                mFeed.Add CStr(a)
            Next
        End If
        CheckCaliforniaRates
    End If
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    CaRatesTest = mSaid
    Set mFeed = Nothing
    mToday = 0
End Function
