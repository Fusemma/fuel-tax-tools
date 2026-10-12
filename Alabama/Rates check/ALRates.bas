' Alabama rates check: what changes in the state's local rates, by the month it starts.
'
' The Alabama DOR posts every city and county sales, use, lodgings and rental rate in one
' file, updated monthly, with the whole history in it. The What changed sheet shows the
' rates you watch (set just below these notes) that changed going into this month and
' going into next month: in September, August to September and September to October.
' Opening the workbook moves it on when the month turns.
'
' Each month:
'   1. Download taxrates.csv from the DOR's Local Cities and Counties Tax Rates Text File
'      page (revenue.alabama.gov, under Sales & Use). It's the first link, the big file
'      with the history. The small taxrates_current.csv won't do.
'   2. Click Check new rates file on the What changed sheet and pick the file.
'   3. Save the workbook if you want to keep it. It never saves itself.
' Since last file lists what changed between this file and the last one that the month
' lists don't show, like a row the state added or fixed with an old start date, or a
' month nobody checked. Log keeps each change once, the first time a file shows it, plus
' a line for each check with nothing new. Saved rates holds the file's rates for next time.
'
' First time: paste this into a module in a new workbook, save it as a macro-enabled
' workbook (.xlsm), then Alt F8 and run CheckNewRatesFile. It sets up the sheets and the
' button, then asks for the file.
'
' A change is a rate that's new, gone, went up or down (the PJ rate too), moved to a new
' administrator, or got a new name. A row the state only re-dated isn't a change.

Option Explicit

' ---- the rates you watch ----------------------------------------------------------
' Separate with commas. Tax types in the file: ST sales, SU sellers use, CU consumers use,
' LL lodgings, RT rental. Rate types: GENER general, AUTO automotive, FARM farm, MACH
' machine, GROC food, VEND vending, AMUSE amusement, WDFEE withdrawal fee, and a few more.
Private Const WATCH_TAX As String = "SU"
Private Const WATCH_RATE As String = "GENER"

' ---- the rest -----------------------------------------------------------------------
Private Const SH_MAIN As String = "What changed"
Private Const SH_SINCE As String = "Since last file"
Private Const SH_LOG As String = "Log"
Private Const SH_SAVED As String = "Saved rates"
Private Const BTN_NAME As String = "CheckRatesButton"
Private Const BTN_TEXT As String = "Check new rates file"
Private Const BOX_TITLE As String = "Alabama rates"
Private Const TITLE_TEXT As String = "Alabama local rates: what changed"
Private Const SINCE_TITLE As String = "Changed since the last file"

' The state's columns, in the order Saved rates keeps them. The file's columns are found
' by name, so a file with its columns moved around still reads right.
Private Const HEADS As String = "Locality Code|Locality Name|County Number|TaxType|Rate Type|" & _
    "Administered|Active Date|Inactive Date|Rate|Indicator|PJ|County Code|PJ_Rate"
Private Const NC As Long = 13
Private Const C_CODE As Long = 1, C_NAME As Long = 2, C_CNUM As Long = 3, C_TAX As Long = 4
Private Const C_RTYPE As Long = 5, C_ADMIN As Long = 6, C_START As Long = 7, C_END As Long = 8
Private Const C_RATE As Long = 9, C_IND As Long = 10, C_PJ As Long = 11, C_CCODE As Long = 12
Private Const C_PJRATE As Long = 13

Private Const LIST_HEADS As String = "Locality|Tax|What changed|Old rate|New rate|Starts|" & _
    "Old PJ rate|New PJ rate|Code|County code|Notes"
Private Const LOG_HEADS As String = "Checked|File|Section|" & LIST_HEADS & "|Change ID"
Private Const LOG_COLS As Long = 15
Private Const SAVED_ROW As Long = 4         ' header row on Saved rates

' One change, as kept while sorting and drawing. 1 to 11 are the list's columns in order.
Private Const F_SORT As Long = 0, F_NAME As Long = 1, F_TAX As Long = 2, F_WHAT As Long = 3
Private Const F_OLD As Long = 4, F_NEW As Long = 5, F_START As Long = 6, F_OLDPJ As Long = 7
Private Const F_NEWPJ As Long = 8, F_CODE As Long = 9, F_CCODE As Long = 10, F_NOTES As Long = 11
Private Const F_RTYPE As Long = 12, F_LOOK As Long = 13, F_KEY As Long = 14, F_ID As Long = 15
Private Const F_WAS As Long = 16, F_NOW As Long = 17     ' the rate's details before and after

Private mFeed As Collection     ' test runs only: answers for the boxes, in order
Private mSaid As String         ' test runs only: what the boxes would have said
Private mToday As Date          ' test runs only: the day to act as today


' ---- the button, and opening the workbook ------------------------------------------

Public Sub CheckNewRatesFile()
    ' Asks for the new rates file, then shows what changed.
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
    ' Opening the workbook moves What changed on to this month and next. That redraw isn't
    ' anything to save, so the workbook is marked unchanged and closing won't ask.
    On Error Resume Next
    If SheetByName(SH_MAIN) Is Nothing Then Exit Sub
    RefreshFront
    ThisWorkbook.Saved = True
    Application.ScreenUpdating = True
End Sub

Private Sub LoadFile()
    Dim picked As Variant, fname As String, problem As String, hasHistory As Boolean
    Dim allRecs As Variant, nAll As Long, keepRecs As Variant, nKeep As Long
    Dim newCur As Variant, nNew As Long, newIdx As Collection, dups As Long
    Dim oldAll As Variant, nOldAll As Long, oldCur As Variant, nOld As Long, oldIdx As Collection
    Dim thisMonth As Collection, nextMonth As Collection, since As Collection, sinceAll As Collection
    Dim others As Long, sinceOthers As Long, today As Date, x As Long
    Dim newest As String, savedNewest As String, savedFile As String, savedWhen As Variant

    today = CheckDate()
    EnsureSheets
    picked = PickFile()
    If VarType(picked) = vbBoolean Then Exit Sub
    fname = FileNameOf(CStr(picked))

    Application.Cursor = xlWait
    Application.StatusBar = "Reading " & fname & "..."
    If Not ReadRates(CStr(picked), allRecs, nAll, hasHistory, problem) Then
        Application.Cursor = xlDefault
        Say problem
        Exit Sub
    End If
    If Not hasHistory Then
        Application.Cursor = xlDefault
        Say "That's the small file, with only today's rates in it, so it can't show what changed " & _
            "month to month." & vbLf & vbLf & "Pick taxrates.csv, the first link on the state's page, " & _
            "the big file with the history in it. Nothing was changed."
        Exit Sub
    End If

    ' The rates in force now, meaning every row with no end date
    Set newIdx = New Collection
    nNew = PickRows(allRecs, nAll, "", newCur, newIdx, dups)
    If nNew = 0 Then
        Application.Cursor = xlDefault
        Say "That file has no rates in force in it. Download taxrates.csv again and pick that." & _
            vbLf & vbLf & "Nothing was changed."
        Exit Sub
    End If

    ' The file loaded last time, if there was one
    nOldAll = ReadSaved(oldAll, savedFile, savedWhen, x)
    If nOldAll > 0 Then
        Set oldIdx = New Collection
        nOld = PickRows(oldAll, nOldAll, "", oldCur, oldIdx, x)
        newest = NewestStart(allRecs, nAll)
        savedNewest = NewestStart(oldAll, nOldAll)
        If newest < savedNewest Then
            Application.Cursor = xlDefault
            If Not Confirm("This file looks older than the one loaded last time." & vbLf & vbLf & _
                    "The newest start date in it is " & ShowDate8(newest) & ". In the saved rates it's " & _
                    ShowDate8(savedNewest) & "." & vbLf & vbLf & _
                    "Load it anyway? It will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
        If nNew < nOld * 0.9 Then
            Application.Cursor = xlDefault
            If Not Confirm("This file has only " & Format$(nNew, "#,##0") & " rates in force. The saved rates have " & _
                    Format$(nOld, "#,##0") & ", so the download may have been cut short." & vbLf & vbLf & _
                    "Load it anyway? It will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
    End If

    ' From here on only the rows still in force at the end of last month matter
    nKeep = KeepRows(allRecs, nAll, Format$(DateSerial(Year(today), Month(today), 0), "yyyymmdd"), keepRecs)

    Application.StatusBar = "Comparing..."
    Set thisMonth = New Collection
    Set nextMonth = New Collection
    MonthChanges keepRecs, nKeep, today, thisMonth, nextMonth, others

    ' Since the last file: what changed between the two files that the months don't show
    Set since = New Collection
    If nOldAll > 0 Then
        Set sinceAll = New Collection
        Compare oldCur, nOld, oldIdx, newCur, nNew, newIdx, allRecs, nAll, sinceAll, sinceOthers
        NotShown sinceAll, thisMonth, nextMonth, since
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Drawing..."
    DrawFront keepRecs, nKeep, thisMonth, nextMonth, others, fname, today, since.Count
    DrawSince since, (nOldAll > 0), fname, today, savedFile, savedWhen, dups
    AddToLog thisMonth, nextMonth, since, fname, today
    SaveCopy keepRecs, nKeep, fname, today, since.Count
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
    SheetByName(SH_MAIN).Range("A1").Select
End Sub

Private Sub RefreshFront()
    ' What changed from the saved rates, as of today
    Dim recs As Variant, n As Long, savedFile As String, savedWhen As Variant, sinceCount As Long
    Dim thisMonth As Collection, nextMonth As Collection, others As Long, loadedOn As Date
    n = ReadSaved(recs, savedFile, savedWhen, sinceCount)
    If n = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Set thisMonth = New Collection
    Set nextMonth = New Collection
    MonthChanges recs, n, CheckDate(), thisMonth, nextMonth, others
    If IsDate(savedWhen) Then loadedOn = CDate(savedWhen) Else loadedOn = CheckDate()
    DrawFront recs, n, thisMonth, nextMonth, others, savedFile, loadedOn, sinceCount
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
    SheetByName(SH_MAIN).Range("A1").Select
    Application.ScreenUpdating = True
End Sub


' ---- reading the state's file ---------------------------------------------------------

Private Function ReadRates(ByVal fullPath As String, ByRef recs As Variant, ByRef n As Long, _
                           ByRef hasHistory As Boolean, ByRef problem As String) As Boolean
    ' Reads the file into recs, one row per line, columns in HEADS order, dates as yyyymmdd.
    ' False, with problem saying why, when it isn't the rates file.
    Dim body As String, lines As Variant, i As Long, j As Long, c As Long, headAt As Long
    Dim fields As Variant, colOf(1 To NC) As Long, need As Variant, gotHead As Boolean

    n = 0: hasHistory = False
    On Error Resume Next
    body = ReadWholeFile(fullPath)
    If Err.Number <> 0 Then
        problem = "Couldn't open " & FileNameOf(fullPath) & ": " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    If Left$(body, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then body = Mid$(body, 4)
    If InStr(body, vbLf) = 0 Then body = Replace(body, vbCr, vbLf) Else body = Replace(body, vbCr, "")
    lines = Split(body, vbLf)

    ' The header is the first line carrying the columns this needs, within the first 20
    For i = 0 To UBound(lines)
        If i > 20 Then Exit For
        If Len(Trim$(lines(i))) > 0 Then
            Erase colOf
            fields = CsvFields(lines(i))
            For j = 0 To UBound(fields)
                c = HeadNumber(CStr(fields(j)))
                If c > 0 Then If colOf(c) = 0 Then colOf(c) = j + 1
            Next
            gotHead = True
            For Each need In Array(C_CODE, C_NAME, C_TAX, C_RTYPE, C_RATE)
                If colOf(need) = 0 Then gotHead = False
            Next
            If gotHead Then headAt = i: Exit For
        End If
    Next
    If Not gotHead Then
        problem = "That doesn't look like the Alabama rates file. " & FileNameOf(fullPath) & _
            " has no Locality Code, TaxType, Rate Type and Rate columns." & vbLf & vbLf & _
            "Download taxrates.csv from the Alabama DOR's Local Cities and Counties Tax Rates " & _
            "Text File page and pick that. Nothing was changed."
        Exit Function
    End If

    ReDim recs(1 To UBound(lines) - headAt + 1, 1 To NC)
    For i = headAt + 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            fields = CsvFields(lines(i))
            n = n + 1
            For c = 1 To NC
                If colOf(c) > 0 And colOf(c) - 1 <= UBound(fields) Then
                    recs(n, c) = Trim$(fields(colOf(c) - 1))
                Else
                    recs(n, c) = ""
                End If
            Next
            If recs(n, C_CODE) = "" Then
                n = n - 1
            Else
                recs(n, C_START) = Date8(recs(n, C_START))
                recs(n, C_END) = Date8(recs(n, C_END))
                If recs(n, C_END) <> "" Then hasHistory = True
            End If
        End If
    Next
    If n = 0 Then
        problem = "That file has no rates in it. Download taxrates.csv again and pick that." & _
            vbLf & vbLf & "Nothing was changed."
        Exit Function
    End If
    ReadRates = True
End Function

Private Function ReadWholeFile(ByVal fullPath As String) As String
    Dim f As Integer, body As String
    f = FreeFile
    Open fullPath For Binary Access Read Shared As #f
    If LOF(f) > 0 Then
        body = Space$(LOF(f))
        Get #f, , body
    End If
    Close #f
    ReadWholeFile = body
End Function

Private Function CsvFields(ByVal s As String) As Variant
    ' One line of the file split at its commas. Quotes around a field come off, and a comma
    ' inside quotes stays part of its field.
    Dim parts As Variant, i As Long, p As String, easy As Boolean
    Dim out() As String, n As Long, cur As String, inQuotes As Boolean, ch As String, j As Long

    parts = Split(s, ",")
    easy = True
    For i = 0 To UBound(parts)
        p = parts(i)
        If InStr(p, """") > 0 Then
            If Len(p) >= 2 And Left$(p, 1) = """" And Right$(p, 1) = """" Then
                p = Mid$(p, 2, Len(p) - 2)
                If InStr(p, """") > 0 Then easy = False: Exit For
                parts(i) = p
            Else
                easy = False: Exit For
            End If
        End If
    Next
    If easy Then CsvFields = parts: Exit Function

    ' The slow way, a letter at a time
    ReDim out(0 To 0)
    j = 1
    Do While j <= Len(s)
        ch = Mid$(s, j, 1)
        If inQuotes Then
            If ch = """" Then
                If Mid$(s, j + 1, 1) = """" Then
                    cur = cur & """"
                    j = j + 1
                Else
                    inQuotes = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQuotes = True
        ElseIf ch = "," Then
            out(n) = cur
            n = n + 1
            ReDim Preserve out(0 To n)
            cur = ""
        Else
            cur = cur & ch
        End If
        j = j + 1
    Loop
    out(n) = cur
    CsvFields = out
End Function

Private Function HeadNumber(ByVal head As String) As Long
    ' Which of the state's columns a header names, 0 for none
    Dim h As Variant, i As Long, want As String
    want = SqueezeHead(head)
    h = Split(HEADS, "|")
    For i = 0 To UBound(h)
        If SqueezeHead(CStr(h(i))) = want Then HeadNumber = i + 1: Exit Function
    Next
End Function

Private Function SqueezeHead(ByVal s As String) As String
    s = UCase$(s)
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, """", "")
    s = Replace(s, vbTab, "")
    SqueezeHead = s
End Function

Private Function Date8(ByVal s As String) As String
    ' A date as yyyymmdd, the way the state writes it. Also takes 9/1/2026 and 2026-09-01.
    Dim p As Variant
    s = Trim$(s)
    If s = "" Then Exit Function
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)     ' drop a time on the end
    If Len(s) = 8 And IsPlainNumber(s) Then Date8 = s: Exit Function
    If InStr(s, "/") > 0 Then
        p = Split(s, "/")
        If UBound(p) = 2 Then
            Date8 = Format$(Val(p(2)), "0000") & Format$(Val(p(0)), "00") & Format$(Val(p(1)), "00")
            Exit Function
        End If
    ElseIf InStr(s, "-") > 0 Then
        p = Split(s, "-")
        If UBound(p) = 2 Then
            Date8 = Format$(Val(p(0)), "0000") & Format$(Val(p(1)), "00") & Format$(Val(p(2)), "00")
            Exit Function
        End If
    End If
    Date8 = s
End Function


' ---- which rows count ------------------------------------------------------------------

Private Function PickRows(ByRef recs As Variant, ByVal n As Long, ByVal asOf As String, _
                          ByRef picked As Variant, ByVal idx As Object, ByRef dups As Long) As Long
    ' The rates in force. With asOf blank, that's every row with no end date. With asOf a
    ' yyyymmdd day, it's every row that had started by then and hadn't ended yet. When two
    ' rows cover the same rate, the one that started later counts. idx gets each rate's key.
    Dim i As Long, c As Long, k As String, at As Long, cnt As Long, st As String, en As String
    dups = 0
    If n = 0 Then Exit Function
    ReDim picked(1 To n, 1 To NC)
    For i = 1 To n
        st = recs(i, C_START)
        en = recs(i, C_END)
        If asOf = "" Then
            If en <> "" Then GoTo NextRow
        Else
            If st > asOf Then GoTo NextRow
            If en <> "" Then If en < asOf Then GoTo NextRow
        End If
        k = RowKey(recs, i)
        at = Found(idx, k)
        If at = 0 Then
            cnt = cnt + 1
            at = cnt
            idx.Add at, k
        Else
            dups = dups + 1
            If st < picked(at, C_START) Then GoTo NextRow
        End If
        For c = 1 To NC
            picked(at, c) = recs(i, c)
        Next
NextRow:
    Next
    PickRows = cnt
End Function

Private Function KeepRows(ByRef recs As Variant, ByVal n As Long, ByVal fromDay As String, _
                          ByRef kept As Variant) As Long
    ' The rows still in force on fromDay or later. Older history isn't needed.
    Dim i As Long, c As Long, cnt As Long
    If n = 0 Then Exit Function
    ReDim kept(1 To n, 1 To NC)
    For i = 1 To n
        If recs(i, C_END) = "" Or recs(i, C_END) >= fromDay Then
            cnt = cnt + 1
            For c = 1 To NC
                kept(cnt, c) = recs(i, c)
            Next
        End If
    Next
    KeepRows = cnt
End Function

Private Function RowKey(ByRef recs As Variant, ByVal i As Long) As String
    ' A rate is one locality, tax type, rate type and county code. A city in two counties has
    ' a row for each county.
    RowKey = UCase$(recs(i, C_CODE) & "|" & recs(i, C_TAX) & "|" & recs(i, C_RTYPE) & "|" & recs(i, C_CCODE))
End Function

Private Function Found(ByVal idx As Object, ByVal k As String) As Long
    On Error Resume Next
    Found = idx(k)
    If Err.Number <> 0 Then Found = 0
    On Error GoTo 0
End Function

Private Function NewestStart(ByRef recs As Variant, ByVal n As Long) As String
    Dim i As Long, best As String
    For i = 1 To n
        If recs(i, C_START) > best Then best = recs(i, C_START)
    Next
    NewestStart = best
End Function


' ---- comparing -------------------------------------------------------------------------

Private Sub MonthChanges(ByRef recs As Variant, ByVal n As Long, ByVal onDay As Date, _
                         ByVal thisMonth As Object, ByVal nextMonth As Object, ByRef others As Long)
    ' thisMonth: the rates in force on last month's last day against this month's last day,
    ' so everything that started this month. nextMonth: the same, a month on.
    Dim d0 As String, d1 As String, d2 As String, x As Long, o1 As Long, o2 As Long
    Dim s0 As Variant, s1 As Variant, s2 As Variant, n0 As Long, n1 As Long, n2 As Long
    Dim i0 As Collection, i1 As Collection, i2 As Collection
    d0 = Format$(DateSerial(Year(onDay), Month(onDay), 0), "yyyymmdd")
    d1 = Format$(DateSerial(Year(onDay), Month(onDay) + 1, 0), "yyyymmdd")
    d2 = Format$(DateSerial(Year(onDay), Month(onDay) + 2, 0), "yyyymmdd")
    Set i0 = New Collection
    n0 = PickRows(recs, n, d0, s0, i0, x)
    Set i1 = New Collection
    n1 = PickRows(recs, n, d1, s1, i1, x)
    Set i2 = New Collection
    n2 = PickRows(recs, n, d2, s2, i2, x)
    Compare s0, n0, i0, s1, n1, i1, recs, n, thisMonth, o1
    Compare s1, n1, i1, s2, n2, i2, recs, n, nextMonth, o2
    others = o1 + o2
End Sub

Private Sub Compare(ByRef oldRecs As Variant, ByVal nOld As Long, ByVal oldIdx As Object, _
                    ByRef newRecs As Variant, ByVal nNew As Long, ByVal newIdx As Object, _
                    ByRef allRecs As Variant, ByVal nAll As Long, ByVal changes As Object, _
                    ByRef others As Long)
    ' Every watched rate that's new, gone or different between two sets of rates. others
    ' counts the same for the rates not watched.
    Dim i As Long, j As Long, what As String, look As String, notes As String, en As String
    Dim goneAt As Collection, ended As Variant, v As Variant

    others = 0
    For i = 1 To nNew
        j = Found(oldIdx, RowKey(newRecs, i))
        If j = 0 Then
            If Watched(newRecs(i, C_TAX), newRecs(i, C_RTYPE)) Then
                notes = ""
                If newRecs(i, C_ADMIN) <> "" Then notes = "Administered by " & newRecs(i, C_ADMIN) & "."
                changes.Add MakeChange("New", "new", notes, oldRecs, 0, newRecs, i)
            Else
                others = others + 1
            End If
        ElseIf Differs(oldRecs, j, newRecs, i, what, look, notes) Then
            If Watched(newRecs(i, C_TAX), newRecs(i, C_RTYPE)) Then
                changes.Add MakeChange(what, look, notes, oldRecs, j, newRecs, i)
            Else
                others = others + 1
            End If
        End If
    Next

    ' Gone: in the old rates but not the new
    Set goneAt = New Collection
    For j = 1 To nOld
        If Found(newIdx, RowKey(oldRecs, j)) = 0 Then
            If Watched(oldRecs(j, C_TAX), oldRecs(j, C_RTYPE)) Then
                goneAt.Add j, RowKey(oldRecs, j)
            Else
                others = others + 1
            End If
        End If
    Next
    If goneAt.Count = 0 Then Exit Sub

    ' When it ended: the old row's own end date, or else the latest one the history has
    ReDim ended(1 To nOld)
    For i = 1 To nAll
        If allRecs(i, C_END) <> "" Then
            j = Found(goneAt, RowKey(allRecs, i))
            If j > 0 Then If allRecs(i, C_END) > ended(j) Then ended(j) = allRecs(i, C_END)
        End If
    Next
    For Each v In goneAt
        j = v
        en = oldRecs(j, C_END)
        If en = "" Then en = ended(j)
        notes = ""
        If en <> "" Then notes = "Ended " & ShowDate8(en) & "."
        changes.Add MakeChange("Gone", "gone", notes, oldRecs, j, newRecs, 0)
    Next
End Sub

Private Function Differs(ByRef oldRecs As Variant, ByVal j As Long, ByRef newRecs As Variant, _
                         ByVal i As Long, ByRef what As String, ByRef look As String, _
                         ByRef notes As String) As Boolean
    ' True when anything changed besides the dates and the indicator. what is the headline,
    ' notes has the rest.
    Dim oldRate As Double, newRate As Double, oldPJ As Double, newPJ As Double, bits As String
    what = "": look = "": notes = ""
    oldRate = Num(oldRecs(j, C_RATE)): newRate = Num(newRecs(i, C_RATE))
    oldPJ = Num(oldRecs(j, C_PJRATE)): newPJ = Num(newRecs(i, C_PJRATE))
    If Abs(newRate - oldRate) > 0.000001 Then
        If newRate > oldRate Then what = "Went up": look = "up" Else what = "Went down": look = "down"
    ElseIf Abs(newPJ - oldPJ) > 0.000001 Then
        If newPJ > oldPJ Then what = "PJ rate went up": look = "up" Else what = "PJ rate went down": look = "down"
    End If
    If Not SameText(oldRecs(j, C_PJ), newRecs(i, C_PJ)) Then
        bits = bits & " PJ was " & OrBlank(oldRecs(j, C_PJ)) & ", now " & OrBlank(newRecs(i, C_PJ)) & "."
        If what = "" Then what = "PJ changed": look = "other"
    End If
    If Not SameText(oldRecs(j, C_ADMIN), newRecs(i, C_ADMIN)) Then
        bits = bits & " Administrator was " & OrBlank(oldRecs(j, C_ADMIN)) & ", now " & _
            OrBlank(newRecs(i, C_ADMIN)) & "."
        If what = "" Then what = "New administrator": look = "other"
    End If
    If Not SameText(oldRecs(j, C_NAME), newRecs(i, C_NAME)) Then
        bits = bits & " Was named " & oldRecs(j, C_NAME) & "."
        If what = "" Then what = "Name changed": look = "other"
    End If
    If Not SameNumber(oldRecs(j, C_CNUM), newRecs(i, C_CNUM)) Then
        bits = bits & " County number was " & OrBlank(oldRecs(j, C_CNUM)) & ", now " & _
            OrBlank(newRecs(i, C_CNUM)) & "."
        If what = "" Then what = "County number changed": look = "other"
    End If
    notes = Trim$(bits)
    Differs = (what <> "")
End Function

Private Function MakeChange(ByVal what As String, ByVal look As String, ByVal notes As String, _
                            ByRef oldRecs As Variant, ByVal j As Long, _
                            ByRef newRecs As Variant, ByVal i As Long) As Variant
    ' One line of a list. i is its row in the new rates and j in the old, 0 for none.
    Dim f(0 To 17) As Variant, nm As String, tax As String, rtype As String, k As String
    Dim code As String, ccode As String, st As String, newRate As String, newPJ As String
    If i > 0 Then
        nm = newRecs(i, C_NAME): tax = newRecs(i, C_TAX): rtype = newRecs(i, C_RTYPE)
        code = newRecs(i, C_CODE): ccode = newRecs(i, C_CCODE): st = newRecs(i, C_START)
        newRate = newRecs(i, C_RATE): newPJ = newRecs(i, C_PJRATE)
        k = RowKey(newRecs, i)
    Else
        nm = oldRecs(j, C_NAME): tax = oldRecs(j, C_TAX): rtype = oldRecs(j, C_RTYPE)
        code = oldRecs(j, C_CODE): ccode = oldRecs(j, C_CCODE)
        k = RowKey(oldRecs, j)
    End If
    f(F_NAME) = nm
    f(F_TAX) = TaxWords(tax, rtype)
    f(F_WHAT) = what
    If j > 0 Then
        f(F_OLD) = Num(oldRecs(j, C_RATE))
        If oldRecs(j, C_PJRATE) <> "" Then f(F_OLDPJ) = Num(oldRecs(j, C_PJRATE))
    End If
    If i > 0 Then
        f(F_NEW) = Num(newRate)
        If newPJ <> "" Then f(F_NEWPJ) = Num(newPJ)
    End If
    If st <> "" Then f(F_START) = DateFrom8(st)
    f(F_CODE) = NumberOrText(code)
    If ccode <> "" Then f(F_CCODE) = NumberOrText(ccode)
    f(F_NOTES) = notes
    f(F_RTYPE) = UCase$(rtype)
    f(F_LOOK) = look
    f(F_KEY) = k
    If j > 0 Then f(F_WAS) = Details(oldRecs, j) Else f(F_WAS) = ""
    If i > 0 Then f(F_NOW) = Details(newRecs, i) Else f(F_NOW) = ""
    ' The same change seen again, by a later file or a later month, has the same ID
    f(F_ID) = k & "|" & what & "|" & st & "|" & NumText(newRate) & "|" & NumText(newPJ) & "|" & UCase$(notes)
    f(F_SORT) = UCase$(nm) & Chr$(1) & UCase$(ccode) & Chr$(1) & TaxOrder(tax) & Chr$(1) & _
        UCase$(rtype) & Chr$(1) & UCase$(code)
    MakeChange = f
End Function

Private Sub NotShown(ByVal all As Object, ByVal thisMonth As Object, ByVal nextMonth As Object, _
                     ByVal out As Object)
    ' The changes in all that the month lists don't already show: the same rate, the same
    ' before and the same after. A rate that changed this month and was also quietly renamed
    ' still counts, because its before isn't the same.
    Dim shown As Collection, ch As Variant
    Set shown = New Collection
    On Error Resume Next
    For Each ch In thisMonth
        shown.Add 1, CStr(ch(F_KEY) & "#" & ch(F_WAS) & "#" & ch(F_NOW))
    Next
    For Each ch In nextMonth
        shown.Add 1, CStr(ch(F_KEY) & "#" & ch(F_WAS) & "#" & ch(F_NOW))
    Next
    On Error GoTo 0
    For Each ch In all
        If Found(shown, CStr(ch(F_KEY) & "#" & ch(F_WAS) & "#" & ch(F_NOW))) = 0 Then out.Add ch
    Next
End Sub

Private Function Details(ByRef recs As Variant, ByVal i As Long) As String
    ' Everything about a rate that counts as a change, in one piece of text
    Dim cnum As String
    cnum = Trim$(recs(i, C_CNUM))
    If IsPlainNumber(cnum) Then cnum = Trim$(Str$(Val(cnum)))
    Details = UCase$(Trim$(recs(i, C_NAME))) & "|" & NumText(recs(i, C_RATE)) & "|" & _
        NumText(recs(i, C_PJRATE)) & "|" & UCase$(Trim$(recs(i, C_PJ))) & "|" & _
        UCase$(Trim$(recs(i, C_ADMIN))) & "|" & cnum
End Function

Private Function SortedOrder(ByVal changes As Object) As Variant
    ' The changes by locality name, then county code, then tax: a place's changes sit together.
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


' ---- the sheets ------------------------------------------------------------------------

Private Sub EnsureSheets()
    ' First run: makes the four sheets. Later runs only put back the button.
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
    If SheetByName(SH_SINCE) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_MAIN))
        ws.Name = SH_SINCE
        ws.Range("A1").Value = "Nothing loaded yet."
        ws.Range("A1").Font.Size = 14
    End If
    If SheetByName(SH_LOG) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_SINCE))
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
        .Value = "Click the button and pick the rates file."
        .Font.Size = 20
    End With
    ws.Rows(3).RowHeight = 34
    ws.Range("A4").Value = "Get it from the Alabama DOR: revenue.alabama.gov, Sales & Use, " & _
        "Local Cities and Counties Tax Rates Text File, then taxrates.csv."
    ws.Rows(4).RowHeight = 24
    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
    DrawButton ws
End Sub

Private Sub DrawFront(ByRef recs As Variant, ByVal n As Long, ByVal thisMonth As Object, _
                      ByVal nextMonth As Object, ByVal others As Long, ByVal fname As String, _
                      ByVal loadedOn As Date, ByVal sinceCount As Long)
    ' What changed: which file, then this month's changes and next month's, in big type.
    Dim ws As Worksheet, r As Long, today As Date, m0 As Date, m1 As Date, m2 As Date
    Dim newest As String, s As String

    today = CheckDate()
    m0 = DateSerial(Year(today), Month(today), 1)
    m1 = DateSerial(Year(today), Month(today) + 1, 1)
    m2 = DateSerial(Year(today), Month(today) + 2, 1)
    newest = NewestStart(recs, n)

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
    ws.Range("A2").Value = "From " & fname & ", loaded " & Format$(loadedOn, "m/d/yyyy") & _
        ". The newest start date in it is " & ShowDate8(newest) & "."
    s = "Showing only " & JoinWords(WATCH_TAX, True) & ", " & JoinWords(WATCH_RATE, False) & "."
    If others = 1 Then s = s & " 1 other change in these two months isn't shown."
    If others > 1 Then s = s & " " & Format$(others, "#,##0") & " other changes in these two months aren't shown."
    ws.Range("A3").Value = s
    ws.Range("A3").Font.Color = RGB(89, 89, 89)
    ws.Rows(2).RowHeight = 24
    ws.Rows(3).RowHeight = 24
    r = 4

    ' A file loaded in an earlier month, with nothing starting this month, is out of date
    If loadedOn < m0 And newest < Format$(m0, "yyyymmdd") Then
        Band ws, r, "GET THE NEW FILE. THIS ONE WAS LOADED " & Format$(loadedOn, "m/d/yyyy") & ".", _
            RGB(255, 199, 206), RGB(156, 0, 6), 24
        r = r + 1
    End If
    If sinceCount > 0 Then
        If sinceCount = 1 Then
            s = "Also 1 change since the last file that isn't below. See the Since last file tab."
        Else
            s = "Also " & sinceCount & " changes since the last file that aren't below. See the Since last file tab."
        End If
        Band ws, r, s, LookFill("other"), LookText("other"), 16
        r = r + 1
    End If
    ws.Rows(r).RowHeight = 12
    r = r + 1

    MonthSection ws, r, m0, thisMonth, (newest >= Format$(m0, "yyyymmdd")), True
    ws.Rows(r).RowHeight = 16
    r = r + 1
    MonthSection ws, r, m1, nextMonth, (newest >= Format$(m1, "yyyymmdd")), False

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

Private Sub MonthSection(ByVal ws As Worksheet, ByRef r As Long, ByVal monthStart As Date, _
                         ByVal changes As Object, ByVal covered As Boolean, ByVal isThis As Boolean)
    ' One month: a big line with the count, a line saying what it covers, then the list.
    Dim n As Long, heading As String, countText As String, fillColor As Long, textColor As Long
    Dim sub1 As String, mname As String
    n = changes.Count
    mname = Format$(monthStart, "mmmm yyyy")
    heading = UCase$(Format$(monthStart - 1, "mmmm") & " to " & Format$(monthStart, "mmmm"))
    If n > 0 Then
        If n = 1 Then countText = "1 CHANGE" Else countText = Format$(n, "#,##0") & " CHANGES"
        fillColor = RGB(255, 217, 102): textColor = RGB(0, 0, 0)
        If isThis Then sub1 = "Took effect in " & mname & "." Else sub1 = "Take effect in " & mname & "."
    ElseIf covered Then
        countText = "NO CHANGES"
        fillColor = RGB(198, 239, 206): textColor = RGB(0, 97, 0)
        sub1 = "Nothing you watch changes in " & mname & "."
    Else
        countText = "NONE POSTED YET"
        fillColor = RGB(221, 235, 247): textColor = RGB(31, 78, 121)
        sub1 = "Nothing in this file starts in " & mname & " yet. Check again once the state posts " & _
            Format$(monthStart, "mmmm") & "'s file."
    End If
    Band ws, r, heading & ": " & countText, fillColor, textColor, 28
    r = r + 1
    ws.Cells(r, 1).Value = sub1
    ws.Rows(r).RowHeight = 24
    r = r + 1
    If n > 0 Then DrawList ws, r, changes
End Sub

Private Sub Band(ByVal ws As Worksheet, ByVal r As Long, ByVal words As String, _
                 ByVal fillColor As Long, ByVal textColor As Long, ByVal pts As Long)
    ' A colored line across the list, for the things that need to be seen
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
        .Merge
        .Value = words
        .Font.Size = pts
        .Font.Bold = True
        .Font.Color = textColor
        .Interior.Color = fillColor
        .HorizontalAlignment = xlLeft
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
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 11))
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .WrapText = True
    End With
    ws.Rows(r).RowHeight = 44
    order = SortedOrder(changes)
    ReDim out(1 To n, 1 To 11)
    For i = 1 To n
        ch = changes(order(i))
        For c = 1 To 11
            out(i, c) = ch(c)
        Next
    Next
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 11))
        .Value = out
        .Font.Size = 16
    End With
    For i = 1 To n
        ch = changes(order(i))
        rr = r + i
        ws.Range(ws.Cells(rr, 4), ws.Cells(rr, 5)).NumberFormat = RateFormat(CStr(ch(F_RTYPE)))
        ws.Range(ws.Cells(rr, 4), ws.Cells(rr, 5)).Font.Bold = True
        ws.Range(ws.Cells(rr, 7), ws.Cells(rr, 8)).NumberFormat = RateFormat(CStr(ch(F_RTYPE)))
        ws.Cells(rr, 6).NumberFormat = "m/d/yyyy"
        With ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 11)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
        With ws.Cells(rr, 3)
            .Interior.Color = LookFill(CStr(ch(F_LOOK)))
            .Font.Color = LookText(CStr(ch(F_LOOK)))
            .Font.Bold = True
        End With
    Next
    ws.Range(ws.Cells(r + 1, 9), ws.Cells(r + n, 10)).NumberFormat = "0"
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).WrapText = True
    ws.Range(ws.Cells(r + 1, 11), ws.Cells(r + n, 11)).WrapText = True
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).EntireRow.AutoFit
    For rr = r + 1 To r + n
        If ws.Rows(rr).RowHeight < 30 Then ws.Rows(rr).RowHeight = 30
    Next
    r = r + n + 1
End Sub

Private Sub DrawSince(ByVal since As Object, ByVal compared As Boolean, ByVal fname As String, _
                      ByVal loadedOn As Date, ByVal savedFile As String, ByVal savedWhen As Variant, _
                      ByVal dups As Long)
    ' Since last file: the safety net, what the two files differ on that the months miss
    Dim ws As Worksheet, n As Long, r As Long, s As String
    Set ws = SheetByName(SH_SINCE)
    n = since.Count
    ClearSheet ws
    ListWidths ws
    ws.Cells.Font.Size = 14
    ws.Cells.VerticalAlignment = xlCenter
    With ws.Range("A1")
        .Value = SINCE_TITLE
        .Font.Size = 26
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 50
    If compared Then
        s = "From " & fname & ", loaded " & Format$(loadedOn, "m/d/yyyy") & ", against " & savedFile & _
            ", loaded " & ShowWhen(savedWhen) & "."
    Else
        s = "First file, so nothing to compare yet. The next file gets checked against this one."
    End If
    ws.Range("A2").Value = s
    s = "What changed between the two files that the What changed tab doesn't show: a row the " & _
        "state added or fixed with an old start date, or a change in a month nobody checked."
    If dups = 1 Then s = s & " The file repeats 1 row; the newer one was used."
    If dups > 1 Then s = s & " The file repeats " & dups & " rows; the newest of each was used."
    ws.Range("A3").Value = s
    ws.Range("A3").Font.Color = RGB(89, 89, 89)
    ws.Rows(2).RowHeight = 24
    ws.Rows(3).RowHeight = 24
    ws.Rows(4).RowHeight = 12
    r = 5
    If n > 0 Then
        If n = 1 Then s = "1 CHANGE" Else s = Format$(n, "#,##0") & " CHANGES"
        Band ws, r, s, RGB(255, 217, 102), RGB(0, 0, 0), 28
        r = r + 2
        DrawList ws, r, since
    ElseIf compared Then
        Band ws, r, "NOTHING HERE", RGB(198, 239, 206), RGB(0, 97, 0), 28
    Else
        Band ws, r, "NOTHING TO COMPARE YET", RGB(221, 235, 247), RGB(31, 78, 121), 28
    End If
    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub ListWidths(ByVal ws As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(50, 18, 30, 12, 12, 14, 15, 15, 9, 11, 50)
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next
End Sub

Private Function JoinWords(ByVal list As String, ByVal taxes As Boolean) As String
    ' "sales and sellers use", or "general rate"
    Dim p As Variant, i As Long, w As String, s As String
    p = Split(Replace(UCase$(list), " ", ""), ",")
    For i = 0 To UBound(p)
        If taxes Then w = LCase$(TaxWords(CStr(p(i)), "GENER")) Else w = RateWords(CStr(p(i)))
        If i = 0 Then
            s = w
        ElseIf i = UBound(p) Then
            s = s & " and " & w
        Else
            s = s & ", " & w
        End If
    Next
    If Not taxes Then
        If UBound(p) > 0 Then s = s & " rates" Else s = s & " rate"
    End If
    JoinWords = s
End Function

Private Sub AddToLog(ByVal thisMonth As Object, ByVal nextMonth As Object, ByVal since As Object, _
                     ByVal fname As String, ByVal loadedOn As Date)
    ' Each change goes on the Log once, the first time a file shows it. A check with nothing
    ' new still gets a line.
    Dim ws As Worksheet, r As Long, last As Long, ids As Collection, v As Variant, i As Long
    Dim out As Variant, n As Long, lists As Variant, labels As Variant, s As Long
    Dim order As Variant, ch As Variant, c As Long, m0 As Date, lst As Object, looks() As String

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

    m0 = DateSerial(Year(loadedOn), Month(loadedOn), 1)
    lists = Array(thisMonth, nextMonth, since)
    labels = Array(Format$(m0 - 1, "mmmm") & " to " & Format$(m0, "mmmm"), _
                   Format$(m0, "mmmm") & " to " & Format$(DateSerial(Year(m0), Month(m0) + 1, 1), "mmmm"), _
                   SH_SINCE)
    ReDim out(1 To thisMonth.Count + nextMonth.Count + since.Count + 1, 1 To LOG_COLS)
    ReDim looks(0 To 0)
    For s = 0 To 2
        Set lst = lists(s)
        If lst.Count > 0 Then
            order = SortedOrder(lst)
            For i = 1 To lst.Count
                ch = lst(order(i))
                If Found(ids, CStr(ch(F_ID))) = 0 Then
                    ids.Add 1, CStr(ch(F_ID))
                    n = n + 1
                    out(n, 1) = loadedOn
                    out(n, 2) = fname
                    out(n, 3) = labels(s)
                    For c = 1 To 11
                        out(n, c + 3) = ch(c)
                    Next
                    out(n, LOG_COLS) = ch(F_ID)
                    ReDim Preserve looks(0 To n)
                    looks(n) = CStr(ch(F_LOOK)) & "|" & CStr(ch(F_RTYPE))
                End If
            Next
        End If
    Next
    If n = 0 Then
        n = 1
        out(1, 1) = loadedOn
        out(1, 2) = fname
        out(1, 6) = "Nothing new"
        out(1, 14) = "Everything this file shows was already on the Log."
        ReDim Preserve looks(0 To 1)
        looks(1) = ""
    End If

    r = last + 1
    If r < 2 Then r = 2
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, LOG_COLS))
        .Value = out
        .Font.Size = 14
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 1)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 9), ws.Cells(r + n - 1, 9)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 12), ws.Cells(r + n - 1, 13)).NumberFormat = "0"
    ws.Range(ws.Cells(r, 14), ws.Cells(r + n - 1, 14)).WrapText = True
    For i = 1 To n
        If looks(i) <> "" Then
            v = Split(looks(i), "|")
            ws.Range(ws.Cells(r + i - 1, 7), ws.Cells(r + i - 1, 8)).NumberFormat = RateFormat(CStr(v(1)))
            ws.Range(ws.Cells(r + i - 1, 10), ws.Cells(r + i - 1, 11)).NumberFormat = RateFormat(CStr(v(1)))
            With ws.Cells(r + i - 1, 6)
                .Interior.Color = LookFill(CStr(v(0)))
                .Font.Color = LookText(CStr(v(0)))
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
    widths = Array(13, 16, 24, 44, 18, 24, 11, 11, 13, 12, 12, 8, 10, 50, 10)
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
                     ByVal loadedOn As Date, ByVal sinceCount As Long)
    ' Saved rates: the file's rates from the end of last month on, for opening later and for
    ' the next file to be checked against
    Dim ws As Worksheet, out As Variant, i As Long, c As Long, titles As Variant, col As Variant
    Dim inForce As Long
    Set ws = SheetByName(SH_SAVED)
    ws.Unprotect
    ws.Cells.Clear
    ws.Cells.Font.Size = 11
    With ws.Range("A1")
        .Value = "The rates from the last file loaded. What changed is drawn from these, and the " & _
            "next file gets checked against them. Loading a file replaces them, so leave this sheet alone."
        .Font.Size = 14
        .Font.Bold = True
    End With
    For i = 1 To n
        If recs(i, C_END) = "" Then inForce = inForce + 1
    Next
    ws.Range("A2").Value = "File"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = fname
    ws.Range("C2").Value = "Loaded"
    ws.Range("D2").Value = loadedOn
    ws.Range("D2").NumberFormat = "m/d/yyyy"
    ws.Range("E2").Value = "Rates in force"
    ws.Range("F2").Value = inForce
    ws.Range("G2").Value = "Since last file"
    ws.Range("H2").Value = sinceCount
    ws.Range("A2,C2,E2,G2").Font.Bold = True
    titles = Split(HEADS, "|")
    For c = 1 To NC
        ws.Cells(SAVED_ROW, c).Value = titles(c - 1)
    Next
    ws.Range(ws.Cells(SAVED_ROW, 1), ws.Cells(SAVED_ROW, NC)).Font.Bold = True
    ReDim out(1 To n, 1 To NC)
    For i = 1 To n
        For c = 1 To NC
            out(i, c) = CellValue(CStr(recs(i, c)), c)
        Next
    Next
    ' Words stay words, whatever they look like
    For Each col In Array(C_NAME, C_TAX, C_RTYPE, C_ADMIN, C_IND, C_PJ)
        ws.Range(ws.Cells(SAVED_ROW + 1, col), ws.Cells(SAVED_ROW + n, col)).NumberFormat = "@"
    Next
    ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + n, NC)).Value = out
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 34
    ws.Range(ws.Columns(3), ws.Columns(NC)).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 14
    ws.Protect
End Sub

Private Function ReadSaved(ByRef recs As Variant, ByRef savedFile As String, ByRef savedWhen As Variant, _
                           ByRef sinceCount As Long) As Long
    ' Every row on Saved rates, as text the way the file had it. 0 when nothing's saved yet.
    Dim ws As Worksheet, last As Long, n As Long, i As Long, c As Long
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Cells(SAVED_ROW, 1).Value) <> "Locality Code" Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last <= SAVED_ROW Then Exit Function
    n = last - SAVED_ROW
    recs = ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(last, NC)).Value
    For i = 1 To n
        For c = 1 To NC
            recs(i, c) = AsText(recs(i, c))
        Next
    Next
    savedFile = CStr(ws.Range("B2").Value)
    savedWhen = ws.Range("D2").Value
    sinceCount = Val(CStr(ws.Range("H2").Value))
    ReadSaved = n
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
    b.OnAction = "CheckNewRatesFile"
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

Private Function IsBlankSheet(ByVal ws As Worksheet) As Boolean
    IsBlankSheet = (Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0)
End Function

Private Function Watched(ByVal tax As String, ByVal rtype As String) As Boolean
    Watched = InList(tax, WATCH_TAX) And InList(rtype, WATCH_RATE)
End Function

Private Function InList(ByVal s As String, ByVal list As String) As Boolean
    InList = InStr(1, "," & Replace(UCase$(list), " ", "") & ",", "," & UCase$(Trim$(s)) & ",", vbBinaryCompare) > 0
End Function

Private Function TaxOrder(ByVal tax As String) As String
    ' Sales before sellers use, the order they're listed in WATCH_TAX
    Dim p As Variant, i As Long
    p = Split(Replace(UCase$(WATCH_TAX), " ", ""), ",")
    For i = 0 To UBound(p)
        If p(i) = UCase$(Trim$(tax)) Then TaxOrder = Format$(i + 1, "00"): Exit Function
    Next
    TaxOrder = "99"
End Function

Private Function TaxWords(ByVal tax As String, ByVal rtype As String) As String
    Dim t As String
    Select Case UCase$(Trim$(tax))
        Case "ST": t = "Sales"
        Case "SU": t = "Sellers use"
        Case "CU": t = "Consumers use"
        Case "LL": t = "Lodgings"
        Case "RT": t = "Rental"
        Case Else: t = tax
    End Select
    If UCase$(Trim$(rtype)) <> "GENER" Then t = t & ", " & RateWords(rtype)
    TaxWords = t
End Function

Private Function RateWords(ByVal rtype As String) As String
    Select Case UCase$(Trim$(rtype))
        Case "GENER": RateWords = "general"
        Case "AUTO": RateWords = "automotive"
        Case "FARM": RateWords = "farm"
        Case "MACH": RateWords = "machine"
        Case "GROC": RateWords = "food"
        Case "VEND": RateWords = "vending"
        Case "AMUSE": RateWords = "amusement"
        Case "WDFEE": RateWords = "withdrawal fee"
        Case "LINEN": RateWords = "linen"
        Case "RMFEE": RateWords = "room fee"
        Case "ALCOH": RateWords = "alcohol"
        Case "REST": RateWords = "food service"
        Case "TRANV": RateWords = "vehicle rental"
        Case "MINE": RateWords = "mining"
        Case Else: RateWords = LCase$(rtype)
    End Select
End Function

Private Function RateFormat(ByVal rtype As String) As String
    ' Rates in the file are percents already: 2.5 means 2.5%. The withdrawal fee is dollars.
    Select Case UCase$(rtype)
        Case "WDFEE": RateFormat = "$#,##0.00"
        Case "RMFEE": RateFormat = "0.00##"
        Case Else: RateFormat = "0.00##""%"""
    End Select
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

Private Function SameText(ByVal a As String, ByVal b As String) As Boolean
    SameText = (UCase$(Trim$(a)) = UCase$(Trim$(b)))
End Function

Private Function SameNumber(ByVal a As String, ByVal b As String) As Boolean
    If IsPlainNumber(Trim$(a)) And IsPlainNumber(Trim$(b)) Then
        SameNumber = (Val(a) = Val(b))
    Else
        SameNumber = SameText(a, b)
    End If
End Function

Private Function OrBlank(ByVal s As String) As String
    If Trim$(s) = "" Then OrBlank = "blank" Else OrBlank = Trim$(s)
End Function

Private Function Num(ByVal s As String) As Double
    Num = Val(Replace(Replace(Replace(Trim$(s), "$", ""), "%", ""), ",", ""))
End Function

Private Function NumText(ByVal s As String) As String
    ' A rate the same way however the file wrote it: 2.5000 and 2.5 both come out 2.5
    If Trim$(s) = "" Then NumText = "" Else NumText = Trim$(Str$(Num(s)))
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

Private Function NumberOrText(ByVal s As String) As Variant
    ' Codes as numbers, unless a leading zero would get lost
    If IsPlainNumber(s) And Not (Len(s) > 1 And Left$(s, 1) = "0") Then
        NumberOrText = Val(s)
    Else
        NumberOrText = s
    End If
End Function

Private Function CellValue(ByVal s As String, ByVal c As Long) As Variant
    ' How a value from the file goes onto Saved rates
    Select Case c
        Case C_NAME, C_TAX, C_RTYPE, C_ADMIN, C_IND, C_PJ
            If s = "" Then CellValue = Empty Else CellValue = s
        Case C_CODE, C_CNUM, C_CCODE
            If IsPlainNumber(s) And Not (Len(s) > 1 And Left$(s, 1) = "0") Then
                CellValue = Val(s)
            ElseIf s = "" Then
                CellValue = Empty
            Else
                CellValue = "'" & s
            End If
        Case Else
            If IsPlainNumber(s) Then
                CellValue = Val(s)
            ElseIf s = "" Then
                CellValue = Empty
            Else
                CellValue = "'" & s
            End If
    End Select
End Function

Private Function AsText(ByVal v As Variant) As String
    ' A value read back off Saved rates, as text the way the file had it
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        AsText = Trim$(v)
    ElseIf IsNumeric(v) Then
        If v = Int(v) Then AsText = Format$(v, "0") Else AsText = Trim$(Str$(v))
    Else
        AsText = Trim$(CStr(v))
    End If
End Function

Private Function DateFrom8(ByVal s As String) As Date
    DateFrom8 = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 5, 2)), Val(Mid$(s, 7, 2)))
End Function

Private Function ShowDate8(ByVal s As String) As String
    If Len(s) = 8 Then ShowDate8 = Format$(DateFrom8(s), "m/d/yyyy") Else ShowDate8 = s
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
    PickFile = Application.GetOpenFilename("Rates file (*.csv),*.csv", 1, _
        "Pick the new Alabama rates file (taxrates.csv)")
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

Public Function RatesTest(ByVal answers As String, ByVal todayText As String) As String
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
        CheckNewRatesFile
    End If
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    RatesTest = mSaid
    Set mFeed = Nothing
    mToday = 0
End Function
