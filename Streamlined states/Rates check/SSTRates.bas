' Streamlined rates check: what changed in one state's sales tax rates since the last file.
'
' One macro for every Streamlined state. You keep one workbook per state, each with this
' same macro: Ohio, Washington, Nebraska and Oklahoma. After its first file, a workbook
' takes only that state's files.
'
' Setting up all four, once:
'   1. Open a new blank workbook. Alt F11, then Insert > Module. If the new window already
'      has a line in it, Ctrl A and Delete first. Paste all of this in, then close that window.
'   2. File > Save As > Excel Macro-Enabled Workbook (.xlsm). Name it SST Rates Check OH.
'   3. Before running anything, Save As three more times the same way: SST Rates Check WA,
'      SST Rates Check NE, SST Rates Check OK. Each copy has the macro and no state yet.
'   4. In each one: Alt F8, pick CheckNewRatesFile, Run. It makes the tabs and the big
'      button, then asks for that state's file. If you don't have the file yet, click Cancel.
'
' After that, the Read here tab in each workbook has everything: where to get the file,
' the steps each quarter, what the colors and lines mean, and how to set up another state.
'
' What it watches: each place's general rate, in-state and out-of-state. Food and drug
' rates aren't watched. A row the state only re-dated isn't a change.

Option Explicit

Private Const SH_MAIN As String = "What changed"
Private Const SH_LOG As String = "Log"
Private Const SH_SAVED As String = "Saved rates"
Private Const SH_HOW As String = "Read here"
Private Const BTN_NAME As String = "CheckRatesButton"
Private Const BTN_TEXT As String = "Check new rates file"
Private Const SITE As String = "www.streamlinedsalestax.org/ratesandboundry/Rates/"

' One rate row: the state file's nine columns, in its order. Dates are kept as yyyymmdd.
Private Const NC As Long = 9
Private Const R_ST As Long = 1, R_TYPE As Long = 2, R_CODE As Long = 3, R_GEN As Long = 4
Private Const R_GENX As Long = 5, R_FOOD As Long = 6, R_FOODX As Long = 7, R_BEGIN As Long = 8
Private Const R_END As Long = 9
Private Const SAVED_HEADS As String = "State|Type|Code|General rate|General rate, out-of-state|" & _
    "Food and drug rate|Food and drug rate, out-of-state|Begins|Ends"
Private Const SAVED_ROW As Long = 5         ' header row on Saved rates

Private Const LIST_HEADS As String = "Place|Kind|Code|What changed|Old rate|New rate|Starts|" & _
    "Old out-of-state rate|New out-of-state rate|Notes"
Private Const NL As Long = 10               ' columns in a list
Private Const LOG_HEADS As String = "Checked|File|Section|" & LIST_HEADS & "|Change ID"
Private Const LOG_COLS As Long = 14

' One change, as kept while sorting and drawing. 1 to 10 are the list's columns in order.
Private Const F_SORT As Long = 0, F_PLACE As Long = 1, F_KIND As Long = 2, F_CODE As Long = 3
Private Const F_WHAT As Long = 4, F_OLD As Long = 5, F_NEW As Long = 6, F_START As Long = 7
Private Const F_OLDX As Long = 8, F_NEWX As Long = 9, F_NOTES As Long = 10
Private Const F_LOOK As Long = 11, F_KEY As Long = 12, F_ID As Long = 13

Private Const EPS As Double = 0.0000001

Private mFeed As Collection     ' test runs only: answers for the boxes, in order
Private mSaid As String         ' test runs only: what the boxes would have said
Private mToday As Date          ' test runs only: the day to act as today
Private mFips As String         ' the state being shown, as its two-digit code
Private mNames As Collection    ' that state's place names, by type|code


' ---- the button --------------------------------------------------------------------

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

Private Sub LoadFile()
    Dim picked As Variant, fname As String, problem As String, fips As String, savedFips As String
    Dim recs As Variant, n As Long, today As Date, fileDay As String, nextDay As String, x As Long
    Dim newCur As Variant, nNew As Long, newIdx As Collection, dups As Long
    Dim oldAll As Variant, nOldAll As Long, oldCur As Variant, nOld As Long, oldIdx As Collection
    Dim savedFile As String, savedDay As String, savedWhen As Variant, newest As String, savedNewest As String
    Dim baseCur As Variant, nBase As Long, baseIdx As Collection
    Dim aheadCur As Variant, nAhead As Long, aheadIdx As Collection
    Dim first As Collection, ahead As Collection, food1 As Long, zip1 As Long, food2 As Long, zip2 As Long
    Dim zero1 As Long, zero2 As Long
    Dim compared As Boolean, label1 As String, label2 As String

    today = CheckDate()
    EnsureSheets
    picked = PickFile()
    If VarType(picked) = vbBoolean Then Exit Sub
    fname = FileNameOf(CStr(picked))
    If LCase$(Right$(fname, 4)) = ".zip" Then
        Say "That's the .zip file. Right-click it in Downloads, pick Extract All, then click the " & _
            "button again and pick the .csv file inside the new folder." & vbLf & vbLf & "Nothing was changed."
        Exit Sub
    End If

    Application.Cursor = xlWait
    Application.StatusBar = "Reading " & fname & "..."
    If Not ReadRates(CStr(picked), recs, n, fips, problem) Then
        Application.Cursor = xlDefault
        Say problem
        Exit Sub
    End If

    ' One state per workbook
    savedFips = SavedState()
    If savedFips <> "" And savedFips <> fips Then
        Application.Cursor = xlDefault
        Say "This workbook is " & StateName(savedFips) & ". That file is " & StateName(fips) & "." & _
            vbLf & vbLf & "Pick " & StateName(savedFips) & "'s rate file, or set up a separate workbook " & _
            "for " & StateName(fips) & ". Nothing was changed."
        Exit Sub
    End If
    mFips = fips
    LoadNames fips

    ' The rates in force on the first day of the quarter the file is for
    fileDay = DayOfFile(fname, today)
    Set newIdx = New Collection
    nNew = PickRows(recs, n, fileDay, newCur, newIdx, dups)
    If nNew = 0 Then
        Application.Cursor = xlDefault
        Say "That file has no rates in force on " & ShowDate8(fileDay) & ", the day it's for. Download " & _
            "the state's newest rate file and pick that." & vbLf & vbLf & "Nothing was changed."
        Exit Sub
    End If

    ' The file loaded last time, if there was one
    newest = NewestStart(recs, n)
    nOldAll = ReadSaved(oldAll, savedFile, savedDay, savedWhen, savedNewest)
    If nOldAll > 0 Then
        Set oldIdx = New Collection
        nOld = PickRows(oldAll, nOldAll, savedDay, oldCur, oldIdx, x)
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

    Application.StatusBar = "Comparing..."
    Set first = New Collection
    If nOldAll > 0 Then
        ' The must: this file's rates against the last file's
        compared = True
        Compare oldCur, nOld, oldIdx, newCur, nNew, newIdx, recs, n, fileDay, savedDay, False, first, food1, zip1, zero1
        label1 = "Since last file"
    Else
        ' First file: what the file itself changes on its day, against the day before
        Set baseIdx = New Collection
        nBase = PickRows(recs, n, DayBefore(fileDay), baseCur, baseIdx, x)
        Compare baseCur, nBase, baseIdx, newCur, nNew, newIdx, recs, n, fileDay, "", False, first, food1, zip1, zero1
        label1 = "On " & LongDate(fileDay)
    End If
    ' What the file already has for the quarter after
    nextDay = NextQuarter(fileDay)
    Set aheadIdx = New Collection
    nAhead = PickRows(recs, n, nextDay, aheadCur, aheadIdx, x)
    Set ahead = New Collection
    Compare newCur, nNew, newIdx, aheadCur, nAhead, aheadIdx, recs, n, nextDay, "", True, ahead, food2, zip2, zero2
    label2 = "Posted for " & LongDate(nextDay)

    Application.ScreenUpdating = False
    Application.StatusBar = "Drawing..."
    DrawFront fname, fileDay, today, compared, savedFile, savedDay, savedWhen, first, ahead, nextDay, _
        food1 + food2, zip1 + zip2, zero1 + zero2, dups, OutDiffers(newCur, nNew)
    AddToLog first, ahead, fname, today, label1, label2
    SaveCopy recs, n, fileDay, fname, today, nNew, newest
    DrawInstructions
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
    SheetByName(SH_MAIN).Range("A1").Select
End Sub


' ---- reading the state's file ---------------------------------------------------------

Private Function ReadRates(ByVal fullPath As String, ByRef recs As Variant, ByRef n As Long, _
                           ByRef fips As String, ByRef problem As String) As Boolean
    ' Reads the file into recs, one rate row per line. Lines that aren't rate rows (a header
    ' line, a blank) are skipped. False, with problem saying why, when it isn't a rate file.
    Dim body As String, lines As Variant, i As Long, f As Variant, nm As String, st As String
    Dim nonBlank As Long, bounds As Long, heads As Long, mixed As Boolean

    n = 0: fips = ""
    nm = FileNameOf(fullPath)
    On Error Resume Next
    body = ReadWholeFile(fullPath)
    If Err.Number <> 0 Then
        problem = "Couldn't open " & nm & ": " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    If Left$(body, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then body = Mid$(body, 4)
    If InStr(body, vbLf) = 0 Then body = Replace(body, vbCr, vbLf) Else body = Replace(body, vbCr, "")
    lines = Split(body, vbLf)
    If UBound(lines) < 0 Then lines = Array("")
    ReDim recs(1 To UBound(lines) + 1, 1 To NC)

    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            nonBlank = nonBlank + 1
            f = CsvFields(lines(i))
            If IsRateRow(f) Then
                n = n + 1
                st = Pad(f(0), 2)
                If fips = "" Then
                    fips = st
                ElseIf st <> fips Then
                    mixed = True
                End If
                recs(n, R_ST) = st
                recs(n, R_TYPE) = Pad(f(1), 2)
                recs(n, R_CODE) = CodeText(CStr(recs(n, R_TYPE)), f(2))
                recs(n, R_GEN) = Num(f(3))
                recs(n, R_GENX) = Num(f(4))
                recs(n, R_FOOD) = Num(f(5))
                recs(n, R_FOODX) = Num(f(6))
                recs(n, R_BEGIN) = Date8(f(7))
                recs(n, R_END) = Date8(f(8))
                If recs(n, R_END) = "" Then recs(n, R_END) = "99991231"
            ElseIf LooksLikeBoundary(f) Then
                bounds = bounds + 1
            ElseIf InStr(1, lines(i), "Jurisdiction", vbTextCompare) > 0 Then
                heads = heads + 1
            End If
        End If
    Next

    If n = 0 Then
        If nonBlank = 0 Then
            problem = nm & " is empty. Download the state's rate file again and pick that." & _
                vbLf & vbLf & "Nothing was changed."
        ElseIf bounds > 0 Then
            problem = "That's a boundary file, the one with B after the state's letters. This needs the " & _
                "rate file, the one with R after them, like OHR2026Q4SEP17." & vbLf & vbLf & "Nothing was changed."
        ElseIf heads > 0 Then
            problem = nm & " has the column names but no rates in it. Download the state's rate file " & _
                "again and pick that." & vbLf & vbLf & "Nothing was changed."
        Else
            problem = "That doesn't look like a Streamlined rate file. " & nm & " has no lines like " & _
                "39,00,001,0.015,0.015,0.015,0.015,20060401,99991231." & vbLf & vbLf & _
                "Get the state's rate file from " & SITE & " (its letters, then R) and pick that. " & _
                "Nothing was changed."
        End If
        Exit Function
    End If
    If mixed Then
        problem = "That file has rates for more than one state in it, so it isn't a state's rate file." & _
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

Private Function IsRateRow(ByVal f As Variant) As Boolean
    ' State, type, code, four rates, begin date, end date
    Dim i As Long, d As String
    If UBound(f) < 8 Then Exit Function
    If Not IsDigits(Trim$(f(0)), 1, 2) Then Exit Function
    If Not IsDigits(Trim$(f(1)), 1, 2) Then Exit Function
    If Len(Trim$(f(2))) = 0 Then Exit Function
    For i = 3 To 6
        If Not IsPlainNumber(Trim$(f(i))) Then Exit Function
    Next
    If Not IsDigits(Date8(f(7)), 8, 8) Then Exit Function
    d = Date8(f(8))
    If d <> "" Then If Not IsDigits(d, 8, 8) Then Exit Function
    IsRateRow = True
End Function

Private Function LooksLikeBoundary(ByVal f As Variant) As Boolean
    ' Boundary rows start with A, Z or 4, then two dates, and run to dozens of columns
    If UBound(f) < 20 Then Exit Function
    Select Case UCase$(Trim$(f(0)))
        Case "A", "Z", "4": LooksLikeBoundary = IsDigits(Trim$(f(1)), 8, 8)
    End Select
End Function

Private Function CodeText(ByVal jtype As String, ByVal raw As String) As String
    ' Codes the way the state writes them, leading zeros and all, even if the file went
    ' through Excel and lost them: 3 digits for a county, 2 for the state, 5 for the rest.
    raw = UCase$(Trim$(raw))
    If IsDigits(raw, 1, 10) Then
        Select Case jtype
            Case "00": raw = Pad(raw, 3)
            Case "45": raw = Pad(raw, 2)
            Case Else: raw = Pad(raw, 5)
        End Select
    End If
    CodeText = raw
End Function

Private Function Pad(ByVal s As String, ByVal w As Long) As String
    s = Trim$(s)
    If Len(s) < w And IsDigits(s, 1, w) Then s = String$(w - Len(s), "0") & s
    Pad = s
End Function

Private Function Date8(ByVal s As String) As String
    ' A date as yyyymmdd, the way the state writes it. Also takes 10/1/2026 and 2026-10-01.
    Dim p As Variant
    s = Trim$(s)
    If s = "" Then Exit Function
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)     ' drop a time on the end
    If Len(s) = 8 And IsDigits(s, 8, 8) Then Date8 = s: Exit Function
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

Private Function DayOfFile(ByVal fname As String, ByVal today As Date) As String
    ' The day whose rates count: the first day of a quarter. Each quarter's file goes up the
    ' month before, so it's the quarter a month from now, or the quarter in the file's name
    ' if that's later (OHR2026Q4SEP17 is Ohio's file for 2026's fourth quarter, October 1).
    ' A state only posts a new file when something changes, so an older file still holds
    ' the rates for now.
    Dim s As String, i As Long, q As Long, d As Date, named As String, soon As String
    s = UCase$(fname)
    For i = 1 To Len(s) - 8
        If Mid$(s, i + 2, 1) = "R" And Mid$(s, i + 7, 1) = "Q" Then
            If IsLetters(Mid$(s, i, 2)) And IsDigits(Mid$(s, i + 3, 4), 4, 4) Then
                q = Val(Mid$(s, i + 8, 1))
                If q >= 1 And q <= 4 Then
                    named = Format$(DateSerial(Val(Mid$(s, i + 3, 4)), (q - 1) * 3 + 1, 1), "yyyymmdd")
                    Exit For
                End If
            End If
        End If
    Next
    d = today + 31
    soon = Format$(DateSerial(Year(d), ((Month(d) - 1) \ 3) * 3 + 1, 1), "yyyymmdd")
    If named > soon Then DayOfFile = named Else DayOfFile = soon
End Function

Private Function NextQuarter(ByVal day8 As String) As String
    Dim d As Date
    d = DateFrom8(day8)
    NextQuarter = Format$(DateSerial(Year(d), ((Month(d) - 1) \ 3) * 3 + 4, 1), "yyyymmdd")
End Function

Private Function DayBefore(ByVal day8 As String) As String
    DayBefore = Format$(DateFrom8(day8) - 1, "yyyymmdd")
End Function


' ---- which rows count ------------------------------------------------------------------

Private Function PickRows(ByRef recs As Variant, ByVal n As Long, ByVal asOf As String, _
                          ByRef picked As Variant, ByVal idx As Object, ByRef dups As Long) As Long
    ' The rates in force on asOf (yyyymmdd): every row that had begun by then and hadn't ended.
    ' When two rows cover the same place, the one that began later counts. idx gets each key.
    Dim i As Long, c As Long, k As String, at As Long, cnt As Long
    dups = 0
    If n = 0 Then Exit Function
    ReDim picked(1 To n, 1 To NC)
    For i = 1 To n
        If recs(i, R_BEGIN) <= asOf And recs(i, R_END) >= asOf Then
            k = RowKey(recs, i)
            at = Found(idx, k)
            If at = 0 Then
                cnt = cnt + 1
                at = cnt
                idx.Add at, k
            Else
                dups = dups + 1
                If recs(i, R_BEGIN) < picked(at, R_BEGIN) Then GoTo NextRow
            End If
            For c = 1 To NC
                picked(at, c) = recs(i, c)
            Next
        End If
NextRow:
    Next
    PickRows = cnt
End Function

Private Function RowKey(ByRef recs As Variant, ByVal i As Long) As String
    ' A place is one jurisdiction type and code: county 001, city 00415, and so on
    RowKey = recs(i, R_TYPE) & "|" & recs(i, R_CODE)
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
        If recs(i, R_BEGIN) > best Then best = recs(i, R_BEGIN)
    Next
    NewestStart = best
End Function

Private Function OutDiffers(ByRef recs As Variant, ByVal n As Long) As Long
    ' How many places have an out-of-state rate that isn't the in-state rate
    Dim i As Long
    For i = 1 To n
        If Not IsZipLevel(CStr(recs(i, R_TYPE)), CStr(recs(i, R_CODE))) Then
            If Abs(recs(i, R_GEN) - recs(i, R_GENX)) > EPS Then OutDiffers = OutDiffers + 1
        End If
    Next
End Function


' ---- comparing -------------------------------------------------------------------------

Private Sub Compare(ByRef oldRecs As Variant, ByVal nOld As Long, ByVal oldIdx As Object, _
                    ByRef newRecs As Variant, ByVal nNew As Long, ByVal newIdx As Object, _
                    ByRef allRecs As Variant, ByVal nAll As Long, ByVal newDay As String, _
                    ByVal lastDay As String, ByVal ahead As Boolean, ByVal changes As Object, _
                    ByRef foodOnly As Long, ByRef zipOnly As Long, ByRef zeroOnly As Long)
    ' Every place that's new, gone, or has a different general rate between two sets of
    ' rates. allRecs is the newer file, for when a gone rate ended. lastDay is the last
    ' file's day, to spot a change dated back before it. The counts are what isn't listed:
    ' food and drug rate changes, the state's ZIP code default rates, and places added or
    ' dropped at 0%, which have no tax either way.
    Dim i As Long, j As Long, what As String, look As String, notes As String
    Dim goneAt As Collection, ended As Variant, v As Variant

    foodOnly = 0: zipOnly = 0: zeroOnly = 0
    For i = 1 To nNew
        j = Found(oldIdx, RowKey(newRecs, i))
        If j = 0 Then
            If IsZipLevel(CStr(newRecs(i, R_TYPE)), CStr(newRecs(i, R_CODE))) Then
                zipOnly = zipOnly + 1
            ElseIf IsZero(newRecs, i) Then
                zeroOnly = zeroOnly + 1
            Else
                changes.Add MakeChange("New", "new", LateNote(CStr(newRecs(i, R_BEGIN)), lastDay), _
                    oldRecs, 0, newRecs, i)
            End If
        ElseIf Differs(oldRecs, j, newRecs, i, what, look, notes) Then
            If IsZipLevel(CStr(newRecs(i, R_TYPE)), CStr(newRecs(i, R_CODE))) Then
                zipOnly = zipOnly + 1
            Else
                changes.Add MakeChange(what, look, JoinNotes(notes, LateNote(CStr(newRecs(i, R_BEGIN)), lastDay)), _
                    oldRecs, j, newRecs, i)
            End If
        ElseIf FoodDiffers(oldRecs, j, newRecs, i) Then
            If IsZipLevel(CStr(newRecs(i, R_TYPE)), CStr(newRecs(i, R_CODE))) Then
                zipOnly = zipOnly + 1
            Else
                foodOnly = foodOnly + 1
            End If
        End If
    Next

    ' Gone: in the old rates but not the new
    Set goneAt = New Collection
    For j = 1 To nOld
        If Found(newIdx, RowKey(oldRecs, j)) = 0 Then
            If IsZipLevel(CStr(oldRecs(j, R_TYPE)), CStr(oldRecs(j, R_CODE))) Then
                zipOnly = zipOnly + 1
            ElseIf IsZero(oldRecs, j) Then
                zeroOnly = zeroOnly + 1
            Else
                goneAt.Add j, RowKey(oldRecs, j)
            End If
        End If
    Next
    If goneAt.Count = 0 Then Exit Sub

    ' When it ended: the latest end date the newer file has for it before the new day
    ReDim ended(1 To nOld)
    For i = 1 To nAll
        j = Found(goneAt, RowKey(allRecs, i))
        If j > 0 Then
            If allRecs(i, R_END) < newDay And allRecs(i, R_END) > ended(j) Then ended(j) = allRecs(i, R_END)
        End If
    Next
    For Each v In goneAt
        j = v
        If ahead Then
            notes = "Ends " & ShowDate8(CStr(oldRecs(j, R_END))) & ". Nothing after it is posted yet."
        ElseIf ended(j) <> "" Then
            notes = "Ended " & ShowDate8(CStr(ended(j))) & "."
        Else
            notes = "No rate for it on " & ShowDate8(newDay) & " in the new file."
        End If
        changes.Add MakeChange("Gone", "gone", notes, oldRecs, j, newRecs, 0)
    Next
End Sub

Private Function Differs(ByRef oldRecs As Variant, ByVal j As Long, ByRef newRecs As Variant, _
                         ByVal i As Long, ByRef what As String, ByRef look As String, _
                         ByRef notes As String) As Boolean
    ' True when the general rate changed, in-state or out-of-state. what is the headline.
    Dim og As Double, ng As Double, ox As Double, nx As Double
    what = "": look = "": notes = ""
    og = oldRecs(j, R_GEN): ng = newRecs(i, R_GEN)
    ox = oldRecs(j, R_GENX): nx = newRecs(i, R_GENX)
    If Abs(ng - og) > EPS Then
        If ng > og Then what = "Went up": look = "up" Else what = "Went down": look = "down"
    ElseIf Abs(nx - ox) > EPS Then
        If nx > ox Then
            what = "Out-of-state rate went up": look = "up"
        Else
            what = "Out-of-state rate went down": look = "down"
        End If
    End If
    If what = "" Then Exit Function
    ' The food rate only gets a word when it isn't just the general rate again
    If FoodDiffers(oldRecs, j, newRecs, i) Then
        If Abs(oldRecs(j, R_FOOD) - og) > EPS Or Abs(newRecs(i, R_FOOD) - ng) > EPS Then
            notes = "Food and drug rate was " & Pct(CDbl(oldRecs(j, R_FOOD))) & ", now " & _
                Pct(CDbl(newRecs(i, R_FOOD))) & "."
        End If
    End If
    Differs = True
End Function

Private Function IsZero(ByRef recs As Variant, ByVal i As Long) As Boolean
    ' A place with no general tax, in-state or out-of-state
    IsZero = Abs(recs(i, R_GEN)) < EPS And Abs(recs(i, R_GENX)) < EPS
End Function

Private Function FoodDiffers(ByRef oldRecs As Variant, ByVal j As Long, ByRef newRecs As Variant, _
                             ByVal i As Long) As Boolean
    FoodDiffers = Abs(newRecs(i, R_FOOD) - oldRecs(j, R_FOOD)) > EPS Or _
        Abs(newRecs(i, R_FOODX) - oldRecs(j, R_FOODX)) > EPS
End Function

Private Function LateNote(ByVal begin8 As String, ByVal lastDay As String) As String
    ' A change dated on or before the last file's day is one the last file didn't have
    If lastDay = "" Then Exit Function
    If begin8 <= lastDay Then LateNote = "Dated before the last file's rates, so it's a late fix or add."
End Function

Private Function JoinNotes(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        JoinNotes = b
    ElseIf b = "" Then
        JoinNotes = a
    Else
        JoinNotes = a & " " & b
    End If
End Function

Private Function MakeChange(ByVal what As String, ByVal look As String, ByVal notes As String, _
                            ByRef oldRecs As Variant, ByVal j As Long, _
                            ByRef newRecs As Variant, ByVal i As Long) As Variant
    ' One line of a list. i is its row in the new rates and j in the old, 0 for none.
    Dim f(0 To 13) As Variant, t As String, c As String, k As String, st As String
    If i > 0 Then
        t = newRecs(i, R_TYPE): c = newRecs(i, R_CODE): st = newRecs(i, R_BEGIN)
    Else
        t = oldRecs(j, R_TYPE): c = oldRecs(j, R_CODE)
    End If
    k = t & "|" & c
    f(F_PLACE) = PlaceName(t, c)
    f(F_KIND) = KindWords(t)
    f(F_CODE) = c
    f(F_WHAT) = what
    If j > 0 Then
        f(F_OLD) = CDbl(oldRecs(j, R_GEN))
        f(F_OLDX) = CDbl(oldRecs(j, R_GENX))
    End If
    If i > 0 Then
        f(F_NEW) = CDbl(newRecs(i, R_GEN))
        f(F_NEWX) = CDbl(newRecs(i, R_GENX))
    End If
    If st <> "" Then f(F_START) = DateFrom8(st)
    f(F_NOTES) = notes
    f(F_LOOK) = look
    f(F_KEY) = k
    ' The same change seen again, in a later file, has the same ID
    f(F_ID) = k & "|" & what & "|" & st & "|" & NumText(f(F_NEW)) & "|" & NumText(f(F_NEWX))
    f(F_SORT) = UCase$(f(F_PLACE)) & Chr$(1) & UCase$(f(F_KIND)) & Chr$(1) & c
    MakeChange = f
End Function

Private Function SortedOrder(ByVal changes As Object) As Variant
    ' The changes by place name, so a place's changes sit together
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

Private Function ShowsOut(ByVal changes As Object) As Boolean
    ' True when a change in the list has an out-of-state rate that isn't its in-state rate
    Dim ch As Variant
    For Each ch In changes
        If Not IsEmpty(ch(F_OLD)) Then If Abs(ch(F_OLD) - ch(F_OLDX)) > EPS Then ShowsOut = True
        If Not IsEmpty(ch(F_NEW)) Then If Abs(ch(F_NEW) - ch(F_NEWX)) > EPS Then ShowsOut = True
    Next
End Function


' ---- the sheets ------------------------------------------------------------------------

Private Sub EnsureSheets()
    ' First run: makes the four sheets. Later runs put back the button and write Read here
    ' fresh, so it always matches this macro.
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
    If SheetByName(SH_HOW) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_MAIN))
        ws.Name = SH_HOW
    End If
    If SheetByName(SH_LOG) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_HOW))
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
    DrawInstructions
    DrawButton SheetByName(SH_MAIN)
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
End Sub

Private Sub DrawWelcome(ByVal ws As Worksheet)
    ClearSheet ws
    ListWidths ws
    ws.Cells.Font.Size = 14
    With ws.Range("A1")
        .Value = "Streamlined rates: what changed"
        .Font.Size = 26
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 50
    With ws.Range("A3")
        .Value = "Click the button and pick the state's rate file."
        .Font.Size = 20
    End With
    ws.Rows(3).RowHeight = 34
    ws.Range("A4").Value = "Get it from " & SITE & ": the state's two letters, then R, like OHR2026Q4SEP17.csv."
    ws.Rows(4).RowHeight = 24
    ws.Range("A5").Value = "The Read here tab has the steps, where to get the file, and what everything means."
    ws.Rows(5).RowHeight = 24
    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
    DrawButton ws
End Sub

Private Sub DrawInstructions()
    ' The Read here sheet: where to get the file, the steps each quarter, what the sheets
    ' mean, and how to set up another state. Each line is a kind and its words: T title,
    ' H heading, P a line, L the link to the list of files.
    Dim ws As Worksheet, st As String, w As Collection, v As Variant, r As Long, txt As String
    Set ws = SheetByName(SH_HOW)
    If ws Is Nothing Then Exit Sub
    st = SavedState()
    Set w = New Collection
    w.Add "T|Read here: how to work it"
    If st = "" Then
        w.Add "P|One workbook for each state, all with the same macro. This one takes its state from the " & _
            "first file you load, and after that it only takes that state's files."
    Else
        w.Add "P|One workbook for each state, all with the same macro. This one is " & StateName(st) & _
            "'s, so it only takes " & StateName(st) & "'s files."
    End If

    w.Add "H|Where to get the file"
    w.Add "P|Each quarter the states post a new rate file by the first of the month before the quarter: " & _
        "early September for October, early December for January, and so on."
    w.Add "P|Go to this page. It's a plain list with one file per state:"
    w.Add "L|" & SITE
    w.Add "P|Click the state's file: its two letters, then R, like OHR2026Q4SEP17.csv. Ohio is OHR, " & _
        "Washington WAR, Nebraska NER, Oklahoma OKR."
    If st <> "" Then w.Add "P|This workbook's file starts with " & StateAbbr(st) & "R."
    w.Add "P|Some come as a .zip. Washington, Nebraska and Oklahoma do. Right-click the .zip in Downloads, " & _
        "pick Extract All, and use the .csv file inside."
    w.Add "P|If the state's file has the same name as last quarter, the state posted nothing new."

    w.Add "H|Each quarter"
    w.Add "P|1. Download the state's new file, the way it says above."
    w.Add "P|2. On the What changed tab, click Check new rates file. The box opens in Downloads. Pick the .csv file."
    w.Add "P|3. Read What changed. The next part says what everything on it means."
    w.Add "P|4. Save the workbook. It never saves itself."

    w.Add "H|What the What changed tab means"
    w.Add "P|The big colored band says how many places changed since the last file you loaded. " & _
        "NO CHANGES means nothing you watch changed."
    w.Add "P|Each line is one place: its name, what kind of place it is, its code in the state's file, " & _
        "what changed, the old rate, the new rate, and the day the new rate starts."
    w.Add "P|Red went up. Green went down. Blue is new. Gray is gone."
    w.Add "P|The out-of-state rate is the rate on sales shipped in from another state. Those columns " & _
        "only show up when the state has places where it differs from the regular rate, like " & _
        "Oklahoma's city use taxes."
    w.Add "P|Already posted for: changes the file already has for the quarter after. Gone there means " & _
        "the rate ends and nothing after it is posted yet. The next file may post its new rate."
    w.Add "P|A late fix or add, in Notes, means the state changed something dated back before the " & _
        "last file. Worth a look."
    w.Add "P|The gray Not listed line counts what the lists leave out on purpose: food and drug rate " & _
        "changes, the state's ZIP code default rates, and places added or dropped at 0%."
    w.Add "P|The first file you load only sets the starting point. That time, What changed shows what " & _
        "the file itself changes on the first day of its quarter. Every file after that is checked " & _
        "against the one before."

    w.Add "H|The other tabs"
    w.Add "P|Log: every change once, with the day you loaded it and the file it came from. Use the " & _
        "filter buttons on top to find a place."
    w.Add "P|Saved rates: the last file's rates, for the next file to be checked against. It's locked. Leave it alone."

    w.Add "H|If something looks off"
    w.Add "P|A yellow bar says macros are disabled: click Enable Content."
    w.Add "P|It says the file looks older than last time, or looks cut short: answer No and download the file again."
    w.Add "P|It says this workbook is one state and that file is another: you picked the other state's " & _
        "file. Open that state's workbook instead."
    w.Add "P|It says it's a boundary file: you grabbed the file with B after the state's letters. " & _
        "Get the one with R."
    w.Add "P|It says it's a .zip file: right-click it, pick Extract All, and pick the .csv file inside."

    w.Add "H|Setting up another state"
    w.Add "P|1. Copy the macro from this workbook: press Alt F11, double-click Module1 on the left, " & _
        "then Ctrl A and Ctrl C. Close that window."
    w.Add "P|2. Open a new blank workbook. Press Alt F11, click Insert, then Module, then Ctrl V. Close that window."
    w.Add "P|3. File, Save As, Excel Macro-Enabled Workbook. Name it for the state."
    w.Add "P|4. Press Alt F8, pick CheckNewRatesFile, and click Run. It sets up the tabs and the button, " & _
        "then asks for that state's file. If you don't have it yet, click Cancel and use the button later."

    ClearSheet ws
    ws.Hyperlinks.Delete
    ws.Cells.Font.Size = 16
    ws.Columns(1).ColumnWidth = 110
    ws.Columns(1).WrapText = True
    ws.Cells.VerticalAlignment = xlTop
    r = 1
    For Each v In w
        txt = Mid$(v, 3)
        Select Case Left$(v, 1)
            Case "T"
                ws.Cells(r, 1).Value = txt
                ws.Cells(r, 1).Font.Size = 26
                ws.Cells(r, 1).Font.Bold = True
            Case "H"
                r = r + 1
                ws.Cells(r, 1).Value = txt
                ws.Cells(r, 1).Font.Size = 20
                ws.Cells(r, 1).Font.Bold = True
                ws.Cells(r, 1).Font.Color = RGB(31, 78, 121)
            Case "L"
                ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="https://" & txt, TextToDisplay:=txt
                ws.Cells(r, 1).Font.Size = 16
            Case Else
                ws.Cells(r, 1).Value = txt
        End Select
        r = r + 1
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(r, 1)).EntireRow.AutoFit
    ws.Rows(1).RowHeight = 50
    ThisWorkbook.Activate
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
End Sub

Private Sub DrawFront(ByVal fname As String, ByVal fileDay As String, ByVal loadedOn As Date, _
                      ByVal compared As Boolean, ByVal savedFile As String, ByVal savedDay As String, _
                      ByVal savedWhen As Variant, ByVal first As Object, ByVal ahead As Object, _
                      ByVal nextDay As String, ByVal foodOnly As Long, ByVal zipOnly As Long, _
                      ByVal zeroOnly As Long, ByVal dups As Long, ByVal outDiff As Long)
    ' What changed: which files, then the changes, in big type
    Dim ws As Worksheet, r As Long, s As String, n As Long, showOut As Boolean
    Dim fillColor As Long, textColor As Long

    showOut = (outDiff > 0) Or ShowsOut(first) Or ShowsOut(ahead)
    Set ws = SheetByName(SH_MAIN)
    ClearSheet ws
    ListWidths ws
    ws.Cells.Font.Size = 14
    ws.Cells.VerticalAlignment = xlCenter

    With ws.Range("A1")
        .Value = StateName(mFips) & " rates: what changed"
        .Font.Size = 26
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 50
    ws.Range("A2").Value = "From " & fname & ", loaded " & Format$(loadedOn, "m/d/yyyy") & ": rates as of " & _
        LongDate(fileDay) & "."
    If compared Then
        s = "Against " & savedFile & ", loaded " & ShowWhen(savedWhen) & ": rates as of " & LongDate(savedDay) & "."
    Else
        s = "First file in this workbook, so there's nothing to compare it to yet. The next file gets checked against this one."
    End If
    ws.Range("A3").Value = s
    ws.Range("A4").Value = "Showing each place's general sales tax rate. Food and drug rates aren't watched."
    If outDiff = 0 Then
        s = "Out-of-state rates (on sales shipped in from another state) are the same as in-state rates in this file."
    ElseIf outDiff = 1 Then
        s = "Out-of-state rates (on sales shipped in from another state) differ from in-state rates for 1 place in this file, so both are shown."
    Else
        s = "Out-of-state rates (on sales shipped in from another state) differ from in-state rates for " & _
            Format$(outDiff, "#,##0") & " places in this file, so both are shown."
    End If
    ws.Range("A5").Value = s
    r = 6
    s = NotListed(foodOnly, zipOnly, zeroOnly, dups)
    If s <> "" Then
        ws.Range("A6").Value = s
        r = 7
    End If
    ws.Range(ws.Cells(4, 1), ws.Cells(r - 1, 1)).Font.Color = RGB(89, 89, 89)
    ws.Range(ws.Cells(2, 1), ws.Cells(r - 1, 1)).RowHeight = 24
    ws.Rows(r).RowHeight = 12
    r = r + 1

    ' The changes: since the last file, or on a first file what the file itself changes
    n = first.Count
    If n > 0 Then
        fillColor = RGB(255, 217, 102): textColor = RGB(0, 0, 0)
    Else
        fillColor = RGB(198, 239, 206): textColor = RGB(0, 97, 0)
    End If
    If compared Then s = "SINCE THE LAST FILE" Else s = "ON " & UCase$(LongDate(fileDay))
    Band ws, r, s & ": " & CountText(n), fillColor, textColor, 28
    r = r + 1
    If compared And savedDay = fileDay Then
        If n > 0 Then
            s = "Both files have rates as of " & LongDate(fileDay) & ", so these are fixes the state made."
        Else
            s = "Both files have rates as of " & LongDate(fileDay) & ", and they match."
        End If
    ElseIf compared Then
        If n > 0 Then
            s = "This file's rates as of " & LongDate(fileDay) & " against the last file's as of " & _
                LongDate(savedDay) & "."
        Else
            s = "Nothing you watch changed between the last file's rates as of " & LongDate(savedDay) & _
                " and this file's as of " & LongDate(fileDay) & "."
        End If
    ElseIf n > 0 Then
        s = "What this file changes on " & LongDate(fileDay) & ", against the day before."
    Else
        s = "Nothing you watch changes on " & LongDate(fileDay) & " in this file."
    End If
    ws.Cells(r, 1).Value = s
    ws.Rows(r).RowHeight = 24
    r = r + 1
    If n > 0 Then DrawList ws, r, first

    ' What the file already has for the next quarter
    ws.Rows(r).RowHeight = 16
    r = r + 1
    If ahead.Count > 0 Then
        Band ws, r, "ALREADY POSTED FOR " & UCase$(LongDate(nextDay)) & ": " & CountText(ahead.Count), _
            RGB(221, 235, 247), RGB(31, 78, 121), 22
        r = r + 1
        ws.Cells(r, 1).Value = "The state posted these ahead. Gone means the rate ends and nothing after it is posted yet."
        ws.Rows(r).RowHeight = 24
        r = r + 1
        DrawList ws, r, ahead
    Else
        ws.Cells(r, 1).Value = "Nothing in this file changes on " & LongDate(nextDay) & "."
        ws.Rows(r).RowHeight = 24
    End If

    ws.Range(ws.Columns(8), ws.Columns(9)).Hidden = Not showOut
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

Private Function NotListed(ByVal foodOnly As Long, ByVal zipOnly As Long, ByVal zeroOnly As Long, _
                           ByVal dups As Long) As String
    ' The gray line for what the lists leave out
    Dim parts As Collection, s As String, i As Long
    Set parts = New Collection
    If zipOnly = 1 Then parts.Add "1 change in the state's ZIP code default rates"
    If zipOnly > 1 Then parts.Add Format$(zipOnly, "#,##0") & " changes in the state's ZIP code default rates"
    If zeroOnly = 1 Then parts.Add "1 place added or dropped at 0%"
    If zeroOnly > 1 Then parts.Add Format$(zeroOnly, "#,##0") & " places added or dropped at 0%"
    If foodOnly = 1 Then parts.Add "1 food and drug rate change"
    If foodOnly > 1 Then parts.Add Format$(foodOnly, "#,##0") & " food and drug rate changes"
    For i = 1 To parts.Count
        If i = 1 Then
            s = "Not listed: " & parts(i)
        ElseIf i = parts.Count Then
            s = s & " and " & parts(i)
        Else
            s = s & ", " & parts(i)
        End If
    Next
    If s <> "" Then s = s & "."
    If dups = 1 Then s = Trim$(s & " The file repeats 1 rate; the newer row was used.")
    If dups > 1 Then s = Trim$(s & " The file repeats " & Format$(dups, "#,##0") & " rates; the newest row of each was used.")
    NotListed = s
End Function

Private Function CountText(ByVal n As Long) As String
    If n = 0 Then
        CountText = "NO CHANGES"
    ElseIf n = 1 Then
        CountText = "1 CHANGE"
    Else
        CountText = Format$(n, "#,##0") & " CHANGES"
    End If
End Function

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
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, NL))
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .WrapText = True
    End With
    ws.Rows(r).RowHeight = 44
    order = SortedOrder(changes)
    ReDim out(1 To n, 1 To NL)
    For i = 1 To n
        ch = changes(order(i))
        For c = 1 To NL
            out(i, c) = ch(c)
        Next
    Next
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + n, 3)).NumberFormat = "@"
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, NL))
        .Value = out
        .Font.Size = 16
    End With
    ws.Range(ws.Cells(r + 1, 5), ws.Cells(r + n, 6)).NumberFormat = "0.00##%"
    ws.Range(ws.Cells(r + 1, 8), ws.Cells(r + n, 9)).NumberFormat = "0.00##%"
    ws.Range(ws.Cells(r + 1, 5), ws.Cells(r + n, 6)).Font.Bold = True
    ws.Range(ws.Cells(r + 1, 7), ws.Cells(r + n, 7)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + n, 3)).HorizontalAlignment = xlLeft
    For i = 1 To n
        ch = changes(order(i))
        rr = r + i
        NoTextMark ws.Cells(rr, 3)
        With ws.Range(ws.Cells(rr, 1), ws.Cells(rr, NL)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
        With ws.Cells(rr, 4)
            .Interior.Color = LookFill(CStr(ch(F_LOOK)))
            .Font.Color = LookText(CStr(ch(F_LOOK)))
            .Font.Bold = True
        End With
    Next
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).WrapText = True
    ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + n, 4)).WrapText = True
    ws.Range(ws.Cells(r + 1, NL), ws.Cells(r + n, NL)).WrapText = True
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).EntireRow.AutoFit
    For rr = r + 1 To r + n
        If ws.Rows(rr).RowHeight < 30 Then ws.Rows(rr).RowHeight = 30
    Next
    r = r + n + 1
End Sub

Private Sub NoTextMark(ByVal cell As Range)
    ' Codes are text on purpose, so Excel's green corner mark for numbers kept as text goes
    On Error Resume Next
    cell.Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
End Sub

Private Sub ListWidths(ByVal ws As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(44, 18, 10, 26, 12, 12, 14, 16, 16, 50)
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next
End Sub

Private Sub AddToLog(ByVal first As Object, ByVal ahead As Object, ByVal fname As String, _
                     ByVal loadedOn As Date, ByVal label1 As String, ByVal label2 As String)
    ' Each change goes on the Log once, the first time a file shows it. A check with nothing
    ' new still gets a line.
    Dim ws As Worksheet, r As Long, last As Long, ids As Collection, v As Variant, i As Long
    Dim out As Variant, n As Long, lists As Variant, labels As Variant, s As Long
    Dim order As Variant, ch As Variant, c As Long, lst As Object, looks() As String

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

    lists = Array(first, ahead)
    labels = Array(label1, label2)
    ReDim out(1 To first.Count + ahead.Count + 1, 1 To LOG_COLS)
    ReDim looks(0 To 0)
    For s = 0 To 1
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
                    For c = 1 To NL
                        out(n, c + 3) = ch(c)
                    Next
                    out(n, LOG_COLS) = ch(F_ID)
                    ReDim Preserve looks(0 To n)
                    looks(n) = CStr(ch(F_LOOK))
                End If
            Next
        End If
    Next
    If n = 0 Then
        n = 1
        out(1, 1) = loadedOn
        out(1, 2) = fname
        out(1, 7) = "Nothing new"
        out(1, 13) = "Everything this file shows was already on the Log."
        ReDim Preserve looks(0 To 1)
        looks(1) = ""
    End If

    r = last + 1
    If r < 2 Then r = 2
    ws.Range(ws.Cells(r, 6), ws.Cells(r + n - 1, 6)).NumberFormat = "@"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, LOG_COLS))
        .Value = out
        .Font.Size = 14
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 1)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 10), ws.Cells(r + n - 1, 10)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 8), ws.Cells(r + n - 1, 9)).NumberFormat = "0.00##%"
    ws.Range(ws.Cells(r, 11), ws.Cells(r + n - 1, 12)).NumberFormat = "0.00##%"
    ws.Range(ws.Cells(r, 6), ws.Cells(r + n - 1, 6)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(r, 13), ws.Cells(r + n - 1, 13)).WrapText = True
    For i = 1 To n
        NoTextMark ws.Cells(r + i - 1, 6)
        If looks(i) <> "" Then
            With ws.Cells(r + i - 1, 7)
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
    widths = Array(13, 34, 26, 40, 16, 10, 30, 11, 11, 13, 13, 13, 44, 10)
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

Private Sub SaveCopy(ByRef recs As Variant, ByVal n As Long, ByVal fileDay As String, ByVal fname As String, _
                     ByVal loadedOn As Date, ByVal inForce As Long, ByVal newest As String)
    ' Saved rates: the file's rows still in force on its day or later, for the next file to
    ' be checked against. The state's code sits in B2 and decides which files fit.
    Dim ws As Worksheet, out As Variant, i As Long, c As Long, titles As Variant, cnt As Long
    Set ws = SheetByName(SH_SAVED)
    ws.Unprotect
    ws.Cells.Clear
    ws.Cells.Font.Size = 11
    With ws.Range("A1")
        .Value = "The rates from the last file loaded. The next file gets checked against them. " & _
            "Loading a file replaces them, so leave this sheet alone."
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Range("A2").Value = "State"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = mFips
    ws.Range("C2").Value = StateName(mFips)
    ws.Range("A3").Value = "File"
    ws.Range("B3").NumberFormat = "@"
    ws.Range("B3").Value = fname
    ws.Range("C3").Value = "Rates as of"
    ws.Range("D3").Value = DateFrom8(fileDay)
    ws.Range("D3").NumberFormat = "m/d/yyyy"
    ws.Range("E3").Value = "Loaded"
    ws.Range("F3").Value = loadedOn
    ws.Range("F3").NumberFormat = "m/d/yyyy"
    ws.Range("G3").Value = "Rates in force"
    ws.Range("H3").Value = inForce
    ws.Range("I3").Value = "Newest start"
    ws.Range("J3").Value = DateFrom8(newest)
    ws.Range("J3").NumberFormat = "m/d/yyyy"
    ws.Range("A2,A3,C3,E3,G3,I3").Font.Bold = True
    titles = Split(SAVED_HEADS, "|")
    For c = 1 To NC
        ws.Cells(SAVED_ROW, c).Value = titles(c - 1)
    Next
    ws.Range(ws.Cells(SAVED_ROW, 1), ws.Cells(SAVED_ROW, NC)).Font.Bold = True
    ReDim out(1 To n, 1 To NC)
    For i = 1 To n
        If recs(i, R_END) >= fileDay Then
            cnt = cnt + 1
            For c = 1 To NC
                out(cnt, c) = recs(i, c)
            Next
        End If
    Next
    ' Codes and dates stay text, leading zeros and all
    ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + cnt, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(SAVED_ROW + 1, 8), ws.Cells(SAVED_ROW + cnt, 9)).NumberFormat = "@"
    ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + cnt, NC)).Value = out
    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 8
    ws.Columns(3).ColumnWidth = 10
    ws.Range(ws.Columns(4), ws.Columns(NC)).ColumnWidth = 16
    ws.Protect
End Sub

Private Function ReadSaved(ByRef recs As Variant, ByRef savedFile As String, ByRef savedDay As String, _
                           ByRef savedWhen As Variant, ByRef savedNewest As String) As Long
    ' Every row on Saved rates. 0 when nothing's saved yet.
    Dim ws As Worksheet, last As Long, n As Long, i As Long, c As Long, v As Variant
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Cells(SAVED_ROW, 1).Value) <> "State" Then Exit Function
    If Not IsDate(ws.Range("D3").Value) Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last <= SAVED_ROW Then Exit Function
    n = last - SAVED_ROW
    v = ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(last, NC)).Value
    ReDim recs(1 To n, 1 To NC)
    For i = 1 To n
        For c = 1 To NC
            Select Case c
                Case R_GEN, R_GENX, R_FOOD, R_FOODX
                    recs(i, c) = Num(CStr(v(i, c)))
                Case Else
                    recs(i, c) = Trim$(CStr(v(i, c)))
            End Select
        Next
    Next
    savedFile = CStr(ws.Range("B3").Value)
    savedDay = Format$(CDate(ws.Range("D3").Value), "yyyymmdd")
    savedWhen = ws.Range("F3").Value
    If IsDate(ws.Range("J3").Value) Then savedNewest = Format$(CDate(ws.Range("J3").Value), "yyyymmdd")
    ReadSaved = n
End Function

Private Function SavedState() As String
    ' The state this workbook is for, once a file has been loaded
    Dim ws As Worksheet
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Cells(SAVED_ROW, 1).Value) <> "State" Then Exit Function
    SavedState = Trim$(CStr(ws.Range("B2").Value))
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.RowHeight = ws.StandardHeight
    ws.Cells.EntireColumn.Hidden = False
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
    b.Left = ws.Range("E1").Left + 10
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


' ---- states, places and kinds ----------------------------------------------------------

Private Function StateName(ByVal fips As String) As String
    ' The state's name from its two-digit code
    Dim all As String, p As Long
    all = "|01Alabama|02Alaska|04Arizona|05Arkansas|06California|08Colorado|09Connecticut|10Delaware|" & _
        "11District of Columbia|12Florida|13Georgia|15Hawaii|16Idaho|17Illinois|18Indiana|19Iowa|20Kansas|" & _
        "21Kentucky|22Louisiana|23Maine|24Maryland|25Massachusetts|26Michigan|27Minnesota|28Mississippi|" & _
        "29Missouri|30Montana|31Nebraska|32Nevada|33New Hampshire|34New Jersey|35New Mexico|36New York|" & _
        "37North Carolina|38North Dakota|39Ohio|40Oklahoma|41Oregon|42Pennsylvania|44Rhode Island|" & _
        "45South Carolina|46South Dakota|47Tennessee|48Texas|49Utah|50Vermont|51Virginia|53Washington|" & _
        "54West Virginia|55Wisconsin|56Wyoming|"
    p = InStr(all, "|" & fips)
    If Len(fips) = 2 And p > 0 Then
        StateName = Mid$(all, p + 3, InStr(p + 1, all, "|") - p - 3)
    Else
        StateName = "state " & fips
    End If
End Function

Private Function StateAbbr(ByVal fips As String) As String
    ' The state's two letters from its two-digit code
    Dim all As String, p As Long
    all = "|01AL|02AK|04AZ|05AR|06CA|08CO|09CT|10DE|11DC|12FL|13GA|15HI|16ID|17IL|18IN|19IA|20KS|" & _
        "21KY|22LA|23ME|24MD|25MA|26MI|27MN|28MS|29MO|30MT|31NE|32NV|33NH|34NJ|35NM|36NY|37NC|38ND|" & _
        "39OH|40OK|41OR|42PA|44RI|45SC|46SD|47TN|48TX|49UT|50VT|51VA|53WA|54WV|55WI|56WY|"
    p = InStr(all, "|" & fips)
    If Len(fips) = 2 And p > 0 Then StateAbbr = Mid$(all, p + 3, 2)
End Function

Private Function KindWords(ByVal jtype As String) As String
    ' The kind of place, from the file's jurisdiction type
    Select Case jtype
        Case "45": KindWords = "State"
        Case "00": KindWords = "County"
        Case "01": KindWords = "City"
        Case "02": KindWords = "Town"
        Case "03": KindWords = "Village"
        Case "04": KindWords = "Borough"
        Case "05": KindWords = "Township"
        Case "09": KindWords = "Other city or town"
        Case "10": KindWords = "School district"
        Case "63", "69"
            Select Case mFips
                Case "39": KindWords = "Transit"
                Case "31": KindWords = "Good Life district"
                Case "53": KindWords = "Location code"
                Case Else: KindWords = "Special district"
            End Select
        Case Else
            If Val(jtype) >= 20 And Val(jtype) <= 29 Then KindWords = "Special district" Else KindWords = "Type " & jtype
    End Select
End Function

Private Function IsZipLevel(ByVal jtype As String, ByVal code As String) As Boolean
    ' Some states post a default rate for each ZIP code area, for when only the ZIP is known.
    ' Each city and county still has its own row, so these aren't listed, only counted.
    If jtype <> "63" Then Exit Function
    Select Case mFips
        Case "40": IsZipLevel = (Left$(code, 2) = "98" Or Left$(code, 2) = "99")
        Case "53": IsZipLevel = (Left$(code, 2) = "L9")
    End Select
End Function

Private Sub LoadNames(ByVal fips As String)
    ' The state's place names, from the lists at the bottom of this module
    Dim all As Variant, e As Variant, p As Variant
    Set mNames = New Collection
    Select Case fips
        Case "39": all = Split(NamesOH(), ";")
        Case "31": all = Split(NamesNE(), ";")
        Case "40": all = Split(NamesOK(), ";")
        Case "53": all = Split(NamesWA(), ";")
        Case Else: Exit Sub
    End Select
    On Error Resume Next
    For Each e In all
        p = Split(CStr(e), "|")
        If UBound(p) = 2 Then mNames.Add CStr(p(2)), CStr(p(0)) & "|" & CStr(p(1))
    Next
    On Error GoTo 0
End Sub

Private Function PlaceName(ByVal jtype As String, ByVal code As String) As String
    ' The place's name, or its kind and code when the lists don't have it
    Dim nm As String
    If jtype = "45" Then PlaceName = StateName(mFips): Exit Function
    If Not mNames Is Nothing Then
        On Error Resume Next
        nm = mNames(jtype & "|" & code)
        On Error GoTo 0
    End If
    If nm = "" Then nm = KindWords(jtype) & " " & code
    PlaceName = nm
End Function


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

Private Function Num(ByVal s As String) As Double
    Num = Val(Replace(Replace(Replace(Trim$(s), "$", ""), "%", ""), ",", ""))
End Function

Private Function NumText(ByVal v As Variant) As String
    ' A rate the same way however the file wrote it: 0.01500 and .015 both come out .015
    If IsEmpty(v) Then NumText = "" Else NumText = Trim$(Str$(v))
End Function

Private Function Pct(ByVal x As Double) As String
    Pct = Format$(x * 100, "0.00##") & "%"
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

Private Function IsDigits(ByVal s As String, ByVal minLen As Long, ByVal maxLen As Long) As Boolean
    Dim i As Long, ch As String
    If Len(s) < minLen Or Len(s) > maxLen Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    IsDigits = True
End Function

Private Function IsLetters(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next
    IsLetters = True
End Function

Private Function DateFrom8(ByVal s As String) As Date
    DateFrom8 = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 5, 2)), Val(Mid$(s, 7, 2)))
End Function

Private Function ShowDate8(ByVal s As String) As String
    If Len(s) = 8 Then ShowDate8 = Format$(DateFrom8(s), "m/d/yyyy") Else ShowDate8 = s
End Function

Private Function LongDate(ByVal s As String) As String
    ' October 1, 2026
    If Len(s) = 8 Then LongDate = Format$(DateFrom8(s), "mmmm d, yyyy") Else LongDate = s
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

Private Function BoxTitle() As String
    Dim st As String
    st = SavedState()
    If st = "" Then BoxTitle = "Streamlined rates" Else BoxTitle = StateName(st) & " rates"
End Function

Private Function PickFile() As Variant
    ' The file picked, or False if the box was cancelled. Starts in Downloads when there is one.
    Dim dl As String, st As String, title As String
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
    st = SavedState()
    If st = "" Then title = "Pick the state's rate file" Else title = "Pick the new " & StateName(st) & " rate file"
    PickFile = Application.GetOpenFilename("Rate file (*.csv;*.zip),*.csv;*.zip", 1, title)
End Function

Private Sub Say(ByVal msg As String)
    If Not mFeed Is Nothing Then
        mSaid = mSaid & msg & vbLf
    Else
        MsgBox msg, vbInformation, BoxTitle()
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
    Confirm = (MsgBox(question, vbYesNo + vbQuestion + vbDefaultButton2, BoxTitle()) = vbYes)
End Function

Public Function RatesTest(ByVal answers As String, ByVal todayText As String) As String
    ' For test runs: feeds answers to the boxes (split on |), acts as if today were todayText
    ' (yyyy-mm-dd, blank for the real day), and returns what the boxes would have said.
    ' Not needed day to day.
    Dim a As Variant
    Set mFeed = New Collection
    mSaid = ""
    If todayText <> "" Then
        mToday = DateSerial(Val(Left$(todayText, 4)), Val(Mid$(todayText, 6, 2)), Val(Mid$(todayText, 9, 2)))
    Else
        mToday = 0
    End If
    If answers <> "" Then
        For Each a In Split(answers, "|")
            mFeed.Add CStr(a)
        Next
    End If
    CheckNewRatesFile
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    RatesTest = mSaid
    Set mFeed = Nothing
    mToday = 0
End Function


' ---- place names (made by build_names.py from the Census and state lists) ----

Private Function NamesOH1() As String
    Dim s As String
    s = s & "00|001|Adams County;00|003|Allen County;00|005|Ashland County;00|007|Ashtabula County;00|009|Athens County;00|011|Auglaize County;00|013|Belmont County;00|015|Brown County;00|017|Butler County;00|019|Carroll County;00|021|Champaign County;00|023|Clark County;00|025|Clermont County;00|027|Clinton County;00|029|Columbiana County;00|031|Coshocton County;00|033|Crawford County;00|035|Cuyahoga County;00|037|Darke County;00|039|Defiance County;00|041|Delaware County;00|043|Erie County;00|045|Fairfield County;00|047|Fayette County;00|049|Franklin County;00|051|Fulton County;00|053|Gallia County;00|055|Geauga County;00|057|Greene County;00|059|Guernsey County;00|061|Hamilton County;00|063|Hancock County;00|065|Hardin County;00|067|Harrison County;00|069|Henry County;00|071|Highland County;00|073|Hocking County;00|075|Holmes County;00|077|Huron County;00|079|Jackson County;"
    s = s & "00|081|Jefferson County;00|083|Knox County;00|085|Lake County;00|087|Lawrence County;00|089|Licking County;00|091|Logan County;00|093|Lorain County;00|095|Lucas County;00|097|Madison County;00|099|Mahoning County;00|101|Marion County;00|103|Medina County;00|105|Meigs County;00|107|Mercer County;00|109|Miami County;00|111|Monroe County;00|113|Montgomery County;00|115|Morgan County;00|117|Morrow County;00|119|Muskingum County;00|121|Noble County;00|123|Ottawa County;00|125|Paulding County;00|127|Perry County;00|129|Pickaway County;00|131|Pike County;00|133|Portage County;00|135|Preble County;00|137|Putnam County;00|139|Richland County;00|141|Ross County;00|143|Sandusky County;00|145|Scioto County;00|147|Seneca County;00|149|Shelby County;00|151|Stark County;00|153|Summit County;00|155|Trumbull County;00|157|Tuscarawas County;00|159|Union County;00|161|Van Wert County;00|163|Vinton County;"
    s = s & "00|165|Warren County;00|167|Washington County;00|169|Wayne County;00|171|Williams County;00|173|Wood County;00|175|Wyandot County;63|02000|Allen County transit;63|18000|Cuyahoga County transit (GCRTA);63|25000|Franklin County transit (COTA);63|31000|Hamilton County transit (Metro);63|43000|Lake County transit (Laketran);63|48000|Lucas County transit (TARTA);63|50000|Mahoning County transit (WRTA);63|57000|Montgomery County transit (RTA);63|67000|Portage County transit (PARTA);63|76000|Stark County transit (SARTA);63|77000|Summit County transit (METRO);63|87100|Wood County, TARTA part;63|93000|Fairfield County, COTA part;63|94000|Licking County, COTA part;63|96000|Delaware County, COTA part;63|98000|Union County, COTA part;"
    NamesOH1 = s
End Function

Private Function NamesOH() As String
    NamesOH = NamesOH1()
End Function

Private Function NamesNE1() As String
    Dim s As String
    s = s & "00|001|Adams County;00|003|Antelope County;00|005|Arthur County;00|007|Banner County;00|009|Blaine County;00|011|Boone County;00|013|Box Butte County;00|015|Boyd County;00|017|Brown County;00|019|Buffalo County;00|021|Burt County;00|023|Butler County;00|025|Cass County;00|027|Cedar County;00|029|Chase County;00|031|Cherry County;00|033|Cheyenne County;00|035|Clay County;00|037|Colfax County;00|039|Cuming County;00|041|Custer County;00|043|Dakota County;00|045|Dawes County;00|047|Dawson County;00|049|Deuel County;00|051|Dixon County;00|053|Dodge County;00|055|Douglas County;00|057|Dundy County;00|059|Fillmore County;00|061|Franklin County;00|063|Frontier County;00|065|Furnas County;00|067|Gage County;00|069|Garden County;00|071|Garfield County;00|073|Gosper County;00|075|Grant County;00|077|Greeley County;00|079|Hall County;00|081|Hamilton County;00|083|Harlan County;00|085|Hayes County;"
    s = s & "00|087|Hitchcock County;00|089|Holt County;00|091|Hooker County;00|093|Howard County;00|095|Jefferson County;00|097|Johnson County;00|099|Kearney County;00|101|Keith County;00|103|Keya Paha County;00|105|Kimball County;00|107|Knox County;00|109|Lancaster County;00|111|Lincoln County;00|113|Logan County;00|115|Loup County;00|117|McPherson County;00|119|Madison County;00|121|Merrick County;00|123|Morrill County;00|125|Nance County;00|127|Nemaha County;00|129|Nuckolls County;00|131|Otoe County;00|133|Pawnee County;00|135|Perkins County;00|137|Phelps County;00|139|Pierce County;00|141|Platte County;00|143|Polk County;00|145|Red Willow County;00|147|Richardson County;00|149|Rock County;00|151|Saline County;00|153|Sarpy County;00|155|Saunders County;00|157|Scotts Bluff County;00|159|Seward County;00|161|Sheridan County;00|163|Sherman County;00|165|Sioux County;00|167|Stanton County;"
    s = s & "00|169|Thayer County;00|171|Thomas County;00|173|Thurston County;00|175|Valley County;00|177|Washington County;00|179|Wayne County;00|181|Webster County;00|183|Wheeler County;00|185|York County;01|00205|Abie village;01|00240|Adams village;01|00415|Ainsworth city;01|00555|Albion city;01|00625|Alda village;01|00730|Alexandria village;01|00870|Allen village;01|00905|Alliance city;01|00975|Alma city;01|01150|Alvo village;01|01290|Amherst village;01|01465|Anoka village;01|01500|Anselmo village;01|01535|Ansley village;01|01780|Arapahoe city;01|01850|Arcadia village;01|01990|Arlington village;01|02095|Arnold village;01|02200|Arthur village;01|02305|Ashland city;01|02375|Ashton village;01|02550|Atkinson city;01|02620|Atlanta village;01|02655|Auburn city;01|02690|Aurora city;01|02795|Avoca village;01|02830|Axtell village;01|02865|Ayr village;01|03005|Bancroft village;01|03040|Barada village;"
    s = s & "01|03075|Barneston village;01|03145|Bartlett village;01|03180|Bartley village;01|03215|Bassett city;01|03250|Battle Creek city;01|03285|Bayard city;01|03355|Bazile Mills village;01|03390|Beatrice city;01|03495|Beaver City city;01|03530|Beaver Crossing village;01|03600|Bee village;01|03635|Beemer village;01|03775|Belden village;01|03810|Belgrade village;01|03950|Bellevue city;01|04020|Bellwood village;01|04160|Belvidere village;01|04195|Benedict village;01|04230|Benkelman city;01|04300|Bennet village;01|04405|Bennington city;01|04615|Bertrand village;01|04650|Berwyn village;01|04895|Big Springs village;01|05140|Bladen village;01|05350|Blair city;01|05455|Bloomfield city;01|05490|Bloomington village;01|05560|Blue Hill city;01|05630|Blue Springs city;01|06015|Boys Town village;01|06050|Bradshaw village;01|06085|Brady village;01|06120|Brainard village;01|06260|Brewster village;"
    s = s & "01|06295|Bridgeport city;01|06400|Bristow village;01|06470|Broadwater village;01|06505|Brock village;01|06610|Broken Bow city;01|06750|Brownville village;01|06785|Brule village;01|06820|Bruning village;01|06855|Bruno village;01|06890|Brunswick village;01|07065|Burchard village;01|07205|Burr village;01|07310|Burton village;01|07345|Burwell city;01|07415|Bushnell village;01|07485|Butte village;01|07555|Byron village;01|07625|Cairo village;01|07660|Callaway village;01|07730|Cambridge city;01|07870|Campbell village;01|07975|Carleton village;01|08010|Carroll village;01|08150|Cedar Bluffs village;01|08185|Cedar Creek village;01|08220|Cedar Rapids village;01|08360|Center village;01|08535|Central City city;01|08570|Ceresco village;01|08605|Chadron city;01|08675|Chambers village;01|08780|Chapman village;01|08885|Chappell city;01|09095|Chester village;01|09165|Clarks village;"
    s = s & "01|09200|Clarkson city;01|09270|Clatonia village;01|09375|Clay Center city;01|09445|Clearwater village;01|09655|Clinton village;01|09760|Cody village;01|09865|Coleridge village;01|10005|Colon village;01|10110|Columbus city;01|10180|Comstock village;01|10250|Concord village;01|10390|Cook village;01|10495|Cordova village;01|10600|Cornlea village;01|10635|Cortland village;01|10740|Cotesfield village;01|10985|Cowles village;01|11020|Cozad city;01|11055|Crab Orchard village;01|11090|Craig village;01|11195|Crawford city;01|11230|Creighton city;01|11300|Creston village;01|11370|Crete city;01|11440|Crofton city;01|11475|Crookston village;01|11615|Culbertson village;01|11825|Curtis city;01|11860|Cushing village;01|12000|Dakota City city;01|12070|Dalton village;01|12105|Danbury village;01|12140|Dannebrog village;01|12245|Davenport village;01|12280|Davey village;01|12315|David City city;"
    s = s & "01|12420|Dawson village;01|12455|Daykin village;01|12525|Decatur village;01|12770|Denton village;01|12840|Deshler city;01|12945|Deweese village;01|13015|De Witt village;01|13085|Diller village;01|13190|Dix village;01|13225|Dixon village;01|13295|Dodge village;01|13365|Doniphan village;01|13435|Dorchester village;01|13505|Douglas village;01|13750|Du Bois village;01|13855|Dunbar village;01|13890|Duncan village;01|13960|Dunning village;01|14065|Dwight village;01|14100|Eagle village;01|14380|Eddyville village;01|14450|Edgar city;01|14520|Edison village;01|14555|Elba village;01|14730|Elgin city;01|14975|Elk Creek village;01|15360|Elm Creek village;01|15430|Elmwood village;01|15500|Elsie village;01|15570|Elwood village;01|15605|Elyria village;01|15710|Emerson village;01|15815|Emmet village;01|15920|Endicott village;01|16025|Ericson village;01|16165|Eustis village;01|16270|Ewing village;"
    s = s & "01|16340|Exeter village;01|16410|Fairbury city;01|16445|Fairfield city;01|16550|Fairmont village;01|16655|Falls City city;01|16725|Farnam village;01|16760|Farwell village;01|16830|Filley village;01|16935|Firth village;01|17110|Fordyce village;01|17145|Fort Calhoun city;01|17320|Foster village;01|17530|Franklin city;01|17670|Fremont city;01|17775|Friend city;01|17810|Fullerton city;01|17880|Funk village;01|17950|Gandy village;01|18230|Garland village;01|18300|Garrison village;01|18405|Geneva city;01|18475|Genoa city;01|18580|Gering city;01|18615|Gibbon city;01|18720|Gilead village;01|18825|Giltner village;01|19070|Glenvil village;01|19245|Goehner village;01|19350|Gordon city;01|19385|Gothenburg city;01|19455|Grafton village;01|19595|Grand Island city;01|19910|Grant city;01|20050|Greeley Center village;"
    NamesNE1 = s
End Function

Private Function NamesNE2() As String
    Dim s As String
    s = s & "01|20190|Greenwood village;01|20225|Gresham village;01|20260|Gretna city;01|20295|Gross village;01|20365|Guide Rock village;01|20435|Gurley village;01|20470|Hadar village;01|20540|Haigler village;01|20610|Hallam village;01|20680|Halsey village;01|20750|Hamlet village;01|20785|Hampton village;01|20960|Harbine village;01|20995|Hardy village;01|21240|Harrison village;01|21275|Hartington city;01|21345|Harvard city;01|21415|Hastings city;01|21660|Hayes Center village;01|21730|Hay Springs village;01|21765|Hazard village;01|21835|Heartwell village;01|21905|Hebron city;01|22045|Hemingford village;01|22080|Henderson city;01|22115|Hendley village;01|22150|Henry village;01|22185|Herman village;01|22290|Hershey village;01|22325|Hickman city;01|22430|Hildreth village;01|22605|Holbrook village;01|22640|Holdrege city;01|22815|Holstein village;01|22920|Homer village;01|23025|Hooper city;"
    s = s & "01|23200|Hordville village;01|23235|Hoskins village;01|23270|Howard City (Boelus) village;01|23340|Howells village;01|23375|Hubbard village;01|23410|Hubbell village;01|23445|Humboldt city;01|23480|Humphrey city;01|23550|Huntley village;01|23655|Hyannis village;01|23690|Imperial city;01|23830|Indianola city;01|23970|Inglewood village;01|24075|Inman village;01|24285|Ithaca village;01|24355|Jackson village;01|24530|Jansen village;01|24670|Johnson village;01|24740|Johnstown village;01|24915|Julian village;01|24950|Juniata village;01|25055|Kearney city;01|25160|Kenesaw village;01|25230|Kennard village;01|25405|Kilgore village;01|25475|Kimball city;01|26175|Lamar village;01|26350|Laurel city;01|26385|La Vista city;01|26420|Lawrence village;01|26455|Lebanon village;01|26560|Leigh village;01|26630|Leshara village;01|26805|Lewellen village;01|26875|Lewiston village;01|26910|Lexington city;"
    s = s & "01|26980|Liberty village;01|28000|Lincoln city;01|28105|Lindsay village;01|28245|Linwood village;01|28350|Litchfield village;01|28420|Lodgepole village;01|29050|Long Pine city;01|29085|Loomis village;01|29175|Lorton village;01|29260|Louisville city;01|29470|Loup City city;01|29645|Lushton village;01|29715|Lyman village;01|29750|Lynch village;01|29855|Lyons city;01|29925|McCook city;01|29960|McCool Junction village;01|30030|McGrew village;01|30065|McLean village;01|30240|Madison city;01|30275|Madrid village;01|30310|Magnet village;01|30345|Malcolm village;01|30380|Malmo village;01|30415|Manley village;01|30730|Marquette village;01|30940|Martinsburg village;01|31045|Maskell village;01|31115|Mason City village;01|31220|Maxwell village;01|31325|Maywood village;01|31395|Mead village;01|31465|Meadow Grove village;01|31570|Melbeta village;01|31640|Memphis village;01|31780|Merna village;"
    s = s & "01|31815|Merriman village;01|32060|Milford city;01|32130|Miller village;01|32200|Milligan village;01|32305|Minatare city;01|32340|Minden city;01|32410|Mitchell city;01|32550|Monowi village;01|32585|Monroe village;01|32760|Moorefield village;01|32830|Morrill village;01|32865|Morse Bluff village;01|33250|Mullen village;01|33320|Murdock village;01|33425|Murray village;01|33530|Naper village;01|33565|Naponee village;01|33705|Nebraska City city;01|33740|Nehawka village;01|33775|Neligh city;01|33880|Nelson city;01|33950|Nemaha village;01|33985|Nenzel village;01|34090|Newcastle village;01|34230|Newman Grove city;01|34265|Newport village;01|34300|Nickerson village;01|34370|Niobrara village;01|34545|Nora village;01|34615|Norfolk city;01|34650|Norman village;01|34720|North Bend city;01|34825|North Loup village;01|35000|North Platte city;01|35245|Oak village;01|35385|Oakdale village;"
    s = s & "01|35490|Oakland city;01|35560|Obert village;01|35665|Oconto village;01|35700|Octavia village;01|35735|Odell village;01|35980|Ogallala city;01|36015|Ohiowa village;01|37000|Omaha city;01|37105|O'Neill city;01|37140|Ong village;01|37210|Orchard village;01|37280|Ord city;01|37420|Orleans village;01|37525|Osceola city;01|37560|Oshkosh city;01|37595|Osmond city;01|37630|Otoe village;01|37770|Overton village;01|37910|Oxford village;01|38085|Page village;01|38130|Palisade village;01|38160|Palmer village;01|38190|Palmyra village;01|38225|Panama village;01|38295|Papillion city;01|38575|Pawnee City city;01|38610|Paxton village;01|38750|Pender village;01|38960|Peru city;01|38995|Petersburg village;01|39030|Phillips village;01|39065|Pickrell village;01|39100|Pierce city;01|39135|Pilger village;01|39170|Plainview city;01|39310|Platte Center village;01|39345|Plattsmouth city;"
    s = s & "01|39380|Pleasant Dale village;01|39450|Pleasanton village;01|39590|Plymouth village;01|39660|Polk village;01|39695|Ponca city;01|39870|Potter village;01|39975|Prague village;01|40325|Preston village;01|40360|Primrose village;01|40430|Prosser village;01|40570|Ragan village;01|40605|Ralston city;01|40675|Randolph city;01|40710|Ravenna city;01|40780|Raymond village;01|40920|Red Cloud city;01|41130|Republican City village;01|41235|Reynolds village;01|41375|Richland village;01|41480|Rising City village;01|41515|Riverdale village;01|41760|Riverton village;01|41830|Roca village;01|42075|Rockville village;01|42145|Rogers village;01|42250|Rosalie village;01|42355|Roseland village;01|42495|Royal village;01|42670|Rulo village;01|42775|Rushville city;01|42810|Ruskin village;01|43055|St. Edward city;01|43090|St. Helena village;01|43265|St. Paul city;01|43370|Salem village;01|43475|Santee village;"
    s = s & "01|43685|Sargent city;01|43755|Saronville village;01|44035|Schuyler city;01|44070|Scotia village;01|44245|Scottsbluff city;01|44280|Scribner city;01|44420|Seward city;01|44595|Shelby village;01|44700|Shelton village;01|45085|Shickley village;01|45190|Sholes village;01|45225|Shubert village;01|45295|Sidney city;01|45400|Silver Creek village;01|45575|Smithfield village;01|45610|Snyder village;01|45680|South Bend village;01|46030|South Sioux City city;01|46135|Spalding village;01|46275|Spencer village;01|46380|Sprague village;01|46520|Springfield city;01|46625|Springview village;01|46695|Stamford village;01|46800|Stanton city;01|46835|Staplehurst village;01|46870|Stapleton village;01|47010|Steele City village;01|47080|Steinauer village;01|47115|Stella village;01|47150|Sterling village;01|47220|Stockham village;01|47290|Stockville village;01|47360|Strang village;01|47395|Stratton village;"
    s = s & "01|47465|Stromsburg city;01|47500|Stuart village;01|47675|Sumner village;01|47815|Superior city;01|47850|Surprise village;01|47920|Sutherland village;01|47955|Sutton city;01|48060|Swanton village;01|48235|Syracuse city;01|48270|Table Rock village;01|48305|Talmage village;01|48410|Tarnov village;01|48445|Taylor village;01|48480|Tecumseh city;01|48515|Tekamah city;01|48585|Terrytown city;01|48690|Thayer village;01|48760|Thedford village;01|48900|Thurston village;01|48935|Tilden city;01|49005|Tobias village;01|49145|Trenton village;01|49285|Trumbull village;01|49425|Uehling village;01|49460|Ulysses village;01|49530|Unadilla village;01|49635|Union village;01|49880|Upland village;01|49915|Utica village;01|49950|Valentine city;01|50020|Valley city;01|50125|Valparaiso village;01|50230|Venango village;01|50335|Verdel village;01|50370|Verdigre village;01|50510|Verdon village;"
    s = s & "01|50790|Virginia village;01|50895|Waco village;01|50965|Wahoo city;01|51070|Wakefield city;01|51175|Wallace village;01|51245|Walthill village;01|51595|Washington village;01|51630|Waterbury village;01|51665|Waterloo village;01|51700|Wauneta village;01|51735|Wausa village;01|51770|Waverly city;"
    NamesNE2 = s
End Function

Private Function NamesNE3() As String
    Dim s As String
    s = s & "01|51840|Wayne city;01|52015|Weeping Water city;01|52085|Wellfleet village;01|52295|Western village;01|52540|Weston village;01|52575|West Point city;01|52925|Whitney village;01|52960|Wilber city;01|52995|Wilcox village;01|53240|Wilsonville village;01|53275|Winnebago village;01|53345|Winnetoon village;01|53380|Winside village;01|53415|Winslow village;01|53450|Wisner city;01|53520|Wolbach village;01|53555|Wood Lake village;01|53660|Wood River city;01|53835|Wymore city;01|53905|Wynot village;01|54045|York city;01|54115|Yutan city;63|GL801|Good Life District, Avenue One opted-in retailers;63|GL802|Good Life District GL802;63|GL803|Good Life District GL803;63|GL804|Good Life District GL804;63|GL805|Good Life District GL805;"
    NamesNE3 = s
End Function

Private Function NamesNE() As String
    NamesNE = NamesNE1() & NamesNE2() & NamesNE3()
End Function

Private Function NamesOK1() As String
    Dim s As String
    s = s & "00|001|Adair County;00|003|Alfalfa County;00|005|Atoka County;00|007|Beaver County;00|009|Beckham County;00|011|Blaine County;00|013|Bryan County;00|015|Caddo County;00|017|Canadian County;00|019|Carter County;00|021|Cherokee County;00|023|Choctaw County;00|025|Cimarron County;00|027|Cleveland County;00|029|Coal County;00|031|Comanche County;00|033|Cotton County;00|035|Craig County;00|037|Creek County;00|039|Custer County;00|041|Delaware County;00|043|Dewey County;00|045|Ellis County;00|047|Garfield County;00|049|Garvin County;00|051|Grady County;00|053|Grant County;00|055|Greer County;00|057|Harmon County;00|059|Harper County;00|061|Haskell County;00|063|Hughes County;00|065|Jackson County;00|067|Jefferson County;00|069|Johnston County;00|071|Kay County;00|073|Kingfisher County;00|075|Kiowa County;00|077|Latimer County;00|079|Le Flore County;00|081|Lincoln County;00|083|Logan County;"
    s = s & "00|085|Love County;00|087|McClain County;00|089|McCurtain County;00|091|McIntosh County;00|093|Major County;00|095|Marshall County;00|097|Mayes County;00|099|Murray County;00|101|Muskogee County;00|103|Noble County;00|105|Nowata County;00|107|Okfuskee County;00|109|Oklahoma County;00|111|Okmulgee County;00|113|Osage County;00|115|Ottawa County;00|117|Pawnee County;00|119|Payne County;00|121|Pittsburg County;00|123|Pontotoc County;00|125|Pottawatomie County;00|127|Pushmataha County;00|129|Roger Mills County;00|131|Rogers County;00|133|Seminole County;00|135|Sequoyah County;00|137|Stephens County;00|139|Texas County;00|141|Tillman County;00|143|Tulsa County;00|145|Wagoner County;00|147|Washington County;00|149|Washita County;00|151|Woods County;00|153|Woodward County;01|00100|Achille town;01|00200|Ada city;01|00250|Adair town;01|00300|Adams CDP;01|00350|Adamson CDP;01|00450|Addington town;"
    s = s & "01|00600|Afton town;01|00700|Agra town;01|00950|Albany CDP;01|01050|Albion town;01|01150|Alderson town;01|01250|Alex town;01|01350|Aline town;01|01400|Allen town;01|01700|Altus city;01|01800|Alva city;01|01900|Amber town;01|01950|Ames town;01|02000|Amorita town;01|02050|Anadarko city;01|02250|Antlers city;01|02300|Apache town;01|02500|Arapaho town;01|02550|Arcadia town;01|02600|Ardmore city;01|02650|Arkoma town;01|02750|Armstrong town;01|02800|Arnett town;01|03000|Asher town;01|03050|Ashland town;01|03300|Atoka city;01|03350|Atwood town;01|03500|Avant town;01|03550|Avard CDP;01|03750|Bache CDP;01|03900|Baker CDP;01|04200|Barber CDP;01|04250|Barnsdall city;01|04450|Bartlesville city;01|04700|Bearden town;01|04750|Beaver town;01|05000|Beggs city;01|05400|Bennington town;01|05450|Bentley CDP;01|05550|Bernice town;01|05650|Bessie town;01|05700|Bethany city;01|05800|Bethel Acres town;"
    s = s & "01|05900|Big Cabin town;01|06100|Billings town;01|06200|Binger town;01|06350|Bison CDP;01|06400|Bixby city;01|06450|Blackburn town;01|06600|Blackwell city;01|06650|Blair town;01|06700|Blanchard city;01|06750|Blanco CDP;01|06900|Blue CDP;01|06950|Bluejacket town;01|07300|Boise City city;01|07350|Bokchito town;01|07450|Bokoshe town;01|07500|Boley town;01|07750|Boswell town;01|07950|Bowlegs town;01|08000|Bowring CDP;01|08150|Boynton town;01|08300|Bradley town;01|08400|Braggs town;01|08450|Braman town;01|08550|Bray town;01|08600|Breckenridge town;01|08700|Briartown CDP;01|08725|Bridge Creek town;01|08750|Bridgeport city;01|08900|Bristow city;01|09050|Broken Arrow city;01|09100|Broken Bow city;01|09150|Bromide town;01|09400|Brooksville town;01|09850|Buffalo town;01|09900|Bunch CDP;01|09950|Burbank town;01|10050|Burlington town;01|10150|Burneyville CDP;01|10250|Burns Flat town;"
    s = s & "01|10450|Butler town;01|10550|Byars town;01|10600|Byng town;01|10650|Byron town;01|10700|Cache city;01|10800|Caddo town;01|10950|Calera town;01|11050|Calumet town;01|11100|Calvin town;01|11150|Camargo town;01|11300|Cameron town;01|11450|Canadian town;01|11550|Caney town;01|11600|Canton town;01|11650|Canute town;01|11750|Capron town;01|11990|Carlton Landing town;01|12000|Carmen town;01|12050|Carnegie town;01|12100|Carney town;01|12200|Carrier town;01|12300|Carter town;01|12550|Cartwright CDP;01|12650|Cashion town;01|12750|Castle town;01|12900|Catoosa city;01|12960|Cedar Valley city;01|13000|Cement town;01|13100|Centrahoma city;01|13135|Central High town;01|13500|Chandler city;01|13600|Chattanooga town;01|13650|Checotah city;01|13700|Chelsea town;01|13750|Cherokee city;01|13800|Chester CDP;01|13900|Cheyenne town;01|13950|Chickasha city;01|14200|Choctaw city;01|14300|Chouteau town;"
    s = s & "01|14500|Cimarron City town;01|14700|Claremore city;01|14750|Clarita CDP;01|14900|Clayton town;01|15050|Clearview town;01|15250|Cleora CDP;01|15300|Cleo Springs town;01|15350|Cleveland city;01|15400|Clinton city;01|15800|Coalgate city;01|16050|Colbert town;01|16100|Colcord town;01|16200|Cole town;01|16250|Coleman CDP;01|16350|Collinsville city;01|16400|Colony town;01|16450|Comanche city;01|16500|Commerce city;01|16700|Connerville CDP;01|16950|Cookson CDP;01|17050|Cooperton town;01|17100|Copan town;01|17300|Corn town;01|17350|Cornish town;01|17550|Council Hill town;01|17700|Covington town;01|17800|Coweta city;01|17850|Cowlington town;01|17950|Coyle town;01|18250|Crescent city;01|18450|Cromwell town;01|18550|Crowder town;01|18750|Cumberland CDP;01|18850|Cushing city;01|18950|Custer City town;01|19000|Cyril town;01|19050|Dacoma town;01|19150|Dale CDP;01|19350|Davenport town;"
    s = s & "01|19400|Davidson town;01|19450|Davis city;01|19650|Deer Creek town;01|19850|Delaware town;01|19900|Del City city;01|20300|Depew town;01|20450|Devol town;01|20500|Dewar town;01|20550|Dewey city;01|20650|Dibble town;01|20700|Dickson town;01|20850|Dill City town;01|20900|Disney town;01|21200|Dougherty town;01|21250|Douglas town;01|21350|Dover town;01|21700|Drummond town;01|21750|Drumright city;01|21900|Duncan city;01|22050|Durant city;01|22100|Durham CDP;01|22200|Dustin town;01|22300|Eagle City CDP;01|22350|Eagletown CDP;01|22400|Eakly town;01|22500|Earlsboro town;01|22550|East Duke town;01|23200|Edmond city;01|23350|Eldorado town;01|23450|Elgin city;01|23500|Elk City city;01|23550|Elmer town;01|23600|Elmore City town;01|23700|El Reno city;01|23900|Empire City town;01|23950|Enid city;01|24200|Erick city;01|24300|Erin Springs town;01|24460|Etowah town;01|24650|Eufaula city;"
    s = s & "01|24850|Fairfax town;01|24900|Fairland town;01|24950|Fairmont town;01|25000|Fair Oaks town;01|25100|Fairview city;01|25250|Fallis town;01|25400|Fanshawe town;01|25450|Fargo town;01|25650|Faxon town;01|25700|Fay CDP;01|25850|Felt CDP;01|26000|Finley CDP;01|26200|Fittstown CDP;01|26250|Fitzhugh town;01|26350|Fletcher town;01|26750|Foraker town;01|26850|Forest Park town;"
    NamesOK1 = s
End Function

Private Function NamesOK2() As String
    Dim s As String
    s = s & "01|26900|Forgan town;01|27100|Fort Cobb town;01|27150|Fort Coffee town;01|27200|Fort Gibson town;01|27350|Fort Supply town;01|27400|Fort Towson town;01|27450|Foss town;01|27500|Foster town;01|27550|Fox CDP;01|27600|Foyil town;01|27650|Francis town;01|27800|Frederick city;01|27850|Freedom town;01|27900|Friendship town;01|28250|Gage town;01|28350|Gans town;01|28500|Garber city;01|28700|Garvin town;01|28800|Gate town;01|28900|Geary city;01|28950|Gene Autry town;01|29100|Geronimo town;01|29150|Gerty town;01|29400|Glencoe town;01|29600|Glenpool city;01|29800|Golden CDP;01|29850|Goldsby town;01|29900|Goltry town;01|30200|Goodwell town;01|30300|Gore town;01|30350|Gotebo town;01|30400|Gould town;01|30500|Gowen CDP;01|30600|Gracemont town;01|30750|Grainola town;01|30850|Grandfield city;01|30900|Grand Lake Towne town;01|30950|Granite town;01|31000|Grant CDP;01|31150|Grayson town;"
    s = s & "01|31250|Greenfield town;01|31600|Grove city;01|31700|Guthrie city;01|31750|Guymon city;01|31900|Haileyville city;01|32050|Hallett town;01|32250|Hammon town;01|32350|Hanna town;01|32550|Hardesty town;01|32750|Harrah city;01|32850|Hartshorne city;01|32900|Haskell town;01|33000|Hastings town;01|33150|Haworth town;01|33250|Haywood CDP;01|33300|Headrick town;01|33350|Healdton city;01|33400|Heavener city;01|33450|Helena town;01|33600|Hendrix town;01|33650|Hennepin CDP;01|33700|Hennessey town;01|33750|Henryetta city;01|34150|Hickory town;01|34650|Hillsdale town;01|34750|Hinton town;01|34850|Hitchcock town;01|34900|Hitchita town;01|35000|Hobart city;01|35030|Hochatown CDP;01|35150|Hodgen CDP;01|35300|Hoffman town;01|35400|Holdenville city;01|35500|Hollis city;01|35550|Hollister town;01|35800|Homestead CDP;01|35850|Hominy city;01|36000|Hooker city;01|36020|Hoot Owl town;01|36050|Hopeton CDP;"
    s = s & "01|36090|Horntown town;01|36150|Howe town;01|36200|Hoyt CDP;01|36300|Hugo city;01|36400|Hulbert town;01|36600|Hunter town;01|36700|Hydro town;01|36750|Idabel city;01|36850|Indiahoma town;01|36950|Indianola town;01|37100|Inola town;01|37350|Isabella CDP;01|37410|IXL town;01|37650|Jay city;01|37750|Jefferson town;01|37800|Jenks city;01|37850|Jennings town;01|38000|Jet town;01|38200|Johnson town;01|38350|Jones town;01|38600|Kansas town;01|38700|Katie town;01|38800|Kaw City city;01|38850|Keefton CDP;01|39000|Kellyville town;01|39050|Kemp town;01|39150|Kendrick town;01|39200|Kenefic town;01|39400|Kenton CDP;01|39450|Kenwood CDP;01|39500|Keota town;01|39550|Ketchum town;01|39600|Keyes town;01|39750|Kiefer town;01|39800|Kildare town;01|39850|Kingfisher city;01|39900|Kingston town;01|39950|Kinta town;01|40000|Kiowa town;01|40100|Knowles town;01|40200|Konawa city;01|40300|Krebs city;"
    s = s & "01|40400|Kremlin town;01|40700|Lahoma town;01|40800|Lake Aluma town;01|41250|Lamar town;01|41300|Lambert town;01|41350|Lamont town;01|41400|Lane CDP;01|41500|Langley town;01|41550|Langston town;01|41700|Laverne town;01|41760|Lawrence Creek town;01|41850|Lawton city;01|41900|Leach CDP;01|41950|Lebanon CDP;01|42050|Leedey town;01|42150|Le Flore town;01|42200|Lehigh city;01|42300|Lenapah town;01|42450|Leon town;01|42500|Leonard CDP;01|42600|Lequire CDP;01|42700|Lexington city;01|42860|Liberty town;01|43000|Lima town;01|43150|Lindsay city;01|43450|Loco town;01|43500|Locust Grove town;01|43725|Lone Chimney;01|43750|Lone Grove city;01|43850|Lone Wolf town;01|43950|Longdale town;01|44000|Lookeba town;01|44100|Lotsee town;01|44150|Loveland town;01|44300|Loyal town;01|44350|Lucien CDP;01|44500|Luther town;01|44800|McAlester city;01|44950|McCurtain town;01|45350|McLoud city;01|45500|Macomb town;"
    s = s & "01|45750|Madill city;01|46000|Manchester town;01|46050|Mangum city;01|46150|Manitou town;01|46200|Mannford town;01|46300|Mannsville town;01|46400|Maramec town;01|46450|Marble City town;01|46500|Marietta city;01|46550|Marland town;01|46600|Marlow city;01|46650|Marshall town;01|46700|Martha town;01|46900|Maud city;01|47000|May town;01|47150|Maysville town;01|47200|Mazie CDP;01|47250|Mead town;01|47300|Medford city;01|47350|Medicine Park town;01|47450|Meeker town;01|47750|Meno town;01|47800|Meridian town;01|48000|Miami city;01|48350|Midwest City city;01|48400|Milburn town;01|48450|Milfay CDP;01|48500|Mill Creek town;01|48600|Millerton town;01|48750|Minco city;01|49000|Moffett town;01|49050|Monroe CDP;01|49100|Moodys CDP;01|49200|Moore city;01|49250|Mooreland town;01|49400|Morris city;01|49450|Morrison town;01|49550|Mounds town;01|49600|Mountain Park town;01|49650|Mountain View town;"
    s = s & "01|49750|Moyers CDP;01|49850|Muldrow town;01|49900|Mulhall town;01|50050|Muskogee city;01|50100|Mustang city;01|50150|Mutual town;01|50250|Nardin CDP;01|50300|Nash town;01|50350|Nashoba CDP;01|51050|New Alluwe town;01|51150|Newcastle city;01|51250|New Cordell city;01|51300|Newkirk city;01|51750|New Woodville town;01|51800|Nichols Hills city;01|51850|Nicoma Park city;01|52100|Ninnekah town;01|52150|Noble city;01|52450|Norge town;01|52500|Norman city;01|52600|North Enid town;01|52700|North Miami town;01|52900|Nowata city;01|53210|Oak Grove town;01|53350|Oakhurst CDP;01|53400|Oakland town;01|53550|Oaks town;01|53600|Oakwood town;01|53700|Ochelata town;01|54000|Oilton city;01|54050|Okarche town;01|54100|Okay town;01|54150|Okeene town;01|54200|Okemah city;01|55000|Oklahoma City city;01|55150|Okmulgee city;01|55200|Oktaha town;01|55450|Olustee town;01|55750|Oologah town;01|55850|Optima town;"
    s = s & "01|56050|Orlando town;01|56150|Osage town;01|56650|Owasso city;01|56800|Paden town;01|56900|Panama town;01|56950|Panola CDP;01|57100|Paoli town;01|57150|Paradise Hill town;01|57300|Park Hill CDP;01|57550|Pauls Valley city;01|57600|Pawhuska city;01|57650|Pawnee city;01|57900|Peckham CDP;01|57950|Peggs CDP;01|58050|Pensacola town;01|58100|Peoria town;01|58150|Perkins city;01|58250|Perry city;01|58500|Phillips town;01|58700|Piedmont city;01|59150|Pink town;01|59350|Pittsburg town;01|59500|Platter CDP;01|59700|Pocasset town;01|59750|Pocola town;01|59850|Ponca City city;01|59950|Pond Creek city;01|60000|Pontotoc CDP;01|60150|Porter town;01|60250|Porum town;01|60350|Poteau city;01|60500|Prague city;01|60650|Preston CDP;01|60750|Proctor CDP;01|60850|Prue town;01|61000|Pryor Creek city;01|61150|Purcell city;01|61250|Putnam town;01|61400|Quapaw town;01|61450|Quay CDP;01|61500|Quinlan CDP;"
    s = s & "01|61550|Quinton town;01|61650|Ralston town;01|61700|Ramona town;01|61850|Randlett town;01|61900|Ratliff City town;01|61950|Rattan town;01|62000|Ravia town;01|62200|Redbird town;01|62550|Red Oak town;01|62650|Red Rock town;01|62900|Renfrow town;01|62950|Rentiesville town;01|63100|Reydon town;01|63350|Ringling town;01|63450|Ringwood town;01|63500|Ripley town;"
    NamesOK2 = s
End Function

Private Function NamesOK3() As String
    Dim s As String
    s = s & "01|63600|Rock Island town;01|63700|Rocky town;01|63750|Roff town;01|63800|Roland town;01|63900|Roosevelt town;01|63950|Rose CDP;01|64000|Rosedale town;01|64050|Rosston town;01|64300|Rush Springs town;01|64600|Ryan town;01|64850|St. Louis town;01|64950|Salina town;01|65000|Sallisaw city;01|65300|Sand Springs city;01|65400|Sapulpa city;01|65499|Sapulpa - Tulsa County;01|65500|Sasakwa town;01|65600|Savanna town;01|65650|Sawyer town;01|65700|Sayre city;01|65900|Schulter town;01|65950|Scipio CDP;01|66250|Seiling city;01|66275|Selah;01|66300|Selman CDP;01|66350|Seminole city;01|66400|Sentinel town;01|66500|Seward CDP;01|66550|Shady Point town;01|66600|Shamrock town;01|66650|Sharon town;01|66750|Shattuck town;01|66800|Shawnee city;01|67200|Shidler city;01|67400|Short CDP;01|67550|Silo town;01|67800|Skedee town;01|67850|Skiatook city;01|67950|Slaughterville town;01|68000|Slick town;"
    s = s & "01|68200|Smith Village town;01|68250|Smithville town;01|68400|Snyder city;01|68500|Soper town;01|68600|South Coffeyville town;01|68950|Sparks town;01|69000|Spaulding town;01|69050|Spavinaw town;01|69200|Spencer city;01|69300|Sperry town;01|69350|Spiro town;01|69400|Sportsmen Acres town;01|69500|Springer town;01|70150|Sterling town;01|70200|Stidham town;01|70250|Stigler city;01|70300|Stillwater city;01|70350|Stilwell city;01|70500|Stonewall town;01|70750|Strang town;01|70800|Stratford town;01|70850|Stringtown town;01|70950|Strong City town;01|71000|Stroud city;01|71050|Stuart town;01|71250|Sugden town;01|71350|Sulphur city;01|71450|Summit town;01|71500|Sumner CDP;01|71900|Sweetwater town;01|71950|Swink CDP;01|72050|Taft town;01|72100|Tahlequah city;01|72200|Talala town;01|72250|Talihina town;01|72350|Taloga town;01|72400|Tamaha town;01|72500|Tatums town;01|72650|Tecumseh city;"
    s = s & "01|72750|Temple town;01|72850|Terlton town;01|72900|Terral town;01|73100|Texhoma town;01|73150|Texola town;01|73200|Thackerville town;01|73250|The Village city;01|73450|Thomas city;01|73850|Tipton town;01|73900|Tishomingo city;01|74150|Tonkawa city;01|74400|Tribbey town;01|74550|Tryon town;01|74650|Tullahassee town;01|75000|Tulsa city;01|75050|Tupelo city;01|75150|Turley CDP;01|75200|Turpin CDP;01|75250|Tushka town;01|75300|Tuskahoma CDP;01|75450|Tuttle city;01|75600|Twin Oaks CDP;01|75750|Tyrone town;01|75950|Union City town;01|76350|Utica CDP;01|76500|Valley Brook town;01|76600|Valley Park town;01|76650|Valliant town;01|76850|Vanoss CDP;01|76900|Velma town;01|76950|Vera town;01|77000|Verden town;01|77050|Verdigris town;01|77150|Vernon;01|77250|Vian town;01|77300|Vici town;01|77550|Vinita city;01|77850|Wagoner city;01|77900|Wainwright town;01|77950|Wakita town;01|78150|Walters city;"
    s = s & "01|78200|Wanette town;01|78250|Wann town;01|78300|Wapanucka town;01|78350|Wardville CDP;01|78400|Warner town;01|78500|Warr Acres city;01|78650|Warwick town;01|78700|Washington town;01|78750|Washita CDP;01|78950|Watonga city;01|79100|Watts town;01|79200|Waukomis town;01|79250|Waurika city;01|79300|Wayne town;01|79350|Waynoka city;01|79450|Weatherford city;01|79600|Webb City town;01|79650|Webbers Falls town;01|79750|Welch town;01|79800|Weleetka town;01|79850|Welling CDP;01|79900|Wellston town;01|80000|Welty CDP;01|80250|Westport town;01|80350|West Siloam Springs town;01|80450|Westville town;01|80500|Wetumka city;01|80550|Wewoka city;01|80800|Whitefield town;01|80900|Whitesboro CDP;01|81000|Wilburton city;01|81300|Willow town;01|81500|Wilson city;01|81575|Winchester town;01|81750|Wister town;01|82000|Woodlawn Park town;01|82150|Woodward city;01|82200|Wright City town;"
    s = s & "01|82250|Wyandotte town;01|82400|Wynnewood city;01|82450|Wynona town;01|82500|Yale city;01|82650|Yeager town;01|82950|Yukon city;"
    NamesOK3 = s
End Function

Private Function NamesOK() As String
    NamesOK = NamesOK1() & NamesOK2() & NamesOK3()
End Function

Private Function NamesWA1() As String
    Dim s As String
    s = s & "00|001|Adams County;00|003|Asotin County;00|005|Benton County;00|007|Chelan County;00|009|Clallam County;00|011|Clark County;00|013|Columbia County;00|015|Cowlitz County;00|017|Douglas County;00|019|Ferry County;00|021|Franklin County;00|023|Garfield County;00|025|Grant County;00|027|Grays Harbor County;00|029|Island County;00|031|Jefferson County;00|033|King County;00|035|Kitsap County;00|037|Kittitas County;00|039|Klickitat County;00|041|Lewis County;00|043|Lincoln County;00|045|Mason County;00|047|Okanogan County;00|049|Pacific County;00|051|Pend Oreille County;00|053|Pierce County;00|055|San Juan County;00|057|Skagit County;00|059|Skamania County;00|061|Snohomish County;00|063|Spokane County;00|065|Stevens County;00|067|Thurston County;00|069|Wahkiakum County;00|071|Walla Walla County;00|073|Whatcom County;00|075|Whitman County;00|077|Yakima County;01|00100|Aberdeen city;"
    s = s & "01|00905|Airway Heights city;01|01010|Albion town;01|01290|Algona city;01|01500|Almira town;01|01990|Anacortes city;01|02585|Arlington city;01|03075|Asotin city;01|03180|Auburn city;01|03736|Bainbridge Island city;01|04475|Battle Ground city;01|04895|Beaux Arts Village town;01|05210|Bellevue city;01|05280|Bellingham city;01|05560|Benton City city;01|06085|Bingen city;01|06330|Black Diamond city;01|06505|Blaine city;01|07170|Bonney Lake city;01|07380|Bothell city;01|07695|Bremerton city;01|07835|Brewster city;01|07870|Bridgeport city;01|07940|Brier city;01|08570|Buckley city;01|08605|Bucoda town;01|08850|Burien city;01|08920|Burlington city;01|09480|Camas city;01|09970|Carbonado town;01|10215|Carnation city;01|10495|Cashmere city;01|10565|Castle Rock city;01|10635|Cathlamet town;01|11160|Centralia city;01|11475|Chehalis city;01|11615|Chelan city;01|11825|Cheney city;"
    s = s & "01|12140|Chewelah city;01|12630|Clarkston city;01|12945|Cle Elum city;01|13365|Clyde Hill city;01|13785|Colfax city;01|13855|College Place city;01|13890|Colton town;01|14170|Colville city;01|14310|Conconully town;01|14380|Concrete town;01|14485|Connell city;01|14870|Cosmopolis city;01|15080|Coulee City town;01|15115|Coulee Dam town;01|15185|Coupeville town;01|15290|Covington city;01|15710|Creston town;01|16340|Cusick town;01|16690|Darrington town;01|16795|Davenport city;01|16970|Dayton city;01|17320|Deer Park city;01|17635|Des Moines city;01|18965|DuPont city;01|19035|Duvall city;01|20155|East Wenatchee city;01|20260|Eatonville town;01|20645|Edgewood city;01|20750|Edmonds city;01|21030|Electric City city;01|21240|Ellensburg city;01|21450|Elma city;01|21485|Elmer City town;01|21730|Endicott town;01|22010|Entiat city;01|22045|Enumclaw city;01|22080|Ephrata city;01|22640|Everett city;"
    s = s & "01|22745|Everson city;01|22990|Fairfield town;01|23340|Farmington town;01|23515|Federal Way city;01|23620|Ferndale city;01|23795|Fife city;01|23970|Fircrest city;01|24810|Forks city;01|25615|Friday Harbor town;01|26140|Garfield town;01|26455|George city;01|26735|Gig Harbor city;01|27365|Gold Bar city;01|27435|Goldendale city;01|27855|Grand Coulee city;01|27925|Grandview city;01|27960|Granger city;01|27995|Granite Falls city;01|29255|Hamilton town;01|29710|Harrah town;01|29745|Harrington city;01|29920|Hartline town;01|30060|Hatton town;01|32300|Hoquiam city;01|32755|Hunts Point town;01|33000|Ilwaco city;01|33175|Index town;01|33560|Ione town;01|33805|Issaquah city;01|34575|Kahlotus city;01|34645|Kalama city;01|35065|Kelso city;01|35170|Kenmore city;01|35275|Kennewick city;01|35415|Kent city;01|35485|Kettle Falls city;01|35940|Kirkland city;01|36045|Kittitas city;"
    s = s & "01|36395|Krupp (Marlin) town;01|36710|La Center city;01|36745|Lacey city;01|36780|La Conner town;01|36850|LaCrosse town;01|37270|Lake Forest Park city;01|37900|Lake Stevens city;01|38038|Lakewood city;01|38215|Lamont town;01|38355|Langley city;01|38495|Latah town;01|38845|Leavenworth city;01|39335|Liberty Lake city;01|39510|Lind town;01|40070|Long Beach city;01|40245|Longview city;01|40770|Lyman town;01|40805|Lynden city;01|40840|Lynnwood city;01|40980|Mabton city;01|41225|McCleary city;01|42275|Malden town;01|42800|Mansfield town;01|43150|Maple Valley city;01|43395|Marcus town;01|43955|Marysville city;01|44165|Mattawa city;01|44690|Medical Lake city;01|44725|Medina city;01|45005|Mercer Island city;01|45180|Mesa city;01|45285|Metaline town;01|45320|Metaline Falls town;01|45865|Mill Creek city;01|45985|Millwood city;01|46020|Milton city;01|46685|Monroe city;01|46895|Montesano city;"
    s = s & "01|47175|Morton city;01|47245|Moses Lake city;01|47315|Mossyrock city;01|47490|Mountlake Terrace city;01|47560|Mount Vernon city;01|47665|Moxee city;01|47735|Mukilteo city;01|47805|Naches town;01|47980|Napavine city;01|48540|Nespelem town;01|48645|Newcastle city;01|48820|Newport city;01|49275|Nooksack city;01|49415|Normandy Park city;01|49485|North Bend city;01|49555|North Bonneville city;01|50045|Northport town;01|50325|Oakesdale town;01|50360|Oak Harbor city;01|50430|Oakville city;01|50570|Ocean Shores city;01|50745|Odessa town;01|50920|Okanogan city;01|51300|Olympia city;01|51340|Omak city;01|51970|Oroville city;01|52005|Orting city;01|52215|Othello city;01|52495|Pacific city;01|52950|Palouse city;01|53545|Pasco city;01|53720|Pateros city;01|53930|Pe Ell town;01|55120|Pomeroy city;01|55365|Port Angeles city;01|55785|Port Orchard city;01|55855|Port Townsend city;01|55995|Poulsbo city;"
    s = s & "01|56240|Prescott city;01|56450|Prosser city;01|56625|Pullman city;01|56695|Puyallup city;01|57115|Quincy city;01|57220|Rainier city;01|57430|Raymond city;01|57465|Reardan town;01|57535|Redmond city;01|57745|Renton city;01|57850|Republic city;01|58235|Richland city;01|58410|Ridgefield city;01|58725|Ritzville city;01|58795|Riverside town;01|59145|Rockford town;01|59180|Rock Island city;01|59775|Rosalia town;01|60055|Roslyn city;01|60160|Roy city;01|60230|Royal City city;01|60510|Ruston town;01|60860|St. John town;01|61115|Sammamish city;01|62288|SeaTac city;01|63000|Seattle city;01|63210|Sedro-Woolley city;01|63280|Selah city;01|63385|Sequim city;01|63735|Shelton city;01|63960|Shoreline city;01|64855|Skykomish town;01|65170|Snohomish city;01|65205|Snoqualmie city;01|65345|Soap Lake city;01|65625|South Bend city;01|65765|South Cle Elum town;01|66045|South Prairie town;"
    s = s & "01|66290|Spangle city;01|67000|Spokane city;01|67167|Spokane Valley city;01|67175|Sprague city;01|67210|Springdale town;01|67455|Stanwood city;01|67490|Starbuck town;01|67770|Steilacoom town;01|67875|Stevenson city;01|68260|Sultan city;01|68330|Sumas city;01|68435|Sumner city;01|68750|Sunnyside city;01|70000|Tacoma city;01|70560|Tekoa city;01|70630|Tenino city;01|71400|Tieton city;01|71785|Toledo city;01|71890|Tonasket city;01|71960|Toppenish city;01|72625|Tukwila city;01|72905|Tumwater city;01|73080|Twisp town;01|73290|Union Gap city;01|73360|Uniontown town;01|73465|University Place city;01|73780|Vader city;01|74060|Vancouver city;01|75565|Waitsburg city;01|75775|Walla Walla city;01|76125|Wapato city;01|76160|Warden city;01|76405|Washougal city;"
    NamesWA1 = s
End Function

Private Function NamesWA2() As String
    Dim s As String
    s = s & "01|76440|Washtucna town;01|76510|Waterville town;01|76720|Waverly town;01|77105|Wenatchee city;01|77630|Westport city;01|77665|West Richland city;01|78330|White Salmon city;01|78680|Wilbur town;01|78925|Wilkeson town;01|79135|Wilson Creek town;01|79275|Winlock city;01|79380|Winthrop town;01|79590|Woodinville city;01|79625|Woodland city;01|79835|Woodway city;01|79975|Yacolt town;01|80010|Yakima city;01|80150|Yarrow Point town;01|80220|Yelm city;01|80500|Zillah city;63|L0100|ADAMS COUNTY (location code 0100);63|L0101|HATTON (location code 0101);63|L0102|LIND (location code 0102);63|L0103|OTHELLO (location code 0103);63|L0104|RITZVILLE (location code 0104);63|L0105|WASHTUCNA (location code 0105);63|L0200|ASOTIN COUNTY (location code 0200);63|L0201|ASOTIN CITY (location code 0201);63|L0202|CLARKSTON (location code 0202);63|L0300|BENTON COUNTY (location code 0300);"
    s = s & "63|L0301|BENTON CITY (location code 0301);63|L0302|KENNEWICK (location code 0302);63|L0303|PROSSER (location code 0303);63|L0304|RICHLAND (location code 0304);63|L0305|WEST RICHLAND (location code 0305);63|L0333|BENTON-PTBA (location code 0333);63|L0400|CHELAN COUNTY (location code 0400);63|L0401|CASHMERE (location code 0401);63|L0402|CHELAN CITY (location code 0402);63|L0403|ENTIAT (location code 0403);63|L0404|LEAVENWORTH (location code 0404);63|L0405|WENATCHEE (location code 0405);63|L0500|CLALLAM COUNTY (location code 0500);63|L0501|FORKS (location code 0501);63|L0502|PORT ANGELES (location code 0502);63|L0503|SEQUIM (location code 0503);63|L0507|JAMESTOWN TRIBE - CLALLAM COUNTY (location code 0507);63|L0600|CLARK COUNTY (location code 0600);63|L0601|BATTLE GROUND (location code 0601);63|L0602|CAMAS (location code 0602);63|L0603|LA CENTER (location code 0603);"
    s = s & "63|L0604|RIDGEFIELD (location code 0604);63|L0605|VANCOUVER (location code 0605);63|L0606|WASHOUGAL (location code 0606);63|L0607|YACOLT (location code 0607);63|L0609|COWLITZ TRIBE - CLARK COUNTY (location code 0609);63|L0611|COWLITZ TRIBE - LA CENTER (location code 0611);63|L0666|CLARK-PTBA (location code 0666);63|L0700|COLUMBIA COUNTY (location code 0700);63|L0701|DAYTON (location code 0701);63|L0702|STARBUCK (location code 0702);63|L0800|COWLITZ COUNTY (location code 0800);63|L0801|CASTLE ROCK (location code 0801);63|L0802|KALAMA (location code 0802);63|L0803|KELSO (location code 0803);63|L0804|LONGVIEW (location code 0804);63|L0805|WOODLAND (location code 0805);63|L0900|DOUGLAS COUNTY (location code 0900);63|L0901|BRIDGEPORT (location code 0901);63|L0902|EAST WENATCHEE (location code 0902);63|L0903|MANSFIELD (location code 0903);63|L0904|ROCK ISLAND (location code 0904);"
    s = s & "63|L0905|WATERVILLE (location code 0905);63|L0909|DOUGLAS-PTBA (location code 0909);63|L1000|FERRY COUNTY (location code 1000);63|L1001|REPUBLIC (location code 1001);63|L1100|FRANKLIN COUNTY (location code 1100);63|L1101|CONNELL (location code 1101);63|L1102|KAHLOTUS (location code 1102);63|L1103|MESA (location code 1103);63|L1104|PASCO (location code 1104);63|L1111|FRANKLIN-PTBA (location code 1111);63|L1200|GARFIELD COUNTY (location code 1200);63|L1201|POMEROY (location code 1201);63|L1300|GRANT COUNTY (location code 1300);63|L1301|COULEE CITY (location code 1301);63|L1302|ELECTRIC CITY (location code 1302);63|L1303|EPHRATA (location code 1303);63|L1304|GEORGE (location code 1304);63|L1305|GRAND COULEE (location code 1305);63|L1306|HARTLINE (location code 1306);63|L1307|KRUPP (location code 1307);63|L1308|MATTAWA (location code 1308);63|L1309|MOSES LAKE (location code 1309);"
    s = s & "63|L1310|QUINCY (location code 1310);63|L1311|ROYAL CITY (location code 1311);63|L1312|SOAP LAKE (location code 1312);63|L1313|WARDEN (location code 1313);63|L1315|WILSON CREEK (location code 1315);63|L1329|COLVILLE TRIBES - MOSES LAKE (location code 1329);63|L1400|GRAYS HARBOR COUNTY (location code 1400);63|L1401|ABERDEEN (location code 1401);63|L1402|COSMOPOLIS (location code 1402);63|L1403|ELMA (location code 1403);63|L1404|HOQUIAM (location code 1404);63|L1405|MCCLEARY (location code 1405);63|L1406|MONTESANO (location code 1406);63|L1407|OAKVILLE (location code 1407);63|L1408|WESTPORT (location code 1408);63|L1409|OCEAN SHORES (location code 1409);63|L1411|CHEHALIS TRIBES - GRAYS HARBOR COUNTY (location code 1411);63|L1413|CHEHALIS TRIBES - OAKVILLE (location code 1413);63|L1500|ISLAND COUNTY (location code 1500);63|L1501|COUPEVILLE (location code 1501);"
    s = s & "63|L1502|LANGLEY (location code 1502);63|L1503|OAK HARBOR (location code 1503);63|L1600|JEFFERSON COUNTY (location code 1600);63|L1601|PORT TOWNSEND (location code 1601);63|L1603|SUQUAMISH TRIBE - JEFFERSON COUNTY (location code 1603);63|L1605|JAMESTOWN TRIBE - JEFFERSON COUNTY (location code 1605);63|L1700|KING COUNTY RTA (location code 1700);63|L1701|ALGONA (location code 1701);63|L1702|AUBURN/KING RTA (location code 1702);63|L1703|BEAUX ARTS VILLAGE (location code 1703);63|L1704|BELLEVUE RTA (location code 1704);63|L1705|BLACK DIAMOND (location code 1705);63|L1706|BOTHELL/KING (location code 1706);63|L1707|CARNATION (location code 1707);63|L1708|CLYDE HILL (location code 1708);63|L1709|DES MOINES (location code 1709);63|L1710|DUVALL (location code 1710);63|L1711|ENUMCLAW (location code 1711);63|L1712|COVINGTON (location code 1712);63|L1713|HUNTS POINT (location code 1713);"
    s = s & "63|L1714|ISSAQUAH RTA (location code 1714);63|L1715|KENT RTA (location code 1715);63|L1716|KIRKLAND (location code 1716);63|L1717|LAKE FOREST PARK (location code 1717);63|L1718|MEDINA (location code 1718);63|L1719|MERCER ISLAND (location code 1719);63|L1720|MAPLE VALLEY (location code 1720);63|L1721|NORMANDY PARK (location code 1721);63|L1722|NORTH BEND (location code 1722);63|L1723|PACIFIC/KING (location code 1723);63|L1724|REDMOND RTA (location code 1724);63|L1725|RENTON RTA (location code 1725);63|L1726|SEATTLE (location code 1726);63|L1727|SKYKOMISH (location code 1727);63|L1728|SNOQUALMIE (location code 1728);63|L1729|TUKWILA (location code 1729);63|L1730|YARROW POINT (location code 1730);63|L1731|MILTON/KING (location code 1731);63|L1732|FEDERAL WAY (location code 1732);63|L1733|SEATAC (location code 1733);63|L1734|BURIEN (location code 1734);"
    s = s & "63|L1735|WOODINVILLE RTA (location code 1735);63|L1736|NEWCASTLE (location code 1736);63|L1737|SHORELINE (location code 1737);63|L1738|KENMORE (location code 1738);63|L1739|SAMMAMISH RTA (location code 1739);63|L1741|PUYALLUP TRIBE - FEDERAL WAY (location code 1741);63|L1800|KITSAP COUNTY (location code 1800);63|L1801|BREMERTON (location code 1801);63|L1802|PORT ORCHARD (location code 1802);63|L1803|POULSBO (location code 1803);63|L1804|BAINBRIDGE ISLAND (location code 1804);63|L1805|SUQUAMISH TRIBE - KITSAP COUNTY (location code 1805);63|L1807|SUQUAMISH TRIBE - BREMERTON (location code 1807);63|L1900|KITTITAS COUNTY (location code 1900);63|L1901|CLE ELUM (location code 1901);63|L1902|ELLENSBURG (location code 1902);63|L1903|KITTITAS CITY (location code 1903);63|L1904|ROSLYN (location code 1904);63|L1905|SOUTH CLE ELUM (location code 1905);63|L2000|KLICKITAT COUNTY (location code 2000);"
    s = s & "63|L2001|BINGEN (location code 2001);63|L2002|GOLDENDALE (location code 2002);63|L2003|WHITE SALMON (location code 2003);63|L2100|LEWIS COUNTY (location code 2100);63|L2101|CENTRALIA (location code 2101);63|L2102|CHEHALIS (location code 2102);63|L2103|MORTON (location code 2103);63|L2104|MOSSYROCK (location code 2104);63|L2105|NAPAVINE (location code 2105);63|L2106|PE ELL (location code 2106);63|L2107|TOLEDO (location code 2107);63|L2108|VADER (location code 2108);63|L2109|WINLOCK (location code 2109);63|L2111|CHEHALIS TRIBES - LEWIS COUNTY (location code 2111);63|L2200|LINCOLN COUNTY (location code 2200);63|L2201|ALMIRA (location code 2201);63|L2202|CRESTON (location code 2202);63|L2203|DAVENPORT (location code 2203);63|L2204|HARRINGTON (location code 2204);63|L2205|ODESSA (location code 2205);63|L2206|REARDAN (location code 2206);63|L2207|SPRAGUE (location code 2207);"
    s = s & "63|L2208|WILBUR (location code 2208);63|L2300|MASON COUNTY (location code 2300);63|L2301|SHELTON (location code 2301);63|L2303|SQUAXIN ISLAND TRIBE - MASON COUNTY (location code 2303);63|L2400|OKANOGAN COUNTY (location code 2400);63|L2401|BREWSTER (location code 2401);63|L2402|CONCONULLY (location code 2402);63|L2403|COULEE DAM (location code 2403);63|L2404|ELMER CITY (location code 2404);63|L2405|NESPELEM (location code 2405);63|L2406|OKANOGAN CITY (location code 2406);63|L2407|OMAK (location code 2407);63|L2408|OROVILLE (location code 2408);63|L2409|PATEROS (location code 2409);63|L2410|RIVERSIDE (location code 2410);63|L2411|TONASKET (location code 2411);63|L2412|TWISP (location code 2412);63|L2413|WINTHROP (location code 2413);63|L2424|OKANOGAN-PTBA (location code 2424);63|L2425|COLVILLE TRIBES - OKANOGAN COUNTY (location code 2425);"
    s = s & "63|L2427|COLVILLE TRIBES - OKANOGAN-PTBA (location code 2427);63|L2500|PACIFIC COUNTY (location code 2500);63|L2501|ILWACO (location code 2501);63|L2502|LONG BEACH (location code 2502);63|L2503|RAYMOND (location code 2503);63|L2504|SOUTH BEND (location code 2504);63|L2600|PEND OREILLE COUNTY (location code 2600);63|L2601|CUSICK (location code 2601);63|L2602|IONE (location code 2602);63|L2603|METALINE (location code 2603);63|L2604|METALINE FALLS (location code 2604);63|L2605|NEWPORT (location code 2605);63|L2607|KALISPEL TRIBE - PEND OREILLE COUNTY (location code 2607);63|L2700|PIERCE COUNTY RTA (location code 2700);63|L2701|BONNEY LAKE RTA (location code 2701);63|L2702|BUCKLEY (location code 2702);63|L2703|CARBONADO (location code 2703);63|L2704|DUPONT (location code 2704);63|L2705|EATONVILLE (location code 2705);63|L2706|FIFE (location code 2706);63|L2707|FIRCREST (location code 2707);"
    s = s & "63|L2708|GIG HARBOR (location code 2708);63|L2709|MILTON/PIERCE (location code 2709);63|L2710|ORTING (location code 2710);63|L2711|PUYALLUP (location code 2711);63|L2712|ROY (location code 2712);63|L2713|RUSTON (location code 2713);63|L2714|SOUTH PRAIRIE (location code 2714);63|L2715|STEILACOOM (location code 2715);63|L2716|SUMNER (location code 2716);63|L2717|TACOMA (location code 2717);63|L2718|WILKESON (location code 2718);63|L2719|UNIVERSITY PLACE (location code 2719);63|L2720|EDGEWOOD (location code 2720);63|L2721|LAKEWOOD (location code 2721);63|L2723|PACIFIC/PIERCE (location code 2723);63|L2724|AUBURN/PIERCE (location code 2724);63|L2725|NISQUALLY TRIBE - PIERCE COUNTY RTA (location code 2725);63|L2727|PIERCE-PTBA RTA (location code 2727);63|L2731|NISQUALLY TRIBE - LAKEWOOD (location code 2731);63|L2733|PUYALLUP TRIBE - PUYALLUP (location code 2733);"
    s = s & "63|L2735|PUYALLUP TRIBE - FIFE (location code 2735);63|L2737|PUYALLUP TRIBE - MILTON (location code 2737);63|L2739|PUYALLUP TRIBE - EDGEWOOD (location code 2739);63|L2741|PUYALLUP TRIBE - TACOMA (location code 2741);63|L2745|PUYALLUP TRIBE - PIERCE COUNTY PTBA RTA (location code 2745);63|L2787|PIERCE PTBA - HBZ (location code 2787);63|L2788|GIG HARBOR HBZ (location code 2788);63|L2789|PIERCE NON-RTA HBZ (location code 2789);63|L2800|SAN JUAN COUNTY (location code 2800);63|L2801|FRIDAY HARBOR (location code 2801);63|L2803|LUMMI NATION - SAN JUAN COUNTY (location code 2803);63|L2900|SKAGIT COUNTY (location code 2900);63|L2901|ANACORTES (location code 2901);63|L2902|BURLINGTON (location code 2902);63|L2903|CONCRETE (location code 2903);63|L2904|HAMILTON (location code 2904);63|L2905|LA CONNER (location code 2905);63|L2906|LYMAN (location code 2906);"
    s = s & "63|L2907|MOUNT VERNON (location code 2907);63|L2908|SEDRO WOOLLEY (location code 2908);63|L2909|SWINOMISH TRIBE - SKAGIT COUNTY (location code 2909);63|L2911|SWINOMISH TRIBE - LA CONNER (location code 2911);63|L2929|SKAGIT-PTBA (location code 2929);63|L2931|SWINOMISH TRIBE - SKAGIT - PTBA (location code 2931);63|L3000|SKAMANIA COUNTY (location code 3000);63|L3001|NORTH BONNEVILLE (location code 3001);63|L3002|STEVENSON (location code 3002);63|L3100|SNOHOMISH COUNTY RTA (location code 3100);63|L3101|ARLINGTON (location code 3101);63|L3102|BRIER (location code 3102);63|L3103|DARRINGTON (location code 3103);63|L3104|EDMONDS (location code 3104);63|L3105|EVERETT RTA (location code 3105);63|L3106|GOLD BAR (location code 3106);63|L3107|GRANITE FALLS (location code 3107);63|L3108|INDEX (location code 3108);63|L3109|LAKE STEVENS (location code 3109);63|L3110|LYNNWOOD (location code 3110);"
    s = s & "63|L3111|MARYSVILLE (location code 3111);63|L3112|MONROE (location code 3112);63|L3113|MOUNTLAKE TERRACE (location code 3113);63|L3114|MUKILTEO (location code 3114);"
    NamesWA2 = s
End Function

Private Function NamesWA3() As String
    Dim s As String
    s = s & "63|L3115|SNOHOMISH CITY (location code 3115);63|L3116|STANWOOD (location code 3116);63|L3117|SULTAN (location code 3117);63|L3118|WOODWAY (location code 3118);63|L3119|MILL CREEK (location code 3119);63|L3120|BOTHELL/SNOHOMISH (location code 3120);63|L3121|TULALIP TRIBES - MARYSVILLE (location code 3121);63|L3123|STILLAGUAMISH TRIBE - ARLINGTON (location code 3123);63|L3125|STILLAGUAMISH TRIBE - MARYSVILLE (location code 3125);63|L3131|SNOHOMISH-PTBA RTA (location code 3131);63|L3200|SPOKANE COUNTY (location code 3200);63|L3201|AIRWAY HEIGHTS (location code 3201);63|L3202|CHENEY (location code 3202);63|L3203|DEER PARK (location code 3203);63|L3204|FAIRFIELD (location code 3204);63|L3205|LATAH (location code 3205);63|L3206|MEDICAL LAKE (location code 3206);63|L3207|MILLWOOD (location code 3207);63|L3208|ROCKFORD (location code 3208);63|L3209|SPANGLE (location code 3209);"
    s = s & "63|L3210|SPOKANE CITY (location code 3210);63|L3211|WAVERLY (location code 3211);63|L3212|LIBERTY LAKE (location code 3212);63|L3213|SPOKANE VALLEY (location code 3213);63|L3215|KALISPEL TRIBE - AIRWAY HEIGHTS (location code 3215);63|L3232|SPOKANE-PTBA (location code 3232);63|L3300|STEVENS COUNTY (location code 3300);63|L3301|CHEWELAH (location code 3301);63|L3302|COLVILLE (location code 3302);63|L3303|KETTLE FALLS (location code 3303);63|L3304|MARCUS (location code 3304);63|L3305|NORTHPORT (location code 3305);63|L3306|SPRINGDALE (location code 3306);63|L3400|THURSTON COUNTY (location code 3400);63|L3401|BUCODA (location code 3401);63|L3402|LACEY (location code 3402);63|L3403|OLYMPIA (location code 3403);63|L3404|RAINIER (location code 3404);63|L3405|TENINO (location code 3405);63|L3406|TUMWATER (location code 3406);63|L3407|YELM (location code 3407);"
    s = s & "63|L3409|CHEHALIS TRIBES - THURSTON COUNTY (location code 3409);63|L3411|CHEHALIS TRIBES - TUMWATER (location code 3411);63|L3413|SQUAXIN ISLAND TRIBE - THURSTON COUNTY (location code 3413);63|L3415|NISQUALLY TRIBE - THURSTON COUNTY (location code 3415);63|L3417|NISQUALLY TRIBE - THURSTON COUNTY PTBA (location code 3417);63|L3419|NISQUALLY TRIBE - LACEY (location code 3419);63|L3421|SQUAXIN ISLAND TRIBE - TUMWATER (location code 3421);63|L3434|THURSTON-PTBA (location code 3434);63|L3500|WAHKIAKUM COUNTY (location code 3500);63|L3501|CATHLAMET (location code 3501);63|L3600|WALLA WALLA COUNTY (location code 3600);63|L3601|COLLEGE PLACE (location code 3601);63|L3602|PRESCOTT (location code 3602);63|L3603|WAITSBURG (location code 3603);63|L3604|WALLA WALLA CITY (location code 3604);63|L3636|WALLA WALLA-PTBA (location code 3636);63|L3700|WHATCOM COUNTY (location code 3700);"
    s = s & "63|L3701|BELLINGHAM (location code 3701);63|L3702|BLAINE (location code 3702);63|L3703|EVERSON (location code 3703);63|L3704|FERNDALE (location code 3704);63|L3705|LYNDEN (location code 3705);63|L3706|NOOKSACK (location code 3706);63|L3707|SUMAS (location code 3707);63|L3709|LUMMI NATION - FERNDALE (location code 3709);63|L3711|LUMMI NATION - WHATCOM COUNTY (location code 3711);63|L3713|LUMMI NATION - BELLINGHAM (location code 3713);63|L3737|WHATCOM-PTBA (location code 3737);63|L3739|LUMMI NATION - WHATCOM-PTBA (location code 3739);63|L3800|WHITMAN COUNTY (location code 3800);63|L3801|ALBION (location code 3801);63|L3802|COLFAX (location code 3802);63|L3803|COLTON (location code 3803);63|L3804|ENDICOTT (location code 3804);63|L3805|FARMINGTON (location code 3805);63|L3806|GARFIELD (location code 3806);63|L3807|LA CROSSE (location code 3807);63|L3808|LAMONT (location code 3808);"
    s = s & "63|L3809|MALDEN (location code 3809);63|L3810|OAKESDALE (location code 3810);63|L3811|PALOUSE (location code 3811);63|L3812|PULLMAN (location code 3812);63|L3813|ROSALIA (location code 3813);63|L3814|ST. JOHN (location code 3814);63|L3815|TEKOA (location code 3815);63|L3816|UNIONTOWN (location code 3816);63|L3900|YAKIMA COUNTY (location code 3900);63|L3901|GRANDVIEW (location code 3901);63|L3902|GRANGER (location code 3902);63|L3903|HARRAH (location code 3903);63|L3904|MABTON (location code 3904);63|L3905|MOXEE CITY (location code 3905);63|L3906|NACHES (location code 3906);63|L3907|SELAH (location code 3907);63|L3908|SUNNYSIDE (location code 3908);63|L3909|TIETON (location code 3909);63|L3910|TOPPENISH (location code 3910);63|L3911|UNION GAP (location code 3911);63|L3912|WAPATO (location code 3912);63|L3913|YAKIMA CITY (location code 3913);63|L3914|ZILLAH (location code 3914);"
    s = s & "63|L4000|KING COUNTY NON-RTA (location code 4000);63|L4002|AUBURN/KING NON-RTA (location code 4002);63|L4004|BELLEVUE NON-RTA (location code 4004);63|L4014|ISSAQUAH NON-RTA (location code 4014);63|L4015|KENT NON-RTA (location code 4015);63|L4024|REDMOND NON-RTA (location code 4024);63|L4025|RENTON NON-RTA (location code 4025);63|L4035|WOODINVILLE NON-RTA (location code 4035);63|L4036|NEWCASTLE NON-RTA (location code 4036);63|L4039|SAMMAMISH NON-RTA (location code 4039);63|L4100|PIERCE COUNTY NON-RTA (location code 4100);63|L4101|BONNEY LAKE NON-RTA (location code 4101);63|L4103|NISQUALLY TRIBE - PIERCE COUNTY NON-RTA (location code 4103);63|L4105|PUYALLUP TRIBE - PIERCE COUNTY NON RTA (location code 4105);63|L4127|PIERCE-PTBA NON-RTA (location code 4127);63|L4200|SNOHOMISH COUNTY NON-RTA (location code 4200);63|L4201|STILLAGUAMISH TRIBE - SNOHOMISH COUNTY NON-RTA (location code 4201);"
    s = s & "63|L4205|EVERETT NON-RTA (location code 4205);63|L4231|SNOHOMISH-PTBA NON-RTA (location code 4231);63|L4233|TULALIP TRIBES - SNOHOMISH-PTBA NON-RTA (location code 4233);63|L4237|STILLAGUAMISH TRIBE - SNOHOMISH-PTBA NON-RTA (location code 4237);"
    NamesWA3 = s
End Function

Private Function NamesWA() As String
    NamesWA = NamesWA1() & NamesWA2() & NamesWA3()
End Function

' ---- end of place names ----
