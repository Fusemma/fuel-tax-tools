' NY upload builder: turns New York PBT schedule exports from Crystal into one Web File
' upload file. Pick one export, or the A and B exports of a schedule together.
'
' The same steps as by hand in Notepad. On an export with * between fields, each comma
' becomes a space, since Web File reads any comma as a column break; 999999999 becomes nine
' spaces when it fills a whole field; and each * becomes a comma. A file with no * between
' fields keeps its commas, and only its 999999999s become spaces. If 999999999 is still in a
' line after that, inside a longer value, it says so and CHECK marks the line. Line breaks
' are cleaned up, blank lines are dropped, nothing comes after the last line, files
' join in name order (A before B), and the files you pick are never changed.
'
' Setting it up, once:
'   1. Download this file (the download button at the top right of its page).
'   2. Open a new blank workbook in Excel. Alt F11, Insert, Module, then Ctrl A and Ctrl V to
'      paste all of this in. Close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm). Name it NY Upload Builder.
'
' Each time:
'   1. Export the schedule from Crystal with * as the separator.
'   2. In NY Upload Builder, Alt F8, pick BuildNYUpload, Run.
'   3. Pick the export. For an A and a B, Ctrl click both, then Open.
'   4. It saves "<name> upload.txt" next to the first file. Nothing pops up; the CHECK sheet
'      opens with one line. Green: good to go, upload the .txt in Web File. Red: the lines
'      to look at are right under it, or it says why nothing was saved; fix that and run it
'      again. Every line, and what changed, is folded away under the + on the left.
'
' The settings below are the only lines to change when a rule changes.

Option Explicit

' ---- settings -------------------------------------------------------------------------

' Dummy FEINs, between bars. A field that is exactly one of these becomes the same number
' of spaces, as by hand: 999999999 becomes nine spaces.
Private Const DUMMY_FEINS As String = "|999999999|"

' What each comma inside the data becomes. One space, as by hand.
Private Const COMMA_TO As String = " "

' What Crystal puts between fields, and what Web File wants there.
Private Const IN_SEP As String = "*"
Private Const OUT_SEP As String = ","

' A file is a * export when its first line has at least this many *. The smallest
' New York schedule has five columns, so four.
Private Const MIN_SEPS As Long = 4

' Added to the name of the upload file. A .txt opens in Notepad, not Excel, and Web File
' takes .txt the same as .csv.
Private Const OUT_SUFFIX As String = " upload"
Private Const OUT_EXT As String = ".txt"

' ---- the rest -------------------------------------------------------------------------

Private Const SH_CHECK As String = "CHECK"
Private Const CHECK_HEADS As String = "File|Line|What changed|Upload line|Before"
Private Const LOOK_HEADS As String = "File|Line|What's wrong|Upload line|Before"
Private Const CHECK_COLS As Long = 5
Private Const WORDS_WIDTH As Long = 60  ' widest the words column gets; longer words wrap
Private Const FOLD_HEAD As String = "Every line, and what changed (click + on the left to open)"
Private Const SAID_FIELDS As String = "Has an extra or missing field, so Web File will likely " & _
    "refuse it. Often a stray "     ' then * or comma, and " in a name."
Private Const SAID_DUMMY As String = " is inside a longer number here, so it was left alone. " & _
    "Check the FEIN."
Private Const SAID_CLOSE As String = " is open in another program, close it and run this again."


Public Sub BuildNYUpload()
    ' Asks for the export files, then builds the upload file.
    Dim picked As Variant
    picked = Application.GetOpenFilename("Exports (*.csv;*.txt),*.csv;*.txt,All files (*.*),*.*", _
        1, "Pick the export, or the A and B exports together", , True)
    If Not IsArray(picked) Then Exit Sub
    RunBuild picked
End Sub

Public Function NyUploadTest(ByVal paths As String) As String
    ' Builds from these files, joined by |, without the file box, and returns CHECK's top line.
    RunBuild Split(paths, "|")
    On Error Resume Next
    NyUploadTest = CStr(ThisWorkbook.Worksheets(SH_CHECK).Range("A1").Value)
End Function

Private Sub RunBuild(ByVal picked As Variant)
    Dim why As String
    On Error GoTo Oops
    Build picked
    Exit Sub
Oops:
    why = Trim$(Err.Description)
    If Right$(why, 1) = "." Then why = Left$(why, Len(why) - 1)
    Resume Tell
Tell:
    On Error GoTo 0                     ' if CHECK can't be written either, stop here, don't loop
    Application.ScreenUpdating = True
    NothingSaved "Something went wrong: " & why & "."
End Sub

Private Sub NothingSaved(ByVal why As String)
    ' CHECK with only a red top line, saying why nothing was saved.
    Dim ws As Worksheet
    Set ws = FreshSheet(ThisWorkbook, SH_CHECK)
    StatusLine ws, "Nothing was saved. " & why, False
    ShowCheck ws
End Sub

Private Sub Build(ByVal picked As Variant)
    Dim files() As String, n As Long, i As Long, j As Long
    Dim lines() As Variant, parts As Variant, t As String, cap As Long
    Dim outA() As String, outN As Long, chk() As String, looks() As String, nLook As Long
    Dim isRaw As Boolean, started As Boolean, sep As String, want As Long
    Dim upLine As String, what As String, wrong As String, d As String, gone As Long, blanked As Long
    Dim commas As Long, dummies As Long, fileN As Long
    Dim perFile As String, kept As String, outPath As String, status As String, notes As String

    ' The picked files, A before B.
    n = UBound(picked) - LBound(picked) + 1
    ReDim files(1 To n)
    For i = 1 To n
        files(i) = CStr(picked(LBound(picked) + i - 1))
    Next i
    SortByName files

    ' Read them all first, so a missing file stops the run before anything is saved.
    ReDim lines(1 To n)
    For i = 1 To n
        If Len(Dir(files(i))) = 0 Then
            NothingSaved "Can't find " & FileOf(files(i)) & "."
            Exit Sub
        End If
        If Not TryRead(files(i), t) Then
            NothingSaved "Can't open " & FileOf(files(i)) & ". If it" & SAID_CLOSE
            Exit Sub
        End If
        t = Replace(t, vbCrLf, vbLf)
        t = Replace(t, vbCr, vbLf)
        lines(i) = Split(t, vbLf)
        cap = cap + UBound(lines(i)) + 1
    Next i
    If cap = 0 Then
        NothingSaved "The files you picked are empty."
        Exit Sub
    End If

    ReDim outA(1 To cap)
    ReDim chk(1 To cap, 1 To CHECK_COLS)
    ReDim looks(1 To cap, 1 To CHECK_COLS)
    For i = 1 To n
        parts = lines(i)
        started = False
        fileN = 0
        For j = 0 To UBound(parts)
            t = parts(j)
            If Len(Trim$(t)) > 0 Then
                If Not started Then
                    ' The file's first line says what kind of file it is; most of its
                    ' lines say how many fields a line should have.
                    isRaw = (CountOf(t, IN_SEP) >= MIN_SEPS)
                    sep = IIf(isRaw, IN_SEP, OUT_SEP)
                    want = UsualCount(parts, sep)
                    started = True
                End If
                upLine = CleanLine(t, isRaw, gone, blanked)
                commas = commas + gone
                dummies = dummies + blanked
                what = ""
                wrong = ""
                If CountOf(t, sep) <> want Then
                    what = "different number of fields than most of this file's lines"
                    wrong = SAID_FIELDS & IIf(isRaw, "*", "comma") & " in a name."
                End If
                If gone > 0 Then what = AddTo(what, gone & IIf(gone = 1, " comma to a space", " commas to spaces"))
                If blanked > 0 Then what = AddTo(what, "dummy FEIN to spaces")
                d = DummyIn(upLine)
                If Len(d) > 0 Then
                    ' Notepad's Find and Replace would hit this one; it's left alone and shown.
                    what = AddTo(what, "dummy FEIN still here, inside a longer value")
                    wrong = Trim$(wrong & " " & d & SAID_DUMMY)
                End If
                outN = outN + 1
                fileN = fileN + 1
                outA(outN) = upLine
                chk(outN, 1) = FileOf(files(i))
                chk(outN, 2) = CStr(j + 1)
                chk(outN, 3) = what
                chk(outN, 4) = upLine
                If upLine <> t Then chk(outN, 5) = t
                If Len(wrong) > 0 Then
                    nLook = nLook + 1
                    looks(nLook, 1) = chk(outN, 1)
                    looks(nLook, 2) = chk(outN, 2)
                    looks(nLook, 3) = wrong
                    looks(nLook, 4) = upLine
                    looks(nLook, 5) = t
                End If
            End If
        Next j
        perFile = AddTo(perFile, FileOf(files(i)) & ": " & fileN & IIf(fileN = 1, " line", " lines"))
        If started And Not isRaw Then kept = AddTo(kept, FileOf(files(i)))
    Next i
    If outN = 0 Then
        NothingSaved "No lines in the files you picked."
        Exit Sub
    End If

    ReDim Preserve outA(1 To outN)
    outPath = OutPathFor(files)

    If nLook = 0 Then
        status = "Done. " & FileOf(outPath) & " is saved next to your exports. Nothing needs a look."
    Else
        status = nLook & IIf(nLook = 1, " line needs", " lines need") & " a look before you upload. " & _
            FileOf(outPath) & " is saved next to your exports."
    End If
    notes = "Saved " & FileOf(outPath) & ", " & outN & IIf(outN = 1, " line", " lines") & _
        ", in " & Left$(outPath, InStrRev(outPath, "\")) & vbLf & perFile & "." & vbLf & _
        "Commas turned to spaces: " & commas & ". Dummy FEINs turned to spaces: " & dummies & "."
    If kept <> "" Then notes = notes & vbLf & _
        "No * between fields, so its commas were kept as separators: " & kept & "."

    ' CHECK first, then the file: if saving fails, CHECK turns red and says nothing was saved.
    WriteCheck status, notes, looks, nLook, chk, outN
    If Not TryWrite(outPath, Join(outA, vbCrLf)) Then  ' no line break after the last line, as by hand
        NothingSaved "Can't save " & FileOf(outPath) & ". If it" & SAID_CLOSE
    End If
End Sub

Private Function UsualCount(ByVal parts As Variant, ByVal sep As String) As Long
    ' How many separators most of the file's lines have; on a tie, the line that comes first.
    Dim seen As Object, j As Long, c As Long, best As Long, k As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    For j = 0 To UBound(parts)
        If Len(Trim$(parts(j))) > 0 Then
            c = CountOf(parts(j), sep)
            seen(c) = seen(c) + 1
        End If
    Next j
    best = -1
    For Each k In seen.Keys             ' in the order first seen
        If best < 0 Then
            best = k
        ElseIf seen(k) > seen(best) Then
            best = k
        End If
    Next k
    UsualCount = best
End Function

Private Function TryRead(ByVal path As String, ByRef text As String) As Boolean
    ' Reads the file; False when it can't, like when another program has it open.
    On Error GoTo Nope
    text = ReadText(path)
    TryRead = True
Nope:
End Function

Private Function TryWrite(ByVal path As String, ByVal body As String) As Boolean
    ' Saves the file; False when it can't, like when another program has it open.
    On Error GoTo Nope
    WriteText path, body
    TryWrite = True
Nope:
End Function

Private Function CleanLine(ByVal s As String, ByVal isRaw As Boolean, ByRef gone As Long, _
        ByRef blanked As Long) As String
    ' One line. A * export gets a space for each comma and a comma for each *. Either kind
    ' gets its dummy FEINs turned to spaces.
    Dim t As String, f() As String, k As Long, d As String
    If isRaw Then
        gone = CountOf(s, ",")
        t = Replace(s, ",", COMMA_TO)
        f = Split(t, IN_SEP)
    Else
        gone = 0
        f = Split(s, OUT_SEP)
    End If
    blanked = 0
    For k = 0 To UBound(f)
        d = Trim$(f(k))
        If IsDummy(d) Then
            f(k) = Replace(f(k), d, Space$(Len(d)))
            blanked = blanked + 1
        End If
    Next k
    CleanLine = Join(f, OUT_SEP)
End Function

Private Function IsDummy(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) > 0 Then IsDummy = (InStr(1, DUMMY_FEINS, "|" & s & "|", vbBinaryCompare) > 0)
End Function

Private Function DummyIn(ByVal s As String) As String
    ' The dummy FEIN found anywhere in the line, or "" when there's none.
    Dim d As Variant
    For Each d In Split(DUMMY_FEINS, "|")
        If Len(d) > 0 Then
            If InStr(1, s, d, vbBinaryCompare) > 0 Then
                DummyIn = d
                Exit Function
            End If
        End If
    Next d
End Function

Private Function CountOf(ByVal s As String, ByVal sep As String) As Long
    CountOf = (Len(s) - Len(Replace(s, sep, ""))) \ Len(sep)
End Function

Private Function AddTo(ByVal s As String, ByVal more As String) As String
    ' Adds to a list, with "; " between.
    If s = "" Then AddTo = more Else AddTo = s & "; " & more
End Function

Private Function FileOf(ByVal path As String) As String
    FileOf = Mid$(path, InStrRev(path, "\") + 1)
End Function

Private Function BaseOf(ByVal path As String) As String
    ' The file name without its .csv or .txt.
    Dim nm As String, p As Long
    nm = FileOf(path)
    p = InStrRev(nm, ".")
    If p > 1 Then nm = Left$(nm, p - 1)
    BaseOf = nm
End Function

Private Sub SortByName(files() As String)
    ' Plain name order, ignoring case, so an A export comes before its B.
    Dim i As Long, j As Long, t As String
    For i = LBound(files) + 1 To UBound(files)
        t = files(i)
        j = i - 1
        Do While j >= LBound(files)
            If StrComp(FileOf(files(j)), FileOf(t), vbTextCompare) <= 0 Then Exit Do
            files(j + 1) = files(j)
            j = j - 1
        Loop
        files(j + 1) = t
    Next i
End Sub

Private Function OutPathFor(files() As String) As String
    ' Next to the first file, named for the part of the names the files share.
    Dim nm As String, other As String, p As String, i As Long, k As Long
    nm = BaseOf(files(1))
    For i = 2 To UBound(files)
        other = BaseOf(files(i))
        k = 0
        Do While k < Len(nm) And k < Len(other)
            If StrComp(Mid$(nm, k + 1, 1), Mid$(other, k + 1, 1), vbTextCompare) <> 0 Then Exit Do
            k = k + 1
        Loop
        nm = Left$(nm, k)
    Next i
    nm = TrimEnd(nm, " -_.")
    If Len(nm) < 3 Then nm = BaseOf(files(1))
    p = Left$(files(1), InStrRev(files(1), "\")) & NoSpaceByDash(nm & OUT_SUFFIX) & OUT_EXT
    For i = 1 To UBound(files)
        If StrComp(files(i), p, vbTextCompare) = 0 Then p = Left$(p, Len(p) - Len(OUT_EXT)) & " 2" & OUT_EXT
    Next i
    OutPathFor = p
End Function

Private Function NoSpaceByDash(ByVal s As String) As String
    ' Web File refuses a file name with a space next to a dash.
    Do While InStr(s, " -") > 0
        s = Replace(s, " -", "-")
    Loop
    Do While InStr(s, "- ") > 0
        s = Replace(s, "- ", "-")
    Loop
    NoSpaceByDash = s
End Function

Private Function TrimEnd(ByVal s As String, ByVal chars As String) As String
    Do While Len(s) > 0
        If InStr(1, chars, Right$(s, 1), vbBinaryCompare) = 0 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    TrimEnd = s
End Function

Private Function ReadText(ByVal path As String) As String
    ' The whole file as text, whether it was saved as UTF-8, UTF-16 or plain Windows text.
    Dim st As Object, b() As Byte, cs As String, t As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                         ' bytes
    st.Open
    st.LoadFromFile path
    If st.Size = 0 Then
        st.Close
        Exit Function
    End If
    b = st.Read
    If Starts(b, &HFF, &HFE) Then
        cs = "unicode"
    ElseIf Starts(b, &HFE, &HFF) Then
        cs = "unicodeFFFE"
    ElseIf Starts(b, &HEF, &HBB, &HBF) Then
        cs = "utf-8"
    ElseIf IsUtf8(b) Then
        cs = "utf-8"
    Else
        cs = "windows-1252"
    End If
    st.Position = 0
    st.Type = 2                         ' text
    st.Charset = cs
    t = st.ReadText
    st.Close
    If Len(t) > 0 Then
        If Left$(t, 1) = ChrW(65279) Then t = Mid$(t, 2)
    End If
    ReadText = t
End Function

Private Function Starts(b() As Byte, ByVal x0 As Long, ByVal x1 As Long, _
        Optional ByVal x2 As Long = -1) As Boolean
    ' True when the file begins with these bytes.
    Dim n As Long
    n = UBound(b) - LBound(b) + 1
    If x2 < 0 Then
        If n >= 2 Then Starts = (b(0) = x0 And b(1) = x1)
    Else
        If n >= 3 Then Starts = (b(0) = x0 And b(1) = x1 And b(2) = x2)
    End If
End Function

Private Function IsUtf8(b() As Byte) As Boolean
    ' True when the bytes hold accented UTF-8 letters and nothing that breaks UTF-8.
    Dim i As Long, j As Long, k As Long, last As Long, multi As Boolean
    last = UBound(b)
    i = LBound(b)
    Do While i <= last
        If b(i) < &H80 Then
            k = 0
        ElseIf (b(i) And &HE0) = &HC0 Then
            k = 1
        ElseIf (b(i) And &HF0) = &HE0 Then
            k = 2
        ElseIf (b(i) And &HF8) = &HF0 Then
            k = 3
        Else
            Exit Function
        End If
        If i + k > last Then Exit Function
        For j = 1 To k
            If (b(i + j) And &HC0) <> &H80 Then Exit Function
        Next j
        If k > 0 Then multi = True
        i = i + k + 1
    Loop
    IsUtf8 = multi
End Function

Private Sub WriteText(ByVal path As String, ByVal body As String)
    ' Saves as UTF-8 with no marker bytes in front, the same as plain text for these files.
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                         ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText body
    st.Position = 3                     ' past the marker bytes
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1                        ' bytes
    bin.Open
    st.CopyTo bin
    bin.SaveToFile path, 2              ' 2: replace the file if it is there
    bin.Close
    st.Close
End Sub

Private Sub WriteCheck(ByVal status As String, ByVal notes As String, looks() As String, _
        ByVal nLook As Long, chk() As String, ByVal rows As Long)
    ' Fills CHECK: the top line, then the lines to look at if there are any, then every
    ' line and what changed, folded away under a +.
    Dim ws As Worksheet, info() As String, k As Long, r As Long, fold As Long, top As Long
    Application.ScreenUpdating = False
    Set ws = FreshSheet(ThisWorkbook, SH_CHECK)
    ws.Cells.NumberFormat = "@"         ' text, so 065 stays 065 and dates stay as written
    ws.Columns(2).NumberFormat = "General"  ' line numbers as numbers, so no green corners
    ws.Cells.VerticalAlignment = xlTop  ' a wrapped line reads level with its row
    StatusLine ws, status, (nLook = 0)
    r = 3
    If nLook > 0 Then
        PutTable ws, r, LOOK_HEADS, looks, nLook
        FitColumns ws, r, r + nLook
        r = r + nLook + 2
    End If
    fold = r
    ws.Cells(fold, 1).Value = FOLD_HEAD
    info = Split(notes, vbLf)
    For k = 0 To UBound(info)
        ws.Cells(fold + 1 + k, 1).Value = info(k)
    Next k
    top = fold + UBound(info) + 3       ' one blank row after the notes
    PutTable ws, top, CHECK_HEADS, chk, rows
    FitColumns ws, top, top + rows
    ws.Columns(4).ColumnWidth = 90
    ws.Columns(5).ColumnWidth = 90
    ws.Columns(3).WrapText = True
    ws.Range(ws.Rows(2), ws.Rows(top + rows)).AutoFit
    ws.Outline.SummaryRow = xlAbove     ' the + sits on the fold's heading row
    ws.Range(ws.Rows(fold + 1), ws.Rows(top + rows)).Group
    ws.Outline.ShowLevels RowLevels:=1
    ShowCheck ws
End Sub

Private Sub PutTable(ws As Worksheet, ByVal r As Long, ByVal headText As String, src() As String, _
        ByVal n As Long)
    ' Headings on row r, then n rows from src: file, line, words, upload line, before.
    Dim heads() As String, grid() As Variant, i As Long, c As Long
    heads = Split(headText, "|")
    For c = 0 To UBound(heads)
        ws.Cells(r, 1 + c).Value = heads(c)
    Next c
    ws.Range(ws.Cells(r, 1), ws.Cells(r, CHECK_COLS)).Font.Bold = True
    ReDim grid(1 To n, 1 To CHECK_COLS)
    For i = 1 To n
        For c = 1 To CHECK_COLS
            grid(i, c) = src(i, c)
        Next c
        grid(i, 2) = CLng(src(i, 2))
    Next i
    ws.Cells(r + 1, 1).Resize(n, CHECK_COLS).Value = grid
End Sub

Private Sub FitColumns(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    ' Fits File, Line and the words column to rows r1 to r2, never narrower than they are.
    Dim c As Long, was As Double
    For c = 1 To 3
        was = ws.Columns(c).ColumnWidth
        ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Columns.AutoFit
        If ws.Columns(c).ColumnWidth < was Then ws.Columns(c).ColumnWidth = was
    Next c
    If ws.Columns(3).ColumnWidth > WORDS_WIDTH Then ws.Columns(3).ColumnWidth = WORDS_WIDTH
End Sub

Private Sub StatusLine(ws As Worksheet, ByVal msg As String, ByVal good As Boolean)
    ' Row 1: green when it's good to go, red when it isn't.
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, CHECK_COLS))
        .Interior.Color = IIf(good, RGB(198, 239, 206), RGB(255, 199, 206))
        .Font.Color = IIf(good, RGB(0, 97, 0), RGB(156, 0, 6))
        .Font.Bold = True
    End With
    ws.Cells(1, 1).Value = msg
End Sub

Private Sub ShowCheck(ws As Worksheet)
    ' CHECK to the front, scrolled to the top.
    Application.ScreenUpdating = True
    On Error Resume Next                ' a hidden workbook can't show it; the line is still there
    ws.Visible = xlSheetVisible
    ws.Parent.Activate
    ws.Activate
    Application.Goto ws.Range("A1"), True
End Sub

Private Function FreshSheet(wb As Workbook, ByVal nm As String) As Worksheet
    ' The named sheet, with nothing left from the last run, or a new one at the front.
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            On Error Resume Next        ' no fold to take out is fine
            ws.Cells.ClearOutline
            On Error GoTo 0
            ws.Cells.Clear
            ws.Rows.Hidden = False
            ws.Rows.RowHeight = ws.StandardHeight
            ws.Columns.ColumnWidth = ws.StandardWidth
            Set FreshSheet = ws
            Exit Function
        End If
    Next ws
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = nm
    Set FreshSheet = ws
End Function
