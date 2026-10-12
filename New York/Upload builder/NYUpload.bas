' NY upload builder: turns New York PBT schedule exports from Crystal into one Web File
' upload file. Pick one export, or the A and B exports of a schedule together.
'
' The same steps as by hand in Notepad. On an export with * between fields, each comma
' becomes a space, since Web File reads any comma as a column break; 999999999 becomes nine
' spaces when it fills a whole field; and each * becomes a comma. A file with no * between
' fields keeps its commas, and only its 999999999s become spaces. If 999999999 is still in a
' line after that, inside a longer value, it says so and CHECK marks the line. Every line
' starts on its own line, blank lines are dropped, nothing comes after the last line, files
' join in name order (A before B), and the files you pick are never changed.
'
' Setting it up, once:
'   1. Open this file in Notepad. Ctrl A, Ctrl C.
'   2. Open a new blank workbook in Excel. Alt F11, then Insert > Module. If the new window
'      already has a line in it, Ctrl A and Delete first. Paste, then close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm). Name it NY Upload Builder.
'
' Each time:
'   1. Export the schedule from Crystal with * as the separator, the way you do now.
'   2. In NY Upload Builder, Alt F8, pick BuildNYUpload, Run.
'   3. Pick the export. For an A and a B, Ctrl click both, then Open.
'   4. It saves "<name> upload.txt" next to the first file and fills the CHECK sheet with
'      every line and what changed. Read CHECK, then upload the .txt in Web File.
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
Private Const BOX_TITLE As String = "NY upload"
Private Const CHECK_HEADS As String = "File|Line|What changed|Upload line|Before"
Private Const CHECK_COLS As Long = 5
Private Const CHECK_TOP As Long = 8     ' header row of the table on CHECK

Private mQuiet As Boolean       ' test runs only: no boxes
Private mSaid As String         ' test runs only: what the boxes would have said


Public Sub BuildNYUpload()
    ' Asks for the export files, then builds the upload file.
    Dim picked As Variant
    picked = Application.GetOpenFilename("Exports (*.csv;*.txt),*.csv;*.txt,All files (*.*),*.*", _
        1, "Pick the export, or the A and B exports together", , True)
    If Not IsArray(picked) Then Exit Sub
    RunBuild picked
End Sub

Public Function NyUploadTest(ByVal paths As String) As String
    ' For the test at home: builds from these files, joined by |, with no boxes, and
    ' returns what the boxes would have said.
    mQuiet = True
    mSaid = ""
    RunBuild Split(paths, "|")
    mQuiet = False
    NyUploadTest = mSaid
End Function

Private Sub RunBuild(ByVal picked As Variant)
    On Error GoTo Oops
    Build picked
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Something went wrong: " & Err.Description & vbLf & vbLf & _
        "The files you picked were not changed. If the upload file is open in another " & _
        "program, close it and run this again."
End Sub

Private Sub Say(ByVal msg As String)
    If mQuiet Then
        mSaid = mSaid & msg & vbLf
    Else
        MsgBox msg, vbOKOnly, BOX_TITLE
    End If
End Sub

Private Sub Build(ByVal picked As Variant)
    Dim files() As String, n As Long, i As Long, j As Long
    Dim lines() As Variant, parts As Variant, t As String, cap As Long
    Dim outA() As String, outN As Long, chk() As String
    Dim isRaw As Boolean, started As Boolean, sep As String, want As Long
    Dim upLine As String, what As String, gone As Long, blanked As Long
    Dim commas As Long, dummies As Long, odd As Long, fileN As Long, left9 As Long
    Dim perFile As String, kept As String, outPath As String, summary As String

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
            Say "Can't find " & files(i) & ". Nothing was saved."
            Exit Sub
        End If
        t = ReadText(files(i))
        t = Replace(t, vbCrLf, vbLf)
        t = Replace(t, vbCr, vbLf)
        lines(i) = Split(t, vbLf)
        cap = cap + UBound(lines(i)) + 1
    Next i
    If cap = 0 Then
        Say "The files you picked are empty. Nothing was saved."
        Exit Sub
    End If

    ReDim outA(1 To cap)
    ReDim chk(1 To cap, 1 To CHECK_COLS)
    For i = 1 To n
        parts = lines(i)
        started = False
        fileN = 0
        For j = 0 To UBound(parts)
            t = parts(j)
            If Len(Trim$(t)) > 0 Then
                If Not started Then
                    ' The file's first line says what kind of file it is.
                    isRaw = (CountOf(t, IN_SEP) >= MIN_SEPS)
                    sep = IIf(isRaw, IN_SEP, OUT_SEP)
                    want = CountOf(t, sep)
                    started = True
                End If
                upLine = CleanLine(t, isRaw, gone, blanked)
                commas = commas + gone
                dummies = dummies + blanked
                what = ""
                If CountOf(t, sep) <> want Then
                    what = "different number of fields than this file's first line"
                    odd = odd + 1
                End If
                If gone > 0 Then what = AddTo(what, gone & IIf(gone = 1, " comma to a space", " commas to spaces"))
                If blanked > 0 Then what = AddTo(what, "dummy FEIN to spaces")
                If HasDummy(upLine) Then
                    ' Notepad's Find and Replace would hit this one; it's left alone and shown.
                    what = AddTo(what, "dummy FEIN still here, inside a longer value")
                    left9 = left9 + 1
                End If
                outN = outN + 1
                fileN = fileN + 1
                outA(outN) = upLine
                chk(outN, 1) = FileOf(files(i))
                chk(outN, 2) = CStr(j + 1)
                chk(outN, 3) = what
                chk(outN, 4) = upLine
                If upLine <> t Then chk(outN, 5) = t
            End If
        Next j
        perFile = AddTo(perFile, FileOf(files(i)) & ": " & fileN & IIf(fileN = 1, " line", " lines"))
        If started And Not isRaw Then kept = AddTo(kept, FileOf(files(i)))
    Next i
    If outN = 0 Then
        Say "No lines in the files you picked. Nothing was saved."
        Exit Sub
    End If

    ReDim Preserve outA(1 To outN)
    outPath = OutPathFor(files)
    WriteText outPath, Join(outA, vbCrLf)   ' no line break after the last line, as by hand

    summary = "Saved " & FileOf(outPath) & ", " & outN & IIf(outN = 1, " line", " lines") & _
        ", in " & Left$(outPath, InStrRev(outPath, "\")) & vbLf & perFile & "." & vbLf & _
        "Commas turned to spaces: " & commas & ". Dummy FEINs turned to spaces: " & dummies & "."
    If kept <> "" Then summary = summary & vbLf & _
        "No * between fields, so its commas were kept as separators: " & kept & "."
    If odd > 0 Then summary = summary & vbLf & odd & IIf(odd = 1, " line has", " lines have") & _
        " a different number of fields than its file's first line. Web File will likely " & _
        "refuse " & IIf(odd = 1, "it", "them") & ". CHECK marks which."
    If left9 > 0 Then summary = summary & vbLf & left9 & IIf(left9 = 1, " line still has", _
        " lines still have") & " a dummy FEIN inside a longer value, left alone. CHECK marks " & _
        IIf(left9 = 1, "it", "them") & "; look before you upload."
    WriteCheck summary, chk, outN
    Say summary & vbLf & vbLf & "CHECK lists every line and what changed."
End Sub

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

Private Function HasDummy(ByVal s As String) As Boolean
    ' True when a dummy FEIN is anywhere in the line.
    Dim d As Variant
    For Each d In Split(DUMMY_FEINS, "|")
        If Len(d) > 0 Then
            If InStr(1, s, d, vbBinaryCompare) > 0 Then HasDummy = True
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

Private Sub WriteCheck(ByVal summary As String, chk() As String, ByVal rows As Long)
    ' Fills CHECK: the summary on top, then every upload line and what changed.
    Dim ws As Worksheet, info() As String, heads() As String, grid() As Variant
    Dim k As Long, r As Long, c As Long
    Application.ScreenUpdating = False
    Set ws = FreshSheet(ThisWorkbook, SH_CHECK)
    ws.Cells.NumberFormat = "@"         ' text, so 065 stays 065 and dates stay as written
    ws.Columns(2).NumberFormat = "General"  ' line numbers as numbers, so no green corners
    info = Split(summary, vbLf)
    For k = 0 To UBound(info)
        ws.Cells(1 + k, 1).Value = info(k)
    Next k
    heads = Split(CHECK_HEADS, "|")
    For k = 0 To UBound(heads)
        ws.Cells(CHECK_TOP, 1 + k).Value = heads(k)
    Next k
    ws.Rows(CHECK_TOP).Font.Bold = True
    ReDim grid(1 To rows, 1 To CHECK_COLS)
    For r = 1 To rows
        For c = 1 To CHECK_COLS
            grid(r, c) = chk(r, c)
        Next c
        grid(r, 2) = CLng(chk(r, 2))
    Next r
    ws.Cells(CHECK_TOP + 1, 1).Resize(rows, CHECK_COLS).Value = grid
    ws.Range(ws.Cells(CHECK_TOP, 1), ws.Cells(CHECK_TOP + rows, 3)).Columns.AutoFit
    ws.Columns(4).ColumnWidth = 90
    ws.Columns(5).ColumnWidth = 90
    Application.ScreenUpdating = True
    If Not mQuiet Then ws.Activate
End Sub

Private Function FreshSheet(wb As Workbook, ByVal nm As String) As Worksheet
    ' The named sheet, emptied, or a new one at the front.
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            ws.Cells.Clear
            Set FreshSheet = ws
            Exit Function
        End If
    Next ws
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = nm
    Set FreshSheet = ws
End Function
