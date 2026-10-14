' Recap fill: fills the EFT recap in Word from the Fields sheet.
'
' In the Word file, type a tag like [[amount]] wherever a value goes, in the text, a table,
' a header or a footer. On Fields, list each tag (amount, no brackets) with its value: a
' number, text, or a formula pointing at another sheet. Numbers and dates go into Word the
' way the value cell shows them, so give that cell the format the recap wants.
'
' Fill opens the Word file named on Fields, from this workbook's folder, swaps every tag for
' its value and saves a filled copy beside it with today's date, like Recap 2026-10-11.docx.
' The Word file itself never changes. Then it says what it filled and any tag it couldn't
' match.
'
' Setting it up from this text, once:
'   1. Copy all of this text (click Raw, then Ctrl A, Ctrl C).
'   2. Open a new blank workbook in Excel. Alt F11, then Insert > Module. If the new window
'      already has a line in it, Ctrl A and Delete first. Paste, then close that window.
'   3. File > Save As > Excel Macro-Enabled Workbook (.xlsm), in the folder with the Word file.
'   4. Alt F8, pick Fill, Run. The first run makes the Fields sheet and its Fill button.

Option Explicit

Private Const SHEET_NAME As String = "Fields"
Private Const DEFAULT_FILE As String = "Recap.docx"
Private Const BOX_TITLE As String = "Recap fill"

' Word's numbers for the settings used here
Private Const WD_FIND_STOP As Long = 0
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_NO_SAVE As Long = 0
Private Const WD_FORMAT_DOC As Long = 0     ' the old .doc
Private Const WD_FORMAT_DOCX As Long = 16

Public Sub Fill()
    Dim ws As Worksheet, wd As Object, doc As Object, parts As Collection, part As Object
    Dim x As Variant, lastRow As Long, r As Long, n As Long, i As Long
    Dim tags() As String, vals() As String, tag As String, found As Long
    Dim wordName As String, folder As String, src As String, copyPath As String
    Dim filled As String, missing As String, twice As String, bad As String, leftover As String
    Dim msg As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MakeFieldsSheet
        MsgBox "Made the Fields sheet. Put your tags and values on it, then press Fill.", vbInformation, BOX_TITLE
        Exit Sub
    End If
    On Error GoTo Failed

    ' The tags and values, under the Tag heading
    x = Application.Match("Tag", ws.Columns(1), 0)
    If IsError(x) Then
        MsgBox "Fields needs the heading Tag in column A, with the tags under it.", vbExclamation, BOX_TITLE
        Exit Sub
    End If
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim tags(1 To lastRow), vals(1 To lastRow)
    For r = x + 1 To lastRow
        tag = ""
        If Not IsError(ws.Cells(r, 1).Value) Then tag = Trim$(Replace(Replace(ws.Cells(r, 1).Value, "[[", ""), "]]", ""))
        If tag = "" Then
            ' a blank row
        ElseIf IsError(ws.Cells(r, 2).Value) Then
            bad = bad & "[[" & tag & "]] shows " & ws.Cells(r, 2).Text & vbLf
        ElseIf InStr(1, "|" & Join(tags, "|") & "|", "|" & tag & "|", vbTextCompare) > 0 Then
            twice = twice & "[[" & tag & "]]" & vbLf
        Else
            n = n + 1
            tags(n) = tag
            vals(n) = ValueText(ws.Cells(r, 2))
        End If
    Next
    If bad <> "" Then
        MsgBox "Fix these on Fields first. Nothing was filled." & vbLf & vbLf & bad, vbExclamation, BOX_TITLE
        Exit Sub
    End If
    If n = 0 Then
        MsgBox "No tags on Fields yet. List them under Tag, each with its value.", vbExclamation, BOX_TITLE
        Exit Sub
    End If

    ' The Word file, from this workbook's folder
    x = Application.Match("Word file", ws.Columns(1), 0)
    If Not IsError(x) Then wordName = Trim$(ws.Cells(x, 2).Text)
    If wordName = "" Then wordName = DEFAULT_FILE
    If ThisWorkbook.Path = "" Then
        MsgBox "Save this workbook first, in the folder with the Word file.", vbExclamation, BOX_TITLE
        Exit Sub
    End If
    folder = LocalFolder(ThisWorkbook.Path, Array(Environ$("OneDriveCommercial"), Environ$("OneDriveConsumer"), Environ$("OneDrive")))
    If folder = "" Then
        MsgBox "Can't find this workbook's folder on the PC:" & vbLf & ThisWorkbook.Path & vbLf & vbLf & _
            "Save the workbook and the Word file in a folder on the PC, then try again.", vbExclamation, BOX_TITLE
        Exit Sub
    End If
    src = WordFile(folder, wordName)
    If src = "" Then
        MsgBox "Can't find " & wordName & " in" & vbLf & folder & vbLf & vbLf & _
            "Put the Word file there, or fix its name on Fields next to Word file.", vbExclamation, BOX_TITLE
        Exit Sub
    End If
    copyPath = CopyName(src)

    Application.Cursor = xlWait
    ' A Word of its own that stays out of sight, so nothing already open in Word is touched
    Set wd = CreateObject("Word.Application")
    ' A new document made from the file, so the file itself never changes. It gets a window
    ' inside that hidden Word, since Word won't let a macro into text boxes in a header or
    ' footer without one.
    Set doc = wd.Documents.Add(src, False, 0, True)
    doc.TrackRevisions = False
    Set parts = TextParts(doc)
    For i = 1 To n
        found = 0
        For Each part In parts
            found = found + FillIn(part, "[[" & tags(i) & "]]", vals(i))
        Next
        If found = 0 Then
            missing = missing & "[[" & tags(i) & "]]" & vbLf
        Else
            filled = filled & "[[" & tags(i) & "]]  " & IIf(vals(i) = "", "(blank)", Shorten(vals(i))) & _
                IIf(found > 1, "  (" & found & " places)", "") & vbLf
        End If
    Next
    For Each part In parts
        AddTagsIn part.Text, leftover
    Next
    If filled <> "" Then
        doc.SaveAs2 FileName:=copyPath, AddToRecentFiles:=False, _
            FileFormat:=IIf(LCase$(Right$(copyPath, 4)) = ".doc", WD_FORMAT_DOC, WD_FORMAT_DOCX)
    End If
    doc.Close WD_NO_SAVE
    Set doc = Nothing
    wd.Quit WD_NO_SAVE
    Set wd = Nothing
    Application.Cursor = xlDefault

    If filled = "" Then
        msg = "Nothing filled, so no copy was saved." & vbLf
    Else
        msg = "Saved " & FileOf(copyPath) & " beside " & FileOf(src) & "." & vbLf
    End If
    If missing <> "" Then msg = msg & vbLf & "Not in the Word file:" & vbLf & missing
    If twice <> "" Then msg = msg & vbLf & "On Fields twice, the first one was used:" & vbLf & twice
    If leftover <> "" Then msg = msg & vbLf & "In the Word file but not on Fields:" & vbLf & leftover
    If filled <> "" Then msg = msg & vbLf & "Filled:" & vbLf & filled
    If Len(msg) > 1000 Then msg = Left$(msg, 990) & "..."
    MsgBox msg, IIf(missing & twice & leftover = "", vbInformation, vbExclamation), BOX_TITLE
    Exit Sub

Failed:
    msg = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close WD_NO_SAVE
    If Not wd Is Nothing Then wd.Quit WD_NO_SAVE
    Application.Cursor = xlDefault
    MsgBox "Fill stopped: " & msg, vbExclamation, BOX_TITLE
End Sub

' The value as it should read in Word. Numbers and dates keep the cell's format.
Private Function ValueText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value2
    Select Case VarType(v)
        Case vbEmpty
            ValueText = ""
        Case vbString
            ValueText = v
        Case vbBoolean
            ValueText = UCase$(CStr(v))
        Case Else
            On Error Resume Next
            ValueText = c.Text
            ValueText = Application.WorksheetFunction.Text(v, c.NumberFormat)
            On Error GoTo 0
            ValueText = Trim$(ValueText)
    End Select
    ' Alt Enter in a cell becomes a line break in Word
    ValueText = Replace(Replace(ValueText, vbCrLf, vbLf), vbLf, Chr$(11))
End Function

' Every part of the document a tag can sit in: the text and its tables, headers, footers,
' footnotes and text boxes.
Private Function TextParts(ByVal doc As Object) As Collection
    Dim parts As New Collection, story As Object, shp As Object, x As Long
    x = doc.Sections(1).Headers(1).Range.StoryType      ' Word can skip the headers without this
    For Each story In doc.StoryRanges
        Do Until story Is Nothing
            parts.Add story
            Set story = story.NextStoryRange
        Loop
    Next
    ' Text boxes in headers and footers aren't on that list. Word keeps them all here.
    For Each shp In doc.Sections(1).Headers(1).Shapes
        AddBox parts, shp
    Next
    Set TextParts = parts
End Function

' Adds a shape's text, if it's a shape that holds text.
Private Sub AddBox(ByVal parts As Collection, ByVal shp As Object)
    Dim box As Object
    On Error Resume Next
    Set box = shp.TextFrame.TextRange
    If Not box Is Nothing Then parts.Add box
End Sub

' Swaps every tag in one part for the value. Returns how many it swapped.
Private Function FillIn(ByVal part As Object, ByVal tag As String, ByVal valText As String) As Long
    Dim r As Object
    Set r = part.Duplicate
    With r.Find
        .ClearFormatting
        .Text = tag
        .Forward = True
        .Wrap = WD_FIND_STOP
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Do While r.Find.Execute
        r.Text = valText
        r.Collapse WD_COLLAPSE_END
        FillIn = FillIn + 1
    Loop
End Function

' Adds each [[tag]] in the text to the list, once.
Private Sub AddTagsIn(ByVal s As String, ByRef list As String)
    Dim i As Long, j As Long, t As String
    j = InStr(s, "]]")
    Do While j > 0
        i = InStrRev(s, "[[", j)
        If i > 0 Then
            t = Mid$(s, i, j - i + 2)
            If Len(t) <= 50 And InStr(t, vbCr) = 0 And InStr(3, t, "]]") = Len(t) - 1 Then
                If InStr(1, vbLf & list, vbLf & t & vbLf, vbTextCompare) = 0 Then list = list & t & vbLf
            End If
        End If
        j = InStr(j + 2, s, "]]")
    Loop
End Sub

' A value cut short for the report.
Private Function Shorten(ByVal s As String) As String
    s = Replace(s, Chr$(11), " ")
    If Len(s) > 40 Then s = Left$(s, 37) & "..."
    Shorten = s
End Function

' The Word file's full path, or "" if it isn't there. A name with no extension tries .docx,
' then .doc. A full path works too.
Private Function WordFile(ByVal folder As String, ByVal wordName As String) As String
    Dim p As String, tries As Variant, t As Variant, hit As String
    If wordName Like "?:\*" Or Left$(wordName, 2) = "\\" Then p = wordName Else p = folder & "\" & wordName
    If InStr(FileOf(p), ".") > 0 Then tries = Array(p) Else tries = Array(p & ".docx", p & ".doc")
    On Error Resume Next
    For Each t In tries
        hit = ""
        hit = Dir$(t)
        If Len(hit) > 0 Then
            WordFile = t
            Exit Function
        End If
    Next
End Function

' Recap 2026-10-11.docx beside the Word file, or Recap 2026-10-11 (2).docx if that's taken.
' A .doc stays a .doc; anything else is saved as .docx.
Private Function CopyName(ByVal src As String) As String
    Dim stem As String, ext As String, p As String, k As Long
    stem = Left$(src, InStrRev(src, ".") - 1) & " " & Format$(Date, "yyyy-mm-dd")
    If LCase$(Mid$(src, InStrRev(src, "."))) = ".doc" Then ext = ".doc" Else ext = ".docx"
    p = stem & ext
    k = 1
    Do While Len(Dir$(p)) > 0
        k = k + 1
        p = stem & " (" & k & ")" & ext
    Loop
    CopyName = p
End Function

Private Function FileOf(ByVal p As String) As String
    FileOf = Mid$(p, InStrRev(p, "\") + 1)
End Function

' On OneDrive, Excel gives the workbook's folder as a web address. This finds the same
' folder on the PC under one of the OneDrive folders in roots. Anything else comes back
' as it is, and "" means it couldn't be found.
Private Function LocalFolder(ByVal p As String, ByVal roots As Variant) As String
    Dim bits() As String, root As Variant, tail As String, hit As String, i As Long
    If LCase$(Left$(p, 4)) <> "http" Then
        LocalFolder = p
        Exit Function
    End If
    bits = Split(UrlDecode(p), "/")         ' https:, blank, the site, then the folders
    On Error Resume Next
    For Each root In roots
        If root <> "" Then
            tail = ""
            For i = UBound(bits) To 3 Step -1   ' keeps the longest tail that's there
                tail = "\" & bits(i) & tail
                hit = ""
                hit = Dir$(root & tail, vbDirectory)
                If Len(hit) > 0 Then LocalFolder = root & tail
            Next
            If LocalFolder <> "" Then Exit Function
        End If
    Next
End Function

Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long
    i = InStr(s, "%")
    Do While i > 0 And i <= Len(s) - 2
        If Mid$(s, i + 1, 2) Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
            s = Left$(s, i - 1) & Chr$(CLng("&H" & Mid$(s, i + 1, 2))) & Mid$(s, i + 3)
        End If
        i = InStr(i + 1, s, "%")
    Loop
    UrlDecode = s
End Function

' The Fields sheet: the Word file's name, the sample's three tags, and the Fill button.
Private Sub MakeFieldsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SHEET_NAME
    ws.Range("A1:B1").Value = Array("Word file", DEFAULT_FILE)
    ws.Range("A3:B3").Value = Array("Tag", "Value")
    ws.Range("A1,A3:B3").Font.Bold = True
    ' The sample Recap.docx has these three. Swap them for your own.
    ws.Range("A4:A6").Value = Application.Transpose(Array("period", "amount", "due"))
    ws.Range("B4").NumberFormat = "mmmm yyyy"
    ws.Range("B4").Value = DateSerial(2026, 9, 1)
    ws.Range("B5").NumberFormat = "$#,##0.00"
    ws.Range("B5").Value = 1234.56
    ws.Range("B6").NumberFormat = "mmmm d, yyyy"
    ws.Range("B6").Value = DateSerial(2026, 10, 20)
    ws.Range("B:B").HorizontalAlignment = xlLeft
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 30
    With ws.Buttons.Add(ws.Range("D1").Left, ws.Range("D1").Top, 96, 30)
        .Caption = "Fill"
        .OnAction = "Fill"
    End With
    ws.Range("D4").Value = "In Word, type [[tag]] wherever a value goes."
    ws.Range("D5").Value = "Fill saves a dated copy beside the Word file."
    ws.Range("D4:D5").Font.Color = RGB(89, 89, 89)
End Sub
