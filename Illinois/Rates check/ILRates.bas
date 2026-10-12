' Illinois rates check: what changed in the state's sales tax rates since the last files loaded.
'
' The Illinois Department of Revenue posts every city's and county's sales tax rates as plain
' text files, a new set for each change date (January 1 and July 1 most years; also August 1
' in 2026). The What changed sheet shows the rates you watch (set just below these notes) that
' changed since the files loaded last time: the place, its county, old rate, new rate and when
' it starts.
'
' Each time the state posts new files:
'   1. Go to tax.illinois.gov, Research, Tax Rates, Sales Tax Rate - Machine Readable Files.
'      Under Sales Tax, download the top two links: the Effective file (ordmache-current.txt)
'      and the Addendum Address File under it (ordmacha-current.txt). The second one has the
'      rates for places with business districts. Never the Address Specific zip; it's 2 GB.
'   2. Click Check new rates file on the What changed sheet and pick both files.
'   3. Save the workbook if you want to keep it. It never saves itself.
' Log keeps each change once, the first time a file shows it, plus a line for each check with
' nothing new. Saved rates holds the files' rates for next time. Read here says all this on a
' tab of its own.
'
' First time: paste this into a module in a new workbook, save it as a macro-enabled workbook
' (.xlsm), then Alt F8 and run CheckIllinoisRates. It sets up the sheets and the button, then
' asks for the files. With nothing loaded before, it shows what changed on the files' own start
' date, from the old rates the state lists beside each rate.
'
' A change is a rate that went up or down (the low rate too), a place or business district
' that's new or gone, or a new name. A row the state only re-dated isn't a change.

Option Explicit

' ---- the rates you watch ----------------------------------------------------------
' Separate with commas. GM general merchandise, DM drugs and medical, PGM and PDM the same two
' from the file's Purchases columns, AUTO vehicles. Each has a high and a low rate. They only
' differ where a district tax doesn't reach every seller, like Metro East in Madison and
' St. Clair counties, where the high rate is for sellers who owe it.
Private Const WATCH As String = "GM"

' ---- the rest -----------------------------------------------------------------------
Private Const ALL_KINDS As String = "GM,DM,PGM,PDM,AUTO"
Private Const SH_MAIN As String = "What changed"
Private Const SH_READ As String = "Read here"
Private Const SH_LOG As String = "Log"
Private Const SH_SAVED As String = "Saved rates"
Private Const BTN_NAME As String = "CheckRatesButton"
Private Const BTN_TEXT As String = "Check new rates file"
Private Const BOX_TITLE As String = "Illinois rates"
Private Const TITLE_TEXT As String = "Illinois sales tax rates: what changed"
Private Const MAX_BYTES As Long = 50000000      ' the real files are under 10 MB

' The state's two layouts: each field's width, and the Saved rates column it goes to (0 for
' none). Every line of a file is the same length, so a short last line means a cut-short file.
Private Const MAIN_LEN As Long = 211
Private Const MAIN_WIDTHS As String = "10,25,25,1,8,5,5,5,5,1,5,5,5,5,1,5,5,5,5,1,8,8,5,5,5,5,1,5,5,5,5,1,5,5,5,5,1"
Private Const MAIN_COLS As String = "1,3,4,5,8,9,11,10,12,13,14,16,15,17,18,19,20,21,22,23,24,25,26,28,27,29,30,31,33,32,34,35,36,37,38,39,40"
Private Const ADDR_LEN As Long = 206
Private Const ADDR_WIDTHS As String = "10,3,25,25,35,8,5,5,5,5,1,5,5,5,5,1,8,8,5,5,5,5,1,5,5,5,5,1"
Private Const ADDR_COLS As String = "1,2,3,4,0,8,9,11,10,12,13,14,16,15,17,18,24,25,26,28,27,29,30,31,33,32,34,35"
Private Const GROCERY_LEN As Long = 106

' Saved rates columns. Rates are percents (6.25 means 6.25%), the vehicle fee is dollars.
Private Const HEADS As String = "Code|District|Place|County|Address file|Addresses|Line|Starts|" & _
    "General merchandise|General merchandise low|Drugs and medical|Drugs and medical low|Sales flag|" & _
    "Purchases general merchandise|Purchases general merchandise low|Purchases drugs and medical|" & _
    "Purchases drugs and medical low|Purchases flag|Vehicles|Vehicles low|Vehicle fee percent|" & _
    "Vehicle fee most|Vehicles flag|Old rate started|Old rate ended|" & _
    "Old general merchandise|Old general merchandise low|Old drugs and medical|Old drugs and medical low|" & _
    "Old sales flag|Old purchases general merchandise|Old purchases general merchandise low|" & _
    "Old purchases drugs and medical|Old purchases drugs and medical low|Old purchases flag|Old vehicles|" & _
    "Old vehicles low|Old vehicle fee percent|Old vehicle fee most|Old vehicles flag"
Private Const NC As Long = 40
Private Const C_CODE As Long = 1, C_DIST As Long = 2, C_PLACE As Long = 3, C_COUNTY As Long = 4
Private Const C_AFILE As Long = 5, C_ADDRS As Long = 6, C_LINE As Long = 7, C_START As Long = 8
Private Const C_GM As Long = 9, C_GML As Long = 10, C_DM As Long = 11, C_DML As Long = 12, C_RFLAG As Long = 13
Private Const C_PGM As Long = 14, C_PGML As Long = 15, C_PDM As Long = 16, C_PDML As Long = 17, C_PFLAG As Long = 18
Private Const C_AUTO As Long = 19, C_AUTOL As Long = 20, C_FEEP As Long = 21, C_FEEMAX As Long = 22, C_VFLAG As Long = 23
Private Const C_PSTART As Long = 24, C_PEND As Long = 25
Private Const OLD_OFF As Long = 17          ' an old rate's column is its new rate's plus this
Private Const SAVED_ROW As Long = 4         ' header row on Saved rates

' The What changed sheet, top to bottom
Private Const DUE_ROW As Long = 6           ' red when newer files should be out
Private Const ADDR_ROW As Long = 7          ' yellow when the address file wasn't picked
Private Const BAND_ROW As Long = 9          ' the big count
Private Const LIST_ROW As Long = 11         ' the list's header

Private Const LIST_HEADS As String = "Place|County|Type|What changed|Old rate|New rate|Starts|Rate|Code|Notes"
Private Const LIST_COLS As Long = 10
Private Const LOG_HEADS As String = "Checked|Files|Checked against|" & LIST_HEADS & "|Change ID"
Private Const LOG_COLS As Long = 14

' One change, as kept while sorting and drawing. 1 to 10 are the list's columns in order.
Private Const F_SORT As Long = 0, F_PLACE As Long = 1, F_COUNTY As Long = 2, F_TYPE As Long = 3
Private Const F_WHAT As Long = 4, F_OLD As Long = 5, F_NEW As Long = 6, F_START As Long = 7
Private Const F_KIND As Long = 8, F_CODE As Long = 9, F_NOTES As Long = 10, F_LOOK As Long = 11
Private Const F_ID As Long = 12, F_KEY As Long = 13

Private mFeed As Collection     ' test runs only: answers for the boxes, in order
Private mSaid As String         ' test runs only: what the boxes would have said
Private mToday As Date          ' test runs only: the day to act as today


' ---- the button, and opening the workbook ------------------------------------------

Public Sub CheckIllinoisRates()
    ' Asks for the state's new rates files, then shows what changed.
    On Error GoTo Oops
    LoadFiles
    GoTo Tidy
Oops:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Say "Something went wrong: " & Err.Description & vbLf & vbLf & _
        "Try it again. If it keeps happening, download the files again and pick those."
Tidy:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
End Sub

Public Sub Auto_Open()
    ' Opening the workbook puts up a red line once newer files should be out. That isn't
    ' anything to save, so the workbook is marked unchanged and closing won't ask.
    On Error Resume Next
    If SheetByName(SH_MAIN) Is Nothing Then Exit Sub
    DueLine
    ThisWorkbook.Saved = True
End Sub

Private Sub LoadFiles()
    Dim picks As Variant, mainPath As String, addrPath As String, problem As String
    Dim recs As Variant, n As Long, nMain As Long, nAddr As Long, change As String, addrChange As String
    Dim oldRecs As Variant, nOld As Long, oldMain As Long, savedFiles As String, savedWhen As Variant
    Dim savedChange As String, savedAddr As Boolean, changes As Collection, others As Long
    Dim today As Date, files As String, against As String, i As Long

    today = CheckDate()
    EnsureSheets
    picks = PickFiles(True, "Pick the Illinois rates file and its Addendum Address File (hold Ctrl to pick both)")
    If Not IsArray(picks) Then Exit Sub
    If Not SortPicks(picks, mainPath, addrPath, problem) Then Say problem: Exit Sub
    If mainPath = "" Then
        picks = PickFiles(False, "Now pick the rates file, the Effective one under Sales Tax (ordmache-current.txt)")
        If Not IsArray(picks) Then Exit Sub
        If Not SortPicks(picks, mainPath, addrPath, problem) Then Say problem: Exit Sub
        If mainPath = "" Then
            Say "That's the address file again, not the rates file." & vbLf & vbLf & _
                "Pick ordmache-current.txt, the Effective file under Sales Tax. Nothing was changed."
            Exit Sub
        End If
    End If
    If addrPath = "" Then
        picks = PickFiles(False, "Now pick the Addendum Address File (ordmacha-current.txt). Cancel skips the business districts this time.")
        If IsArray(picks) Then
            If Not SortPicks(picks, mainPath, addrPath, problem) Then Say problem: Exit Sub
        End If
    End If

    Application.Cursor = xlWait
    Application.StatusBar = "Reading the files..."
    If Not ReadFiles(mainPath, addrPath, recs, nMain, nAddr, problem) Then
        Application.Cursor = xlDefault
        Say problem
        Exit Sub
    End If
    n = nMain + nAddr

    ' The files are for the newest start date in them. Two files from one posting agree on it.
    change = NewestStart(recs, 1, nMain)
    If nAddr > 0 Then
        addrChange = NewestStart(recs, nMain + 1, n)
        If addrChange <> change Then
            Application.Cursor = xlDefault
            If Not Confirm("These two files may not be from the same posting." & vbLf & vbLf & _
                    "The newest start date in " & FileNameOf(mainPath) & " is " & ShowDate8(change) & ". In " & _
                    FileNameOf(addrPath) & " it's " & ShowDate8(addrChange) & "." & vbLf & vbLf & _
                    "Load them anyway? They will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
    End If

    ' The files loaded last time, if there were any
    nOld = ReadSaved(oldRecs, savedFiles, savedWhen, savedChange, savedAddr)
    If nOld > 0 Then
        If change < savedChange Then
            Application.Cursor = xlDefault
            If Not Confirm("These files look older than the ones loaded last time." & vbLf & vbLf & _
                    "They're for rates starting " & ShowDate8(change) & ". The saved rates start " & _
                    ShowDate8(savedChange) & "." & vbLf & vbLf & _
                    "Load them anyway? They will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
        For i = 1 To nOld
            If oldRecs(i, C_DIST) = "" Then oldMain = oldMain + 1
        Next
        If nMain < oldMain * 0.9 Then
            Application.Cursor = xlDefault
            If Not Confirm("This rates file has only " & Format$(nMain, "#,##0") & " places. The saved rates have " & _
                    Format$(oldMain, "#,##0") & ", so the download may have been cut short." & vbLf & vbLf & _
                    "Load it anyway? It will replace the saved rates.") Then Exit Sub
            Application.Cursor = xlWait
        End If
    End If

    Application.StatusBar = "Comparing..."
    Set changes = New Collection
    CompareAll oldRecs, nOld, savedAddr, recs, n, (nAddr > 0), change, changes, others

    files = FileNameOf(mainPath)
    If addrPath <> "" Then files = files & " and " & FileNameOf(addrPath)
    If nOld > 0 Then against = savedFiles & ", loaded " & ShowWhen(savedWhen) Else against = "Nothing, first files"

    Application.ScreenUpdating = False
    Application.StatusBar = "Drawing..."
    DrawFront changes, others, files, today, change, (nAddr > 0), (nOld > 0), savedFiles, savedWhen, _
        savedChange, savedAddr
    AddToLog changes, files, today, against
    SaveCopy recs, n, files, today, change, (nAddr > 0)
    DrawReadHere SheetByName(SH_READ)
    DueLine
    ThisWorkbook.Activate
    SheetByName(SH_MAIN).Activate
    SheetByName(SH_MAIN).Range("A1").Select
End Sub


' ---- reading the state's files ---------------------------------------------------------

Private Function SortPicks(ByVal picks As Variant, ByRef mainPath As String, ByRef addrPath As String, _
                           ByRef problem As String) As Boolean
    ' Sorts the files picked into the rates file and the address file. False, with problem
    ' saying why, when one is neither, or there are two of a kind.
    Dim i As Long, p As String, kind As String
    For i = LBound(picks) To UBound(picks)
        p = CStr(picks(i))
        kind = FileKind(p, problem)
        If kind = "" Then Exit Function
        If kind = "main" Then
            If mainPath <> "" And StrComp(mainPath, p, vbTextCompare) <> 0 Then
                problem = "That's two rates files. Pick one rates file (ordmache) and one address file " & _
                    "(ordmacha), from the same day's downloads. Nothing was changed."
                Exit Function
            End If
            mainPath = p
        Else
            If addrPath <> "" And StrComp(addrPath, p, vbTextCompare) <> 0 Then
                problem = "That's two address files. Pick one rates file (ordmache) and one address file " & _
                    "(ordmacha), from the same day's downloads. Nothing was changed."
                Exit Function
            End If
            addrPath = p
        End If
    Next
    SortPicks = True
End Function

Private Function FileKind(ByVal fullPath As String, ByRef problem As String) As String
    ' "main" for the rates file, "addr" for the Addendum Address File, or blank with problem
    ' saying why. Reads only the start of a file, so the big address file is turned away fast.
    Dim nm As String, size As Double, head As String, first As String, p As Long
    nm = FileNameOf(fullPath)
    If LCase$(nm) Like "alladdressrates*" Or LCase$(Right$(nm, 4)) = ".zip" Then
        problem = BigFileWords(nm)
        Exit Function
    End If
    On Error Resume Next
    size = FileLen(fullPath)
    If Err.Number <> 0 Then size = -1
    On Error GoTo 0
    If size < 0 Or size > MAX_BYTES Then
        problem = BigFileWords(nm)
        Exit Function
    End If
    If size > 0 Then head = ReadStart(fullPath, 1000)
    p = InStr(head, vbLf)
    If p = 0 Then p = InStr(head, vbCr)
    If p > 0 Then first = Left$(head, p - 1) Else first = head
    first = Replace(first, vbCr, "")
    If Left$(first, 10) Like "###-####-#" Then
        If Len(first) = MAIN_LEN Then FileKind = "main": Exit Function
        If Len(first) = ADDR_LEN And Mid$(first, 11, 1) = "R" Then FileKind = "addr": Exit Function
        If Len(first) = GROCERY_LEN Then
            problem = "That's the grocery rates file (" & nm & "). This checks the sales tax rates, the top " & _
                "two files under Sales Tax (ordmache-current.txt and ordmacha-current.txt). Pick those." & _
                vbLf & vbLf & "Nothing was changed."
            Exit Function
        End If
    End If
    problem = "That doesn't look like the Illinois rates file. " & nm & " isn't laid out like the state's files." & _
        vbLf & vbLf & "Download the top two files under Sales Tax on the state's Machine Readable Files page " & _
        "(ordmache-current.txt and ordmacha-current.txt) and pick those. Nothing was changed."
End Function

Private Function BigFileWords(ByVal nm As String) As String
    BigFileWords = nm & " is the big address file, or too big to be the rates file. This never reads that " & _
        "one; it's 2 GB zipped and 9 GB unzipped." & vbLf & vbLf & "Download the top two files under " & _
        "Sales Tax instead (ordmache-current.txt and ordmacha-current.txt) and pick those. Nothing was changed."
End Function

Private Function ReadFiles(ByVal mainPath As String, ByVal addrPath As String, ByRef recs As Variant, _
                           ByRef nMain As Long, ByRef nAddr As Long, ByRef problem As String) As Boolean
    ' Both files into recs, columns in HEADS order: the rates file's rows, then one row for each
    ' district in the address file. A district's first address stands for all of them; the
    ' rest are only counted.
    Dim lines As Variant, aLines As Variant, i As Long, j As Long, r As Long, k As String, cnt As Long
    Dim seen As Collection, groups As Collection, firstAt() As Long, addrs() As Long

    nMain = 0: nAddr = 0
    If Not FileLines(mainPath, MAIN_LEN, lines, problem) Then Exit Function
    If addrPath <> "" Then
        If Not FileLines(addrPath, ADDR_LEN, aLines, problem) Then Exit Function
        Set groups = New Collection
        ReDim firstAt(1 To 1000)
        ReDim addrs(1 To 1000)
        For i = 0 To UBound(aLines)
            If Len(aLines(i)) > 0 Then
                k = UCase$(Left$(aLines(i), 13))        ' the place code and the district
                j = Found(groups, k)
                If j = 0 Then
                    nAddr = nAddr + 1
                    If nAddr > UBound(firstAt) Then
                        ReDim Preserve firstAt(1 To nAddr * 2)
                        ReDim Preserve addrs(1 To nAddr * 2)
                    End If
                    groups.Add nAddr, k
                    firstAt(nAddr) = i
                    addrs(nAddr) = 1
                Else
                    addrs(j) = addrs(j) + 1
                End If
            End If
        Next
    End If

    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then nMain = nMain + 1
    Next
    ReDim recs(1 To nMain + nAddr, 1 To NC)
    Set seen = New Collection
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            r = r + 1
            ParseLine CStr(lines(i)), MAIN_WIDTHS, MAIN_COLS, recs, r
            ' A code the state lists twice gets a line number
            k = UCase$(recs(r, C_CODE))
            cnt = Found(seen, k)
            If cnt > 0 Then seen.Remove k
            seen.Add cnt + 1, k
            recs(r, C_LINE) = CStr(cnt + 1)
        End If
    Next
    For j = 1 To nAddr
        r = r + 1
        ParseLine CStr(aLines(firstAt(j))), ADDR_WIDTHS, ADDR_COLS, recs, r
        recs(r, C_ADDRS) = CStr(addrs(j))
        recs(r, C_LINE) = "1"
    Next
    ReadFiles = True
End Function

Private Function FileLines(ByVal fullPath As String, ByVal width As Long, ByRef lines As Variant, _
                           ByRef problem As String) As Boolean
    ' The file's lines, each one checked: the layout's length, starting with a place code.
    Dim body As String, i As Long, last As Long, nm As String
    nm = FileNameOf(fullPath)
    On Error Resume Next
    body = ReadWholeFile(fullPath)
    If Err.Number <> 0 Then
        problem = "Couldn't open " & nm & ": " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    If InStr(body, vbLf) = 0 Then body = Replace(body, vbCr, vbLf) Else body = Replace(body, vbCr, "")
    lines = Split(body, vbLf)
    last = -1
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then last = i
    Next
    If last < 0 Then
        problem = "That file is empty: " & nm & "." & vbLf & vbLf & "Download it again and pick that. Nothing was changed."
        Exit Function
    End If
    For i = 0 To last
        If Len(lines(i)) > 0 Then
            If Len(lines(i)) <> width Or Not (Left$(lines(i), 10) Like "###-####-#") Then
                If i = last Then
                    problem = nm & " looks cut short: its last line stops partway." & vbLf & vbLf & _
                        "Download it again and pick that. Nothing was changed."
                Else
                    problem = "Line " & Format$(i + 1, "#,##0") & " of " & nm & " isn't like the others, so the " & _
                        "file may have been changed or cut short." & vbLf & vbLf & _
                        "Download it again and pick that. Nothing was changed."
                End If
                Exit Function
            End If
        End If
    Next
    FileLines = True
End Function

Private Sub ParseLine(ByVal s As String, ByVal widthList As String, ByVal colList As String, _
                      ByRef recs As Variant, ByVal r As Long)
    ' One line of the file into row r, each field cut at its width
    Dim widths As Variant, cols As Variant, f As Long, pos As Long, w As Long, c As Long
    widths = Split(widthList, ",")
    cols = Split(colList, ",")
    pos = 1
    For f = 0 To UBound(widths)
        w = Val(widths(f))
        c = Val(cols(f))
        If c > 0 Then recs(r, c) = FieldValue(Mid$(s, pos, w), c)
        pos = pos + w
    Next
    For c = 1 To NC
        If IsEmpty(recs(r, c)) Then recs(r, c) = ""
    Next
End Sub

Private Function FieldValue(ByVal raw As String, ByVal c As Long) As String
    ' A field the way Saved rates keeps it: rates as percents (06250 is 6.25), the vehicle fee
    ' in dollars (02000 is 20.00), everything else trimmed.
    Select Case c
        Case C_GM To C_DML, C_PGM To C_PDML, C_AUTO To C_FEEP, C_GM + OLD_OFF To C_DML + OLD_OFF, _
             C_PGM + OLD_OFF To C_PDML + OLD_OFF, C_AUTO + OLD_OFF To C_FEEP + OLD_OFF
            FieldValue = NumText(Val(raw) / 1000)
        Case C_FEEMAX, C_FEEMAX + OLD_OFF
            FieldValue = NumText(Val(raw) / 100)
        Case Else
            FieldValue = Trim$(raw)
    End Select
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

Private Function ReadStart(ByVal fullPath As String, ByVal most As Long) As String
    ' The first few bytes of a file, however big it is. Blank when it won't open.
    Dim f As Integer, body As String, size As Long
    f = FreeFile
    On Error Resume Next
    Open fullPath For Binary Access Read Shared As #f
    If Err.Number <> 0 Then Exit Function
    size = LOF(f)
    If size <= 0 Or size > most Then size = most
    body = Space$(size)
    Get #f, 1, body
    Close #f
    ReadStart = body
End Function

Private Function NewestStart(ByRef recs As Variant, ByVal fromRow As Long, ByVal toRow As Long) As String
    Dim i As Long, best As String
    For i = fromRow To toRow
        If recs(i, C_START) > best Then best = recs(i, C_START)
    Next
    NewestStart = best
End Function


' ---- comparing -------------------------------------------------------------------------

Private Sub CompareAll(ByRef oldRecs As Variant, ByVal nOld As Long, ByVal oldAddr As Boolean, _
                       ByRef newRecs As Variant, ByVal nNew As Long, ByVal newAddr As Boolean, _
                       ByVal change As String, ByVal changes As Object, ByRef others As Long)
    ' Every watched rate that changed between the saved files and the new ones. With nothing
    ' saved, a line that starts on the change date is checked against its own old rates. others
    ' counts lines where only rates not watched changed.
    Dim oIdx As Collection, nIdx As Collection, oYes As Collection, nYes As Collection
    Dim oRow As Variant, nRow As Variant, oUnk As Variant, nUnk As Variant
    Dim no As Long, nn As Long, i As Long, j As Long, r As Long, o As Long

    others = 0
    Set nIdx = New Collection
    Set nYes = New Collection
    nn = BuildLines(newRecs, nNew, nIdx, nRow, nUnk, nYes)
    Set oIdx = New Collection
    Set oYes = New Collection
    If nOld > 0 Then no = BuildLines(oldRecs, nOld, oIdx, oRow, oUnk, oYes)

    For i = 1 To nn
        If Not nUnk(i) Then
            r = nRow(i)
            If nOld = 0 Then
                OwnChanges newRecs, r, change, changes, others
            Else
                j = Found(oIdx, LineKey(newRecs, r))
                If j = 0 Then
                    ' A district line the saved files couldn't have had: no address file then
                    If newRecs(r, C_DIST) <> "" And Not oldAddr Then
                        OwnChanges newRecs, r, change, changes, others
                    Else
                        NewLine newRecs, r, changes, others
                    End If
                ElseIf oUnk(j) Then
                    OwnChanges newRecs, r, change, changes, others
                Else
                    LineChanges oldRecs, CLng(oRow(j)), False, newRecs, r, changes, others
                End If
            End If
        End If
    Next

    ' Gone: lines the saved files had that these don't
    For j = 1 To no
        If Not oUnk(j) Then
            o = oRow(j)
            If Found(nIdx, LineKey(oldRecs, o)) = 0 Then
                ' With no address file this time, its district lines weren't checked, not gone
                If oldRecs(o, C_DIST) = "" Or newAddr Then GoneLine oldRecs, o, nYes, changes, others
            End If
        End If
    Next
End Sub

Private Function BuildLines(ByRef recs As Variant, ByVal n As Long, ByVal idx As Object, ByRef rowOf As Variant, _
                            ByRef unknown As Variant, ByVal yes As Object) As Long
    ' The lines to compare. A place is one line from the rates file, unless the state marks it
    ' "see the address file": then its lines come from the address file, one per district
    ' (R00 is the place outside its districts). idx gets each line's key, rowOf its record,
    ' yes the codes marked for the address file. unknown is True for a marked place with no
    ' address lines loaded, whose rates the rates file leaves at zero.
    Dim i As Long, cnt As Long, k As String, code As String, inAddr As Collection
    Set inAddr = New Collection
    On Error Resume Next
    For i = 1 To n
        code = UCase$(recs(i, C_CODE))
        If recs(i, C_DIST) = "" Then
            If UCase$(recs(i, C_AFILE)) = "Y" Then yes.Add 1, code
        Else
            inAddr.Add 1, code
        End If
    Next
    On Error GoTo 0
    ReDim rowOf(1 To n + 1)
    ReDim unknown(1 To n + 1)
    For i = 1 To n
        code = UCase$(recs(i, C_CODE))
        If recs(i, C_DIST) = "" Then
            If UCase$(recs(i, C_AFILE)) = "Y" And Found(inAddr, code) > 0 Then GoTo NextRec
        ElseIf Found(yes, code) = 0 Then
            GoTo NextRec        ' the rates file has this place's rates, so its line counts
        End If
        k = LineKey(recs, i)
        If Found(idx, k) = 0 Then
            cnt = cnt + 1
            idx.Add cnt, k
            rowOf(cnt) = i
            unknown(cnt) = (recs(i, C_DIST) = "" And UCase$(recs(i, C_AFILE)) = "Y")
        End If
NextRec:
    Next
    BuildLines = cnt
End Function

Private Function LineKey(ByRef recs As Variant, ByVal i As Long) As String
    ' A line is a place code, a district (R00 for a line from the rates file), and a number for
    ' a code the state lists more than once.
    Dim d As String
    d = UCase$(recs(i, C_DIST))
    If d = "" Then d = "R00"
    LineKey = UCase$(recs(i, C_CODE)) & "|" & d & "|" & recs(i, C_LINE)
End Function

Private Sub OwnChanges(ByRef nRecs As Variant, ByVal r As Long, ByVal change As String, ByVal changes As Object, _
                       ByRef others As Long)
    ' With nothing to check against: a line that starts on the change date changed then, from
    ' the old rates the state lists beside it. One with no old rates at all is new.
    If nRecs(r, C_START) <> change Then Exit Sub
    If AnyRate(nRecs, r, True) Then
        LineChanges nRecs, r, True, nRecs, r, changes, others
    Else
        NewLine nRecs, r, changes, others
    End If
End Sub

Private Sub LineChanges(ByRef oRecs As Variant, ByVal o As Long, ByVal oOld As Boolean, ByRef nRecs As Variant, _
                        ByVal r As Long, ByVal changes As Object, ByRef others As Long)
    ' One line's changes, a list line for each watched rate that changed. oOld: the old rates
    ' are the line's own old-rate columns.
    Dim kinds As Variant, i As Long, what As String, look As String, notes As String
    Dim named As String, base As String, shown As Boolean
    base = LineNote(nRecs, r, False)
    If Not oOld Then
        If Not SameText(oRecs(o, C_PLACE), nRecs(r, C_PLACE)) Then named = "Was named " & oRecs(o, C_PLACE) & "."
        If Not SameText(oRecs(o, C_COUNTY), nRecs(r, C_COUNTY)) Then
            named = Trim$(named & " County was " & OrBlank(oRecs(o, C_COUNTY)) & ".")
        End If
    End If
    kinds = WatchList()
    For i = 0 To UBound(kinds)
        If KindDiff(oRecs, o, oOld, nRecs, r, CStr(kinds(i)), what, look, notes) Then
            If Not oOld Then
                If oRecs(o, C_START) = nRecs(r, C_START) Then notes = Trim$(notes & " Changed without a new start date.")
            End If
            changes.Add MakeChange(what, look, JoinNotes(base, notes, named), CStr(kinds(i)), oRecs, o, oOld, nRecs, r)
            shown = True
            named = ""          ' a new name goes on the first line only
        End If
    Next
    If shown Then Exit Sub
    If named <> "" Then
        changes.Add MakeChange("Name changed", "other", JoinNotes(base, "", named), CStr(kinds(0)), oRecs, o, oOld, _
            nRecs, r)
        Exit Sub
    End If
    kinds = Split(ALL_KINDS, ",")
    For i = 0 To UBound(kinds)
        If Not InList(CStr(kinds(i)), WATCH) Then
            If KindDiff(oRecs, o, oOld, nRecs, r, CStr(kinds(i)), what, look, notes) Then
                others = others + 1
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function KindDiff(ByRef oRecs As Variant, ByVal o As Long, ByVal oOld As Boolean, ByRef nRecs As Variant, _
                          ByVal r As Long, ByVal kind As String, ByRef what As String, ByRef look As String, _
                          ByRef notes As String) As Boolean
    ' True when this kind's high or low rate differs (for vehicles, the fee too). what is the
    ' headline; notes has the low rate and the fee.
    Dim hi As Long, lo As Long, f1 As Long, f2 As Long, oh As Double, nh As Double, ol As Double, nl As Double
    Dim op As Double, np As Double, om As Double, nm As Double
    what = "": look = "": notes = ""
    If Not KindCols(kind, hi, lo, f1, f2) Then Exit Function
    oh = RateAt(oRecs, o, hi, oOld): nh = RateAt(nRecs, r, hi, False)
    ol = RateAt(oRecs, o, lo, oOld): nl = RateAt(nRecs, r, lo, False)
    If Abs(nh - oh) > 0.000001 Then
        If nh > oh Then what = "Went up": look = "up" Else what = "Went down": look = "down"
        ' The low rate only gets a note where it isn't the same as the high rate
        If Abs(ol - oh) > 0.000001 Or Abs(nl - nh) > 0.000001 Then
            If Abs(nl - ol) > 0.000001 Then
                notes = "Low rate was " & Pct(ol) & ", now " & Pct(nl) & "."
            Else
                notes = "Low rate " & Pct(nl) & "."
            End If
        End If
    ElseIf Abs(nl - ol) > 0.000001 Then
        ' Only the low rate moved: the list shows the low rates, the note the high one
        If nl > ol Then what = "Low rate went up": look = "up" Else what = "Low rate went down": look = "down"
        notes = "High rate stays " & Pct(nh) & "."
    End If
    If f1 > 0 Then
        op = RateAt(oRecs, o, f1, oOld): np = RateAt(nRecs, r, f1, False)
        om = RateAt(oRecs, o, f2, oOld): nm = RateAt(nRecs, r, f2, False)
        If Abs(np - op) > 0.000001 Or Abs(nm - om) > 0.000001 Then
            notes = Trim$(notes & " Fee was " & Pct(op) & " up to " & Format$(om, "$0.00") & ", now " & Pct(np) & _
                " up to " & Format$(nm, "$0.00") & ".")
            If what = "" Then what = "Fee changed": look = "other"
        End If
    End If
    KindDiff = (what <> "")
End Function

Private Sub NewLine(ByRef nRecs As Variant, ByVal r As Long, ByVal changes As Object, ByRef others As Long)
    Dim kinds As Variant, i As Long, shown As Boolean
    kinds = WatchList()
    For i = 0 To UBound(kinds)
        If HasRate(nRecs, r, CStr(kinds(i)), False) Then
            changes.Add MakeChange("New", "new", LineNote(nRecs, r, True), CStr(kinds(i)), nRecs, 0, False, nRecs, r)
            shown = True
        End If
    Next
    If Not shown Then
        If AnyRate(nRecs, r, False) Then others = others + 1
    End If
End Sub

Private Sub GoneLine(ByRef oRecs As Variant, ByVal o As Long, ByVal newYes As Object, ByVal changes As Object, _
                     ByRef others As Long)
    Dim kinds As Variant, i As Long, shown As Boolean, notes As String, d As String
    notes = "Not in these files."
    d = UCase$(oRecs(o, C_DIST))
    If d = "" Or d = "R00" Then
        If Found(newYes, UCase$(oRecs(o, C_CODE))) > 0 Then notes = "Its rates are by business district now, on its other lines."
    End If
    kinds = WatchList()
    For i = 0 To UBound(kinds)
        If HasRate(oRecs, o, CStr(kinds(i)), False) Then
            changes.Add MakeChange("Gone", "gone", JoinNotes(LineNote(oRecs, o, False), notes, ""), CStr(kinds(i)), _
                oRecs, o, False, oRecs, 0)
            shown = True
        End If
    Next
    If Not shown Then
        If AnyRate(oRecs, o, False) Then others = others + 1
    End If
End Sub

Private Function MakeChange(ByVal what As String, ByVal look As String, ByVal notes As String, ByVal kind As String, _
                            ByRef oRecs As Variant, ByVal o As Long, ByVal oOld As Boolean, _
                            ByRef nRecs As Variant, ByVal r As Long) As Variant
    ' One line of the list. r is its record in the new files and o in the old ones, 0 for
    ' none. oOld: the old rate is the new record's own old-rate column.
    Dim f(0 To 13) As Variant, hi As Long, lo As Long, f1 As Long, f2 As Long, sc As Long
    Dim k As String, st As String, place As String, county As String, code As String, dist As String
    Dim lineNo As String, newHi As String, newLo As String
    KindCols kind, hi, lo, f1, f2
    sc = hi                                         ' the rates shown: the low ones if only they moved
    If Left$(what, 8) = "Low rate" Then sc = lo
    If r > 0 Then
        place = nRecs(r, C_PLACE): county = nRecs(r, C_COUNTY): code = nRecs(r, C_CODE)
        dist = nRecs(r, C_DIST): lineNo = nRecs(r, C_LINE): st = nRecs(r, C_START)
        k = LineKey(nRecs, r)
        f(F_NEW) = RateAt(nRecs, r, sc, False)
        newHi = NumText(RateAt(nRecs, r, hi, False))
        newLo = NumText(RateAt(nRecs, r, lo, False))
        ' A new name's start date is its rate's, not the name's, so it isn't shown
        If st <> "" And what <> "Name changed" Then f(F_START) = DateFrom8(st)
        If oOld Then f(F_OLD) = RateAt(nRecs, r, sc, True)
    Else
        place = oRecs(o, C_PLACE): county = oRecs(o, C_COUNTY): code = oRecs(o, C_CODE)
        dist = oRecs(o, C_DIST): lineNo = oRecs(o, C_LINE)
        k = LineKey(oRecs, o)
    End If
    If o > 0 And Not oOld Then f(F_OLD) = RateAt(oRecs, o, sc, False)
    f(F_PLACE) = place
    f(F_COUNTY) = county
    f(F_TYPE) = TypeWords(code, dist)
    f(F_WHAT) = what
    f(F_KIND) = KindTitle(kind)
    f(F_CODE) = code
    f(F_NOTES) = notes
    f(F_LOOK) = look
    f(F_KEY) = k & "|" & kind
    ' The same change seen again, by a later file, has the same ID
    f(F_ID) = k & "|" & kind & "|" & what & "|" & st & "|" & newHi & "|" & newLo & "|" & UCase$(notes)
    f(F_SORT) = UCase$(place) & Chr$(1) & UCase$(county) & Chr$(1) & TypeOrder(dist) & Chr$(1) & _
        KindOrder(kind) & Chr$(1) & UCase$(code) & Chr$(1) & lineNo
    MakeChange = f
End Function

Private Function LineNote(ByRef recs As Variant, ByVal i As Long, ByVal isNew As Boolean) As String
    ' What a line is, where its type doesn't say it all
    Dim d As String, cnt As Long
    d = UCase$(recs(i, C_DIST))
    If d = "" Then
        If Val(recs(i, C_LINE)) > 1 Then LineNote = "Line " & recs(i, C_LINE) & " for this code in the state's file."
    ElseIf d = "R00" Then
        LineNote = "Outside its business districts."
    Else
        LineNote = "District " & d
        If isNew Then
            cnt = Val(recs(i, C_ADDRS))
            If cnt = 1 Then
                LineNote = LineNote & ", 1 address in the address file"
            Else
                LineNote = LineNote & ", " & Format$(cnt, "#,##0") & " addresses in the address file"
            End If
        End If
        LineNote = LineNote & "."
    End If
End Function

Private Function SortedOrder(ByVal changes As Object) As Variant
    ' The changes by place name, then county, then district: a place's changes sit together.
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


' ---- the rate kinds ----------------------------------------------------------------------

Private Function WatchList() As Variant
    WatchList = Split(Replace(UCase$(WATCH), " ", ""), ",")
End Function

Private Function KindCols(ByVal kind As String, ByRef hi As Long, ByRef lo As Long, ByRef f1 As Long, _
                          ByRef f2 As Long) As Boolean
    ' A kind's columns: high rate, low rate, and for vehicles the fee percent and its most
    hi = 0: lo = 0: f1 = 0: f2 = 0
    Select Case UCase$(Trim$(kind))
        Case "GM": hi = C_GM: lo = C_GML
        Case "DM": hi = C_DM: lo = C_DML
        Case "PGM": hi = C_PGM: lo = C_PGML
        Case "PDM": hi = C_PDM: lo = C_PDML
        Case "AUTO": hi = C_AUTO: lo = C_AUTOL: f1 = C_FEEP: f2 = C_FEEMAX
        Case Else: Exit Function
    End Select
    KindCols = True
End Function

Private Function KindTitle(ByVal kind As String) As String
    Select Case UCase$(Trim$(kind))
        Case "GM": KindTitle = "General merchandise"
        Case "DM": KindTitle = "Drugs and medical"
        Case "PGM": KindTitle = "Purchases, general merchandise"
        Case "PDM": KindTitle = "Purchases, drugs and medical"
        Case "AUTO": KindTitle = "Vehicles"
        Case Else: KindTitle = kind
    End Select
End Function

Private Function WatchWords() As String
    ' "general merchandise rate", or "general merchandise and vehicles rates"
    Dim p As Variant, i As Long, s As String, w As String
    p = WatchList()
    For i = 0 To UBound(p)
        w = LCase$(KindTitle(CStr(p(i))))
        If i = 0 Then
            s = w
        ElseIf i = UBound(p) Then
            s = s & " and " & w
        Else
            s = s & ", " & w
        End If
    Next
    If UBound(p) > 0 Then s = s & " rates" Else s = s & " rate"
    WatchWords = s
End Function

Private Function KindOrder(ByVal kind As String) As String
    ' The order the kinds are listed in WATCH
    Dim p As Variant, i As Long
    p = WatchList()
    For i = 0 To UBound(p)
        If p(i) = UCase$(Trim$(kind)) Then KindOrder = Format$(i + 1, "00"): Exit Function
    Next
    KindOrder = "99"
End Function

Private Function RateAt(ByRef recs As Variant, ByVal i As Long, ByVal c As Long, ByVal useOld As Boolean) As Double
    If c = 0 Or i = 0 Then Exit Function
    If useOld Then c = c + OLD_OFF
    RateAt = Num(recs(i, c))
End Function

Private Function HasRate(ByRef recs As Variant, ByVal i As Long, ByVal kind As String, ByVal useOld As Boolean) As Boolean
    Dim hi As Long, lo As Long, f1 As Long, f2 As Long
    If Not KindCols(kind, hi, lo, f1, f2) Then Exit Function
    HasRate = (RateAt(recs, i, hi, useOld) <> 0 Or RateAt(recs, i, lo, useOld) <> 0)
End Function

Private Function AnyRate(ByRef recs As Variant, ByVal i As Long, ByVal useOld As Boolean) As Boolean
    Dim p As Variant, k As Long
    p = Split(ALL_KINDS, ",")
    For k = 0 To UBound(p)
        If HasRate(recs, i, CStr(p(k)), useOld) Then AnyRate = True: Exit Function
    Next
End Function

Private Function TypeWords(ByVal code As String, ByVal dist As String) As String
    ' Codes 200 and up are other states, there for vehicles. xxx-5000 is a county outside its towns.
    If Val(Left$(code, 3)) >= 200 Then
        TypeWords = "Other state"
    ElseIf dist <> "" And UCase$(dist) <> "R00" Then
        TypeWords = "Business district"
    ElseIf Mid$(code, 5, 4) = "5000" Then
        TypeWords = "County"
    Else
        TypeWords = "City"
    End If
End Function

Private Function TypeOrder(ByVal dist As String) As String
    If dist = "" Or UCase$(dist) = "R00" Then TypeOrder = "0" Else TypeOrder = "1" & UCase$(dist)
End Function


' ---- the sheets ------------------------------------------------------------------------

Private Sub EnsureSheets()
    ' First run: makes the four sheets. Later runs put back the button, and any sheet that
    ' went missing.
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
    If SheetByName(SH_READ) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_MAIN))
        ws.Name = SH_READ
        DrawReadHere ws
    End If
    If SheetByName(SH_LOG) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=SheetByName(SH_READ))
        ws.Name = SH_LOG
        LogHeader ws
    End If
    If SheetByName(SH_SAVED) Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SH_SAVED
        ws.Range("A1").Value = "Nothing saved yet. Loading the files puts their rates here."
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
        .Value = "Click the button and pick the state's two rates files."
        .Font.Size = 20
    End With
    ws.Rows(3).RowHeight = 34
    ws.Range("A4").Value = "Get them from tax.illinois.gov: Research, Tax Rates, Sales Tax Rate - Machine Readable " & _
        "Files. Under Sales Tax, the top two links: the Effective file and the Addendum Address File."
    ws.Rows(4).RowHeight = 24
    ws.Range("A5").Value = "How it works, step by step: the Read here tab."
    ws.Rows(5).RowHeight = 24
    FreezeAt ws, 1
    ActiveWindow.DisplayGridlines = False
    DrawButton ws
End Sub

Private Sub DrawReadHere(ByVal ws As Worksheet)
    ' How to work it, in the same words as the notes posted with the macro. Redrawn on each
    ' load, so it keeps up with a pasted update.
    Dim r As Long
    ClearSheet ws
    ws.Cells.Font.Size = 16
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 105
    r = 1
    ReadLine ws, r, "T", "How this works"
    ReadLine ws, r, "P", "This shows what changed in Illinois's sales tax rates since the last files you loaded: " & _
        "each place whose rate changed, its county, the old rate, the new rate, and when it starts. It watches " & _
        "the " & WatchWords() & "."
    ReadLine ws, r, "H", "When to check"
    ReadLine ws, r, "P", "Illinois posts new rates for January 1 and July 1, and sometimes another day (August 1 " & _
        "was one in 2026). Once that day comes, What changed shows a red GET THE NEW FILES line."
    ReadLine ws, r, "H", "Getting the files"
    ReadLine ws, r, "P", "1. Go to tax.illinois.gov, Research, Tax Rates, Sales Tax Rate - Machine Readable Files."
    ReadLine ws, r, "P", "2. Under Sales Tax, download the top two links: Effective (it saves as " & _
        "ordmache-current.txt) and Addendum Address File (ordmacha-current.txt). The second one has the rates " & _
        "for places with business districts."
    ReadLine ws, r, "P", "3. Not the Address Specific page. That file is 2 GB and this won't read it."
    ReadLine ws, r, "H", "Loading them"
    ReadLine ws, r, "P", "1. Click Check new rates file on the What changed tab. The box opens in Downloads."
    ReadLine ws, r, "P", "2. Hold Ctrl, click both files, then Open. Or pick one and it asks for the other."
    ReadLine ws, r, "P", "3. Save the workbook when you're done. It never saves itself."
    ReadLine ws, r, "H", "Reading What changed"
    ReadLine ws, r, "P", "Each line is a place, or one business district in a place, whose rate changed. " & _
        "Starts is the day the new rate begins."
    ReadLine ws, r, "up", "Red|went up."
    ReadLine ws, r, "down", "Green|went down."
    ReadLine ws, r, "new", "Blue|new."
    ReadLine ws, r, "gone", "Gray|gone."
    ReadLine ws, r, "other", "Yellow|something else, like a new name."
    ReadLine ws, r, "P", "Places with business districts get a line for each district (District R01 and so on). " & _
        "Those rates come from the address file."
    ReadLine ws, r, "P", "The first time, with nothing loaded before, it shows what changed on the files' start date."
    ReadLine ws, r, "P", "BUSINESS DISTRICTS NOT CHECKED means the address file wasn't picked. Load both next time."
    ReadLine ws, r, "H", "The other tabs"
    ReadLine ws, r, "P", "Log: every change once, the first time a file showed it, with the day it was loaded. " & _
        "Nothing new means a load found nothing new. The buttons on its top row filter it."
    ReadLine ws, r, "P", "Saved rates: the rates from the last files loaded. The next files get checked against " & _
        "them. It's locked, so leave it alone."
    ReadLine ws, r, "H", "If a box asks something"
    ReadLine ws, r, "P", "If a yellow bar says macros are disabled, click Enable Content."
    ReadLine ws, r, "P", "If it says the files look older than last time, or cut short, or not from the same " & _
        "posting, answer No and download them again."
    ReadLine ws, r, "N", "This tab comes from the macro and is redrawn each time files load."
    FreezeAt ws, 0
    ActiveWindow.DisplayGridlines = False
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo 0
End Sub

Private Sub ReadLine(ByVal ws As Worksheet, ByRef r As Long, ByVal kind As String, ByVal words As String)
    ' One line of Read here. kind: T title, H heading, P plain, N small gray note, or a look
    ' (up, down, new, gone, other): the color's name in its color, then what it means.
    Dim p As Long
    If kind = "H" Then
        ws.Rows(r).RowHeight = 14           ' a gap above each heading
        r = r + 1
    End If
    Select Case kind
        Case "T", "H"
            With ws.Cells(r, 1)
                .Value = words
                .Font.Bold = True
                If kind = "T" Then
                    .Font.Size = 26
                Else
                    .Font.Size = 20
                    .Font.Color = RGB(31, 78, 121)
                End If
            End With
            If kind = "T" Then ws.Rows(r).RowHeight = 50 Else ws.Rows(r).RowHeight = 34
        Case "P", "N"
            With ws.Cells(r, 2)
                .Value = words
                .WrapText = True
                If kind = "N" Then
                    .Font.Size = 14
                    .Font.Color = RGB(89, 89, 89)
                End If
            End With
            ws.Rows(r).AutoFit
            ws.Rows(r).RowHeight = ws.Rows(r).RowHeight + 6     ' a little air between lines
            If ws.Rows(r).RowHeight < 26 Then ws.Rows(r).RowHeight = 26
        Case Else
            p = InStr(words, "|")
            With ws.Cells(r, 1)
                .Value = Left$(words, p - 1)
                .Font.Bold = True
                .Interior.Color = LookFill(kind)
                .Font.Color = LookText(kind)
                .IndentLevel = 1
            End With
            ws.Cells(r, 2).Value = Mid$(words, p + 1)
            ws.Rows(r).RowHeight = 30
    End Select
    ws.Rows(r).VerticalAlignment = xlCenter
    r = r + 1
End Sub

Private Sub DrawFront(ByVal changes As Object, ByVal others As Long, ByVal files As String, ByVal loadedOn As Date, _
                      ByVal change As String, ByVal hasAddr As Boolean, ByVal compared As Boolean, _
                      ByVal savedFiles As String, ByVal savedWhen As Variant, ByVal savedChange As String, _
                      ByVal savedAddr As Boolean)
    ' What changed: which files, what they were checked against, then the list in big type.
    Dim ws As Worksheet, n As Long, s As String, r As Long
    n = changes.Count
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
    ws.Range("A2").Value = "From " & files & ", loaded " & Format$(loadedOn, "m/d/yyyy") & "."
    ws.Range("A3").Value = "These files are for rates starting " & ShowDate8(change) & _
        ". Illinois rates usually change January 1 and July 1."
    If compared Then
        s = "Checked against " & savedFiles & ", loaded " & ShowWhen(savedWhen) & ", for rates starting " & _
            ShowDate8(savedChange) & "."
        If hasAddr And Not savedAddr Then
            s = s & " Those had no address file, so business districts show what changed on " & ShowDate8(change) & "."
        End If
    Else
        s = "Nothing loaded before, so this shows what changed on " & ShowDate8(change) & _
            ", from the old rates the state lists beside each rate."
    End If
    ws.Range("A4").Value = s
    s = "Showing only the " & WatchWords() & "."
    If others = 1 Then s = s & " 1 other change, in rates you don't watch, isn't shown."
    If others > 1 Then s = s & " " & Format$(others, "#,##0") & " other changes, in rates you don't watch, aren't shown."
    ws.Range("A5").Value = s
    ws.Range("A5").Font.Color = RGB(89, 89, 89)
    For r = 2 To 5
        ws.Rows(r).RowHeight = 24
    Next
    ws.Rows(DUE_ROW).RowHeight = 4
    If hasAddr Then
        ws.Rows(ADDR_ROW).RowHeight = 4
    Else
        Band ws, ADDR_ROW, "BUSINESS DISTRICTS NOT CHECKED: NO ADDRESS FILE.", LookFill("other"), LookText("other"), 20
    End If
    ws.Rows(BAND_ROW - 1).RowHeight = 12

    If n > 0 Then
        If n = 1 Then s = "1 CHANGE" Else s = Format$(n, "#,##0") & " CHANGES"
        Band ws, BAND_ROW, s, RGB(255, 217, 102), RGB(0, 0, 0), 28
        ws.Cells(BAND_ROW + 1, 1).Value = "Red went up, green went down, blue is new, gray is gone, yellow is something else."
        DrawList ws, LIST_ROW, changes
    Else
        Band ws, BAND_ROW, "NO CHANGES", RGB(198, 239, 206), RGB(0, 97, 0), 28
        If compared Then
            ws.Cells(BAND_ROW + 1, 1).Value = "Nothing you watch changed since the last files."
        Else
            ws.Cells(BAND_ROW + 1, 1).Value = "Nothing you watch changed on " & ShowDate8(change) & "."
        End If
    End If
    ws.Rows(BAND_ROW + 1).RowHeight = 24

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

Private Sub DueLine()
    ' The red line on What changed once newer files should be out: from the next January 1 or
    ' July 1 after the saved files' start date.
    Dim ws As Worksheet, sv As Worksheet, change As String, d As Date, due As Date
    Set ws = SheetByName(SH_MAIN)
    Set sv = SheetByName(SH_SAVED)
    If ws Is Nothing Or sv Is Nothing Then Exit Sub
    If CStr(sv.Cells(SAVED_ROW, 1).Value) <> "Code" Then Exit Sub
    change = SavedChangeDate(sv)
    If Len(change) <> 8 Then Exit Sub
    d = DateFrom8(change)
    If Month(d) < 7 Then due = DateSerial(Year(d), 7, 1) Else due = DateSerial(Year(d) + 1, 1, 1)
    With ws.Range(ws.Cells(DUE_ROW, 1), ws.Cells(DUE_ROW, 7))
        .UnMerge
        .Clear
    End With
    If CheckDate() >= due Then
        Band ws, DUE_ROW, "GET THE NEW FILES. RATES USUALLY CHANGE " & Format$(due, "m/d/yyyy") & ".", _
            RGB(255, 199, 206), RGB(156, 0, 6), 24
    Else
        ws.Rows(DUE_ROW).RowHeight = 4
    End If
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
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(r).RowHeight = pts * 1.8
End Sub

Private Sub DrawList(ByVal ws As Worksheet, ByVal r As Long, ByVal changes As Object)
    ' The header row at r, then one row per change
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
    ws.Range(ws.Cells(r + 1, 9), ws.Cells(r + n, 9)).NumberFormat = "@"
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, LIST_COLS))
        .Value = out
        .Font.Size = 16
    End With
    With ws.Range(ws.Cells(r + 1, 5), ws.Cells(r + n, 6))
        .NumberFormat = "0.00##""%"""
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(r + 1, 7), ws.Cells(r + n, 7)).NumberFormat = "m/d/yyyy"
    For i = 1 To n
        ch = changes(order(i))
        With ws.Range(ws.Cells(r + i, 1), ws.Cells(r + i, LIST_COLS)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(191, 191, 191)
        End With
        With ws.Cells(r + i, 4)
            .Interior.Color = LookFill(CStr(ch(F_LOOK)))
            .Font.Color = LookText(CStr(ch(F_LOOK)))
            .Font.Bold = True
        End With
    Next
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 4)).WrapText = True
    ws.Range(ws.Cells(r + 1, LIST_COLS), ws.Cells(r + n, LIST_COLS)).WrapText = True
    ws.Range(ws.Cells(r + 1, LIST_COLS), ws.Cells(r + n, LIST_COLS)).IndentLevel = 1
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + n, 1)).EntireRow.AutoFit
    For rr = r + 1 To r + n
        If ws.Rows(rr).RowHeight < 30 Then ws.Rows(rr).RowHeight = 30
    Next
End Sub

Private Sub ListWidths(ByVal ws As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(30, 18, 25, 24, 12, 12, 15, 28, 16, 50)
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next
End Sub

Private Sub AddToLog(ByVal changes As Object, ByVal files As String, ByVal loadedOn As Date, ByVal against As String)
    ' Each change goes on the Log once, the first time a file shows it. A check with nothing
    ' new still gets a line.
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
                out(n, 2) = files
                out(n, 3) = against
                For c = 1 To LIST_COLS
                    out(n, c + 3) = ch(c)
                Next
                out(n, LOG_COLS) = ch(F_ID)
                looks(n) = CStr(ch(F_LOOK))
            End If
        Next
    End If
    If n = 0 Then
        n = 1
        out(1, 1) = loadedOn
        out(1, 2) = files
        out(1, 3) = against
        out(1, 7) = "Nothing new"
        out(1, 13) = "Everything these files show was already on the Log."
        looks(1) = ""
    End If

    r = last + 1
    If r < 2 Then r = 2
    ws.Range(ws.Cells(r, 12), ws.Cells(r + n - 1, 12)).NumberFormat = "@"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, LOG_COLS))
        .Value = out
        .Font.Size = 14
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 1)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 10), ws.Cells(r + n - 1, 10)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(r, 8), ws.Cells(r + n - 1, 9)).NumberFormat = "0.00##""%"""
    ws.Range(ws.Cells(r, 13), ws.Cells(r + n - 1, 13)).WrapText = True
    For i = 1 To n
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
    widths = Array(13, 30, 30, 28, 16, 22, 22, 11, 11, 13, 26, 13, 50, 10)
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

Private Sub SaveCopy(ByRef recs As Variant, ByVal n As Long, ByVal files As String, ByVal loadedOn As Date, _
                     ByVal change As String, ByVal hasAddr As Boolean)
    ' Saved rates: every row of the files (a district once, its addresses counted), for the
    ' next files to be checked against
    Dim ws As Worksheet, out As Variant, i As Long, c As Long, titles As Variant
    Set ws = SheetByName(SH_SAVED)
    ws.Unprotect
    ws.Cells.Clear
    ws.Cells.Font.Size = 11
    With ws.Range("A1")
        .Value = "The rates from the last files loaded. The next files get checked against them. " & _
            "Loading files replaces them, so leave this sheet alone."
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Range("A2").Value = "Files"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = files
    ws.Range("C2").Value = "Loaded"
    ws.Range("D2").Value = loadedOn
    ws.Range("D2").NumberFormat = "m/d/yyyy"
    ws.Range("E2").Value = "Rates start"
    ws.Range("F2").Value = DateFrom8(change)
    ws.Range("F2").NumberFormat = "m/d/yyyy"
    ws.Range("G2").Value = "Address file"
    If hasAddr Then ws.Range("H2").Value = "Yes" Else ws.Range("H2").Value = "No"
    ws.Range("A2,C2,E2,G2").Font.Bold = True
    titles = Split(HEADS, "|")
    For c = 1 To NC
        ws.Cells(SAVED_ROW, c).Value = titles(c - 1)
    Next
    With ws.Range(ws.Cells(SAVED_ROW, 1), ws.Cells(SAVED_ROW, NC))
        .Font.Bold = True
        .WrapText = True
    End With
    ReDim out(1 To n, 1 To NC)
    For i = 1 To n
        For c = 1 To NC
            out(i, c) = CellValue(CStr(recs(i, c)), c)
        Next
    Next
    For c = 1 To NC
        With ws.Range(ws.Cells(SAVED_ROW + 1, c), ws.Cells(SAVED_ROW + n, c))
            If IsTextColumn(c) Then
                .NumberFormat = "@"
            ElseIf c = C_FEEMAX Or c = C_FEEMAX + OLD_OFF Then
                .NumberFormat = "$0.00"
            ElseIf c = C_ADDRS Or c = C_LINE Then
                .NumberFormat = "0"
            Else
                .NumberFormat = "0.00##""%"""
            End If
        End With
    Next
    ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(SAVED_ROW + n, NC)).Value = out
    ws.Range(ws.Columns(1), ws.Columns(NC)).ColumnWidth = 12
    ws.Columns(C_PLACE).ColumnWidth = 26
    ws.Columns(C_COUNTY).ColumnWidth = 16
    ws.Rows(SAVED_ROW).RowHeight = 45
    ws.Protect
End Sub

Private Function ReadSaved(ByRef recs As Variant, ByRef files As String, ByRef loadedOn As Variant, _
                           ByRef change As String, ByRef hadAddr As Boolean) As Long
    ' Every row on Saved rates, as text the way the files had it. 0 when nothing's saved yet.
    Dim ws As Worksheet, last As Long, n As Long, i As Long, c As Long
    Set ws = SheetByName(SH_SAVED)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Cells(SAVED_ROW, 1).Value) <> "Code" Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last <= SAVED_ROW Then Exit Function
    n = last - SAVED_ROW
    recs = ws.Range(ws.Cells(SAVED_ROW + 1, 1), ws.Cells(last, NC)).Value
    For i = 1 To n
        For c = 1 To NC
            recs(i, c) = AsText(recs(i, c))
        Next
    Next
    files = CStr(ws.Range("B2").Value)
    loadedOn = ws.Range("D2").Value
    change = SavedChangeDate(ws)
    hadAddr = (CStr(ws.Range("H2").Value) = "Yes")
    ReadSaved = n
End Function

Private Function SavedChangeDate(ByVal ws As Worksheet) As String
    If IsDate(ws.Range("F2").Value) Then SavedChangeDate = Format$(CDate(ws.Range("F2").Value), "yyyymmdd")
End Function

Private Function IsTextColumn(ByVal c As Long) As Boolean
    Select Case c
        Case C_CODE, C_DIST, C_PLACE, C_COUNTY, C_AFILE, C_START, C_RFLAG, C_PFLAG, C_VFLAG, C_PSTART, C_PEND, _
             C_RFLAG + OLD_OFF, C_PFLAG + OLD_OFF, C_VFLAG + OLD_OFF
            IsTextColumn = True
    End Select
End Function

Private Function CellValue(ByVal s As String, ByVal c As Long) As Variant
    ' How a value goes onto Saved rates: words stay words, rates and counts go in as numbers
    If s = "" Then
        CellValue = Empty
    ElseIf IsTextColumn(c) Then
        CellValue = s
    Else
        CellValue = Val(s)
    End If
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
    b.Left = ws.Range("E1").Left + 10
    b.Top = ws.Range("A1").Top + 4
    b.Width = 290
    b.Height = ws.Range("A1").Height - 8
    b.Caption = BTN_TEXT
    b.OnAction = "CheckIllinoisRates"
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

Private Function Found(ByVal idx As Object, ByVal k As String) As Long
    On Error Resume Next
    Found = idx(k)
    If Err.Number <> 0 Then Found = 0
    On Error GoTo 0
End Function

Private Function InList(ByVal s As String, ByVal list As String) As Boolean
    InList = InStr(1, "," & Replace(UCase$(list), " ", "") & ",", "," & UCase$(Trim$(s)) & ",", vbBinaryCompare) > 0
End Function

Private Function JoinNotes(ByVal a As String, ByVal b As String, ByVal c As String) As String
    JoinNotes = Trim$(Trim$(Trim$(a) & " " & Trim$(b)) & " " & Trim$(c))
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

Private Function OrBlank(ByVal s As String) As String
    If Trim$(s) = "" Then OrBlank = "blank" Else OrBlank = Trim$(s)
End Function

Private Function Num(ByVal s As String) As Double
    Num = Val(Replace(Replace(Replace(Trim$(s), "$", ""), "%", ""), ",", ""))
End Function

Private Function NumText(ByVal x As Double) As String
    ' A number the same way every time: 6.25, 10, .75
    If x = Int(x) Then NumText = Format$(x, "0") Else NumText = Trim$(Str$(x))
End Function

Private Function Pct(ByVal x As Double) As String
    Pct = Format$(x, "0.00##") & "%"
End Function

Private Function AsText(ByVal v As Variant) As String
    ' A value read back off Saved rates, as text the way the files had it
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

Private Function PickFiles(ByVal several As Boolean, ByVal title As String) As Variant
    ' The files picked, as a list, or False if the box was cancelled. Starts in Downloads.
    Dim dl As String, got As Variant
    If Not mFeed Is Nothing Then
        If mFeed.Count = 0 Then
            PickFiles = False
        Else
            got = mFeed(1)
            mFeed.Remove 1
            If got = "<cancel>" Or got = "" Then PickFiles = False Else PickFiles = Split(got, ";")
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
    got = Application.GetOpenFilename("Rates files (*.txt),*.txt", 1, title, , several)
    If IsArray(got) Then
        PickFiles = got
    ElseIf VarType(got) = vbBoolean Then
        PickFiles = False
    Else
        PickFiles = Array(CStr(got))
    End If
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
    ' For test runs: feeds answers to the boxes (split on |; files picked together split on ;),
    ' acts as if today were todayText (yyyy-mm-dd, blank for the real day), and returns what
    ' the boxes would have said. "<open>" as the answers does what opening the workbook does
    ' instead of the button. Not needed day to day.
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
        CheckIllinoisRates
    End If
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    RatesTest = mSaid
    Set mFeed = Nothing
    mToday = 0
End Function
