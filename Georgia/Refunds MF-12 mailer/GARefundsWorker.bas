Attribute VB_Name = "GARefundsWorker"
' GA refunds, the worker: the three buttons, and everything with folders, files and Outlook;
' it draws CHECK, ENTRY and HOW TO. GARefundsReader reads the tabs. The HOW TO tab has the
' steps. Each packet and each email is worked on its own, so one problem holds only its own;
' it never sends an email, deletes a file or writes over one.

Option Explicit

Private Const BIG_EMAIL_MB As Double = 20      ' a draft bigger than this gets a warning on CHECK
Private Const TAG_FILE As String = "packet.txt"   ' in each packet folder: whose packet it is
' ENTRY's columns, in the portal's order. Keyed is first: the x's are read back by these words.
Private Const ENTRY_HEADS As String = "Keyed|Seller Name|Seller License Number|Date|Invoice Number|Gallons|" & _
    "Tax Paid to Seller|Exempt Portion|Type of Fuel|Running total"

' Where a packet stands, in the order it gets there
Private Const ST_HELD As Long = 1        ' a row or the supplier needs a fix first
Private Const ST_FILES As Long = 2       ' waiting on a VI, or its MF-12 isn't right
Private Const ST_SEND As Long = 3        ' ready for Make drafts
Private Const ST_DRAFT As Long = 4       ' in Outlook Drafts, not sent yet
Private Const ST_OUT As Long = 5         ' sent, waiting on the notarized MF-12
Private Const ST_BACK As Long = 6        ' back notarized: key it from ENTRY, then Mark filed

Private mMoved As Long           ' paid rows moved to Refunds Received this click
Private mMonthDirs As Variant    ' the folders in the packets folder, to find a month's whatever its capitals
Private mTemplate As String      ' the template's bytes
Private mQuiet As Boolean        ' test runs: no popups
Private mFeed As Collection      ' test runs: the answers to the questions, in order
Private mTestDraftDir As String  ' test runs: drafts go to text files here instead of Outlook
Private mNotes As String
Private mRenamed As Collection   ' every rename this click, "old path|new name", for CHECK
Private mLastSaid As String      ' what the last box said, for CHECK when a click stops early
Private mWaitOnly() As Boolean   ' an email held only by an invoice not in yet or an MF-12 not made yet
' the packet.txt files found this click: the folder, the packet it names, and what else it says
Private mTagDir() As String, mTagKey() As String, mTagShort() As String, mTagSubject() As String, mNTag As Long

' ======================================================================================
'  The three macros
' ======================================================================================

Public Sub Refunds1_BuildPackets()
    RunStep 1
End Sub

Public Sub Refunds2_MakeDrafts()
    RunStep 2
End Sub

Public Sub Refunds3_MarkFiled()
    RunStep 3
End Sub

' Runs a step with no popups and hands back what it would have said, for trying it on a copy.
' answers are what you'd type into its questions, in order, split on |.
Public Function RefundsTest(ByVal stepNo As Long, ByVal answers As String, ByVal draftDir As String) As String
    Dim a As Variant
    mQuiet = True
    Set mFeed = New Collection
    If Len(answers) > 0 Then
        For Each a In Split(answers, "|")
            mFeed.Add CStr(a)
        Next a
    End If
    mTestDraftDir = draftDir
    mNotes = ""
    RunStep stepNo
    mQuiet = False
    mTestDraftDir = ""
    Set mFeed = Nothing
    RefundsTest = mNotes
End Function

Private Sub RunStep(ByVal stepNo As Long)
    Dim msg As String, marked As Long, drawIt As Boolean, errText As String

    On Error GoTo Failed
    mStage = "getting started"
    ' Excel keeps these between runs, so start every run from nothing
    mTemplate = "": mMoved = 0: mLastSaid = "": mMonthDirs = Empty
    Erase mWaitOnly
    Set mRenamed = New Collection
    ClearLists True
    Set mBook = Book()
    If mBook Is Nothing Then
        Say "Can't find the Refunds Due sheet. This macro goes in the GA tracking workbook (the one with " & _
            "Refunds Due and Corp License Number) and click the button there.", vbExclamation
        Exit Sub
    End If
    ' From here on a click that has to stop still draws CHECK, saying why.
    If Len(mBook.Path) = 0 Then
        StopEarly stepNo, "Save this workbook first: File, Save As, and pick Excel Macro-Enabled Workbook (.xlsm). " & _
            "Then click Build packets again."
        Exit Sub
    End If

    mStage = "writing the HOW TO tab"
    WriteHowTo

    mStage = "checking the setup"
    msg = EnsureSetup()
    If Len(msg) > 0 Then StopEarly stepNo, msg: Exit Sub
    If Not AskSetup() Then WriteCheckStopped stepNo, mLastSaid: Exit Sub
    msg = ReadSetup()
    If Len(msg) > 0 Then StopEarly stepNo, "Fill these in on the SETUP sheet first:" & vbLf & msg: Exit Sub
    If Not LoadTemplate() Then
        StopEarly stepNo, "The MF-12 template on SETUP isn't there, or isn't the right file:" & vbLf & mTemplatePath & vbLf & _
            "It's " & TEMPLATE_NAME & ", which comes with this macro. Save it beside this workbook, clear the " & _
            "MF-12 template file cell on SETUP, and click Build packets again."
        Exit Sub
    End If
    drawIt = True        ' from here on, every click ends by drawing CHECK and ENTRY

    mStage = "moving paid rows to Refunds Received"
    mMoved = MoveReceived()

    Application.ScreenUpdating = False
    ReadAll
    If stepNo = 1 Then
        mStage = "making folders under " & mBase
        MakeFolders
    End If
    ' only Build packets files them; the other buttons say the same about INBOX without moving anything
    mStage = IIf(stepNo = 1, "filing the INBOX", "looking in the INBOX")
    SortInbox (stepNo = 1)

    mStage = "checking the packet folders"
    CheckFiles
    ' Build packets fills and saves the MF-12s; the other buttons fill them without saving, so a
    ' form that can't be filled (a name too long for its box) shows on CHECK whichever is clicked.
    mStage = "filling the MF-12s"
    WriteForms (stepNo = 1)
    mStage = "checking what's been sent"
    CheckSent (stepNo = 2)
    mStage = "stamping MF-12 Sent and MF-12 Signed"
    WriteStamps
    mStage = "deciding who is ready"
    DecideReady
    If stepNo = 2 Then
        mStage = "making the Outlook drafts"
        MakeDrafts
        SetStages
    ElseIf stepNo = 3 Then
        mStage = "marking packets filed"
        Application.ScreenUpdating = True
        marked = MarkFiled()
        Application.ScreenUpdating = False
        If marked > 0 Then
            ' read it all again, so the filed packets drop off
            ReadAll
            SortInbox False
            CheckFiles
            CheckSent False
            DecideReady
        End If
    End If

    mStage = "writing the ENTRY sheet"
    WriteEntry
    mStage = "writing the CHECK sheet"
    WriteCheck stepNo
    Application.ScreenUpdating = True
    If stepNo <> 3 Or marked > 0 Then Say Finished(stepNo, marked), vbInformation
    Exit Sub

Failed:
    Close
    errText = Err.Description
    Resume Stopped
Stopped:
    ' something no one packet could hold: say so on CHECK too, once there's a CHECK to draw
    On Error Resume Next
    msg = "Stopped with an error while " & mStage & ":" & vbLf & errText
    If drawIt Then
        AddLook "Stopped while " & mStage, errText, "Fix that, then click the button again. Nothing after that step was done on this click."
        WriteEntry
        WriteCheck stepNo
        msg = msg & vbLf & "CHECK lists it."
    ElseIf Not mBook Is Nothing Then
        WriteCheckStopped stepNo, msg
    End If
    Application.ScreenUpdating = True
    Say msg, vbCritical
End Sub

' A click that can't go on: the box says why, and so does CHECK.
Private Sub StopEarly(ByVal stepNo As Long, ByVal msg As String)
    Say msg, vbExclamation
    WriteCheckStopped stepNo, msg
End Sub

' Reads the tabs again from the top, and puts each packet in its folder.
Private Sub ReadAll()
    ClearLists
    LoadAll
    mStage = "finding the packet folders"
    PlaceFolders
End Sub

' Every packet's folder, found one packet at a time, then every rename that's needed, one packet
' at a time. A folder it can't look in or rename holds only that packet.
Private Sub PlaceFolders()
    Dim p As Long, pass As Long, taken As String, waiting() As Boolean

    FixMonthFolders
    mMonthDirs = ListFolders(mBase)
    ReadTags
    For p = 1 To mNPkt
        PlacePacket p, taken
    Next p

    ' A name another packet is giving up gets a second try, once the rest are done.
    If mNPkt = 0 Then Exit Sub
    ReDim waiting(1 To mNPkt)
    For p = 1 To mNPkt
        waiting(p) = (Len(mPkt(p).Folder) > 0 And Not StoppedOn(p))
    Next p
    For pass = 1 To 2
        For p = 1 To mNPkt
            If waiting(p) Then waiting(p) = Not RenamePacket(p, pass = 2)
        Next p
    Next pass
End Sub

' One packet's folder, inside its period's folder: the one whose packet.txt names it; else the
' one with the name it should have (its packet.txt gets written); else a folder the last version
' made (LIC-ACME, or one under an older Short Name that holds its files), to be renamed. Two
' folders whose packet.txt names the same packet, or two old folders that both hold its files,
' hold that packet, and so does a supplier that needs a fix first, so nothing of theirs is
' touched. taken: the old folders other packets claimed this click.
Private Sub PlacePacket(ByVal p As Long, ByRef taken As String)
    Dim i As Long, n As Long, seen As String, want As String, claims As String, olds As Variant
    On Error GoTo Failed
        With mPkt(p)
            .Folder = "": .Subject = "": .TagShort = ""
            If Len(mSup(.Sup).Problem) = 0 Then
                want = MonthFolder(.PerKey) & "\" & .PName
                n = 0: seen = ""
                For i = 1 To mNTag
                    If mTagKey(i) = .Key Then
                        n = n + 1
                        seen = seen & IIf(n > 1, " and ", "") & LastPart(mTagDir(i))
                        If n = 1 Then .Folder = mTagDir(i): .Subject = mTagSubject(i): .TagShort = mTagShort(i)
                    End If
                Next i
                If n > 1 Then
                    .Folder = ""
                    .Problem = "two folders say they're this packet: " & seen
                    AddLook PeriodLabel(.PerKey) & " " & .PName & ": two folders say they're this packet", seen & ", in " & _
                        MonthFolderName(.PerKey) & ". Each has a packet.txt for license " & mSup(.Sup).License & ", " & .Word & ", " & _
                        PeriodLabel(.PerKey) & ".", "Keep the right one: move the other out of " & MonthFolderName(.PerKey) & _
                        ", or delete its packet.txt, then click Build packets again."
                ElseIf n = 0 Then
                    If FolderExists(want) And TagAt(want) > 0 Then
                        .Problem = "the folder " & .PName & " is another packet's"
                        AddLook PeriodLabel(.PerKey) & " " & .PName & ": the folder " & .PName & " is another packet's", _
                            "Its packet.txt is for another supplier, refund type or period.", _
                            "Change this supplier's Short Name on " & CONTACTS_SHEET & ", or move that folder out, then click Build packets again."
                    ElseIf FolderExists(want) Then
                        .Folder = want
                    Else
                        .Folder = want        ' new, unless a folder the last version made holds it
                        If Len(.Prefix) > 0 Then claims = OldClaims(p, taken) Else claims = ""
                        If Len(claims) > 0 Then
                            olds = Split(Mid$(claims, 2), "|")
                            If UBound(olds) = 0 Then
                                .Folder = olds(0)
                                taken = taken & "|" & UCase$(olds(0)) & "|"
                            Else
                                seen = ""
                                For i = 0 To UBound(olds)
                                    seen = seen & IIf(i > 0, " and ", "") & LastPart(CStr(olds(i)))
                                Next i
                                .Folder = ""
                                .Problem = "two folders hold its files: " & seen
                                AddLook PeriodLabel(.PerKey) & " " & .PName & ": two folders hold its files", seen & ", in " & _
                                    MonthFolderName(.PerKey) & ". Each has an invoice or MF-12 of this packet.", _
                                    "Move everything into one of them and the other out of " & MonthFolderName(.PerKey) & _
                                    ", then click Build packets again."
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Exit Sub
Failed:
    mPkt(p).Folder = ""             ' not known this click, so nothing of its is touched
    HoldPacket p, "finding its folder"
End Sub

' True once something stopped partway on this packet this click.
Private Function StoppedOn(ByVal p As Long) As Boolean
    StoppedOn = (Left$(mPkt(p).Problem, 14) = "stopped while ")
End Function

' A period's folder in the capitals the last version gave it (2026-08 AUGUST) gets this
' version's (2026-08 August), unless it holds a packet that's filed: what's filed stays where
' it was filed. Windows finds it either way, so one that isn't renamed goes on as it is.
Private Sub FixMonthFolders()
    Dim p As Long, i As Long, dirs As Variant, want As String, done As String, why As String, errNo As Long
    dirs = ListFolders(mBase)
    If IsEmpty(dirs) Then Exit Sub
    For p = 1 To mNPkt
        want = MonthFolderName(mPkt(p).PerKey)
        If InStr(done, "|" & want & "|") = 0 Then
            done = done & "|" & want & "|"
            For i = 1 To UBound(dirs)
                If UCase$(dirs(i)) = UCase$(want) And dirs(i) <> want Then
                    If Not HoldsClosedPacket(mBase & "\" & dirs(i)) Then
                        errNo = RenameOne(mBase & "\" & dirs(i), mBase & "\" & want, why)
                        If errNo <> 0 Then AddLookItem "Couldn't rename this month folder, so it works as it is meanwhile:", _
                            "Couldn't rename these month folders, so they work as they are meanwhile:", _
                            dirs(i) & " to " & want, why, FixFor(errNo, "Build packets")
                    End If
                End If
            Next i
        End If
    Next p
End Sub

' True when a month folder holds a packet folder whose packet.txt names a packet that isn't open
' this click: one filed or paid already. One it can't look in counts as holding one.
Private Function HoldsClosedPacket(ByVal mf As String) As Boolean
    Dim dirs As Variant, i As Long, p As Long, key As String, sh As String, subj As String, isOpen As Boolean
    On Error GoTo Unsure
    dirs = ListFolders(mf)
    If IsEmpty(dirs) Then Exit Function
    For i = 1 To UBound(dirs)
        If ParseTag(mf & "\" & dirs(i), key, sh, subj) Then
            isOpen = False
            For p = 1 To mNPkt
                If mPkt(p).Key = key Then isOpen = True: Exit For
            Next p
            If Not isOpen Then HoldsClosedPacket = True: Exit Function
        End If
    Next i
    Exit Function
Unsure:
    HoldsClosedPacket = True
End Function

' Every packet.txt in the period folders this click works in.
Private Sub ReadTags()
    Dim p As Long, i As Long, mf As String, done As String, dirs As Variant
    mNTag = 0
    Erase mTagDir, mTagKey, mTagShort, mTagSubject
    For p = 1 To mNPkt
        mf = MonthFolder(mPkt(p).PerKey)
        If InStr(done, "|" & UCase$(mf) & "|") = 0 Then
            done = done & "|" & UCase$(mf) & "|"
            dirs = ListFolders(mf)
            If Not IsEmpty(dirs) Then
                For i = 1 To UBound(dirs)
                    ReadTag mf & "\" & dirs(i)
                Next i
            End If
        End If
    Next p
End Sub

' One folder's packet.txt, kept for this click. One that can't be read, or doesn't say whose
' packet it is, counts as none, so the folder is found by its name instead.
Private Sub ReadTag(ByVal fld As String)
    Dim key As String, sh As String, subj As String
    If Not ParseTag(fld, key, sh, subj) Then Exit Sub
    mNTag = mNTag + 1
    ReDim Preserve mTagDir(1 To mNTag): ReDim Preserve mTagKey(1 To mNTag)
    ReDim Preserve mTagShort(1 To mNTag): ReDim Preserve mTagSubject(1 To mNTag)
    mTagDir(mNTag) = fld
    mTagKey(mNTag) = key
    mTagShort(mNTag) = sh
    mTagSubject(mNTag) = subj
End Sub

' What a folder's packet.txt says: the packet it names, the Short Name and the subject. False
' when there's none, it can't be read, or it doesn't say whose packet it is.
Private Function ParseTag(ByVal fld As String, ByRef key As String, ByRef sh As String, ByRef subj As String) As Boolean
    Dim lines As Variant, i As Long, k As String, v As String, lic As String, tw As String, per As String
    key = "": sh = "": subj = ""
    If Not FileExists(fld & "\" & TAG_FILE) Then Exit Function
    On Error GoTo Unreadable
    lines = Split(Replace(ReadText(fld & "\" & TAG_FILE), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        If InStr(lines(i), ": ") > 0 Then
            k = LCase$(Trim$(Left$(lines(i), InStr(lines(i), ": ") - 1)))
            v = Trim$(Mid$(lines(i), InStr(lines(i), ": ") + 2))
            Select Case k
                Case "supplier": sh = v
                Case "license": lic = v
                Case "refund type": tw = v
                Case "period": per = v
                Case "subject": subj = v
            End Select
        End If
    Next i
    If Len(lic) = 0 Or Len(tw) = 0 Or PeriodKey(per) = 0 Then Exit Function
    key = PacketKey(PeriodKey(per), LicenseKey(lic), tw)
    ParseTag = True
    Exit Function
Unreadable:
    Close
End Function

' The folders the last version made (LIC-ACME) that claim a packet with no packet.txt yet,
' each after a |. A Short Name changed before this version leaves the old folder full and a new
' one beside it nearly empty, so a folder claims the packet when it holds an invoice for one of
' its BOLs or an MF-12 with its fingerprint. The one named for today's Short Name claims it too,
' unless it holds no files and another folder claims it; then it's left where it is.
' taken: the folders other packets claimed this click.
Private Function OldClaims(ByVal p As Long, ByVal taken As String) As String
    Dim dirs As Variant, i As Long, mf As String, fld As String, pre As String, own As String, out As String
    mf = MonthFolder(mPkt(p).PerKey)
    pre = UCase$(mPkt(p).Prefix & "-")
    dirs = ListFolders(mf)
    If IsEmpty(dirs) Then Exit Function
    For i = 1 To UBound(dirs)
        fld = mf & "\" & dirs(i)
        If UCase$(Left$(dirs(i), Len(pre))) = pre And TagAt(fld) = 0 And InStr(taken, "|" & UCase$(fld) & "|") = 0 Then
            If UCase$(dirs(i)) = pre & UCase$(mSup(mPkt(p).Sup).ShortName) Then
                own = fld
            ElseIf HoldsPacket(p, fld) Then
                out = out & "|" & fld
            End If
        End If
    Next i
    If Len(own) > 0 Then
        If HasFiles(own) Or Len(out) = 0 Then out = "|" & own & out
    End If
    OldClaims = out
End Function

' True when a folder holds an invoice for one of the packet's BOLs, or an MF-12 with one of its
' fingerprints.
Private Function HoldsPacket(ByVal p As Long, ByVal fld As String) As Boolean
    Dim files As Variant, i As Long, j As Long, r As Long, f As String, stem As String, ext As String, fp As String
    Dim bolNo As String, kind As String, sq As String
    files = ListFiles(fld)
    If IsEmpty(files) Then Exit Function
    For i = 1 To UBound(files)
        f = files(i)
        SplitExt f, stem, ext
        If InStr(UCase$(f), "MF-12") > 0 Then
            If InStr(UCase$(f), "NOTARIZED") = 0 And LCase$(ext) = ".pdf" Then
                fp = SafeFingerprint(fld & "\" & f)
                For j = 1 To mPkt(p).NForms
                    If Len(fp) > 0 And fp = FormFingerprint(p, j) Then HoldsPacket = True: Exit Function
                Next j
            End If
        ElseIf ParseFiledName(stem, bolNo, kind, sq) Then
            For r = 1 To mNRow
                If mRow(r).Pkt = p And mRow(r).BolK = BolKey(bolNo) Then HoldsPacket = True: Exit Function
            Next r
        End If
    Next i
End Function

' True when a folder has any file in it (folders inside it don't count).
Private Function HasFiles(ByVal fld As String) As Boolean
    Dim files As Variant, i As Long
    files = ListFiles(fld)
    If IsEmpty(files) Then Exit Function
    For i = 1 To UBound(files)
        If Not IsJunkFile(CStr(files(i))) Then HasFiles = True: Exit Function
    Next i
End Function

' Which packet.txt found this click is in this folder, 0 if none.
Private Function TagAt(ByVal fld As String) As Long
    Dim i As Long
    For i = 1 To mNTag
        If UCase$(mTagDir(i)) = UCase$(fld) Then TagAt = i: Exit Function
    Next i
End Function

' Renames one packet's folder to the name it should have now, then the files the macro made in
' it. A folder it can't rename (a file in it open in another program) stays as it is, with a
' line on CHECK, and the packet works from there this click. False only when the new name is
' still taken and this isn't the last try. A folder it can't look in holds only this packet.
Private Function RenamePacket(ByVal p As Long, ByVal lastTry As Boolean) As Boolean
    Dim cur As String, want As String, why As String, errNo As Long, was As String
    On Error GoTo Failed
    RenamePacket = True
    cur = mPkt(p).Folder
    want = MonthFolder(mPkt(p).PerKey) & "\" & mPkt(p).PName
    If Not FolderExists(cur) Then Exit Function         ' new: Build packets makes it
    was = LastPart(cur)
    If was <> mPkt(p).PName Then
        If FolderExists(want) And UCase$(cur) <> UCase$(want) Then
            If Not lastTry Then RenamePacket = False: Exit Function
            AddLookItem "A folder already has the new name, so this packet works under its old one:", _
                "Folders already have the new names, so these packets work under their old ones:", _
                PeriodLabel(mPkt(p).PerKey) & " " & was & " to " & mPkt(p).PName, "", _
                "Move the folder you don't want out of its month folder, then click Build packets again."
            Exit Function
        End If
        errNo = RenameOne(cur, want, why)
        If errNo <> 0 Then
            AddLookItem "Couldn't rename this packet's folder, so it works under its old name this click:", _
                "Couldn't rename these packets' folders, so each works under its old name this click:", _
                PeriodLabel(mPkt(p).PerKey) & " " & was & " to " & mPkt(p).PName, why, FixFor(errNo, "Build packets")
            Exit Function
        End If
        mPkt(p).Folder = want
    End If
    RenameFiles p, was
    Exit Function
Failed:
    HoldPacket p, "renaming its folder and files"
    RenamePacket = True
End Function

' The files the macro made in a packet's folder, under the names they should have now: each
' invoice by its BOL and kind, each MF-12 by the fingerprint inside it, a copy named after an
' MF-12 (the notarized one) to match, and UPLOAD as Upload. was is the folder's name before.
' Anything else in the folder is left alone.
Private Sub RenameFiles(ByVal p As Long, ByVal was As String)
    Dim files As Variant, i As Long, j As Long, f As String, stem As String, ext As String, nw As String
    Dim bolNo As String, kind As String, sq As String, why As String, fp As String, errNo As Long, fld As String

    fld = mPkt(p).Folder
    files = ListFiles(fld)
    If Not IsEmpty(files) Then
        For i = 1 To UBound(files)
            f = files(i): nw = ""
            SplitExt f, stem, ext
            If LCase$(f) = TAG_FILE Or IsJunkFile(f) Or LCase$(Right$(f, 8)) = ".partial" Then
                ' not the macro's to rename
            ElseIf InStr(UCase$(f), "MF-12") > 0 Then
                If InStr(UCase$(f), "NOTARIZED") = 0 And LCase$(ext) = ".pdf" And Not IsExpectedForm(p, f) Then
                    fp = SafeFingerprint(fld & "\" & f)
                    For j = 1 To mPkt(p).NForms
                        If Len(fp) > 0 And fp = FormFingerprint(p, j) Then nw = FormFileName(p, j)
                    Next j
                End If
                If Len(nw) = 0 Then nw = AfterForm(p, f, was)
            ElseIf ParseFiledName(stem, bolNo, kind, sq) Then
                nw = FiledName(p, bolNo, kind, sq)
                If Len(nw) > 0 Then nw = nw & ext
            End If
            If Len(nw) > 0 And nw <> f Then
                If FileExists(fld & "\" & nw) And UCase$(nw) <> UCase$(f) Then
                    AddLookItem "A file already has the new name, so this one keeps its old name:", _
                        "Files already have the new names, so these keep their old ones:", _
                        f & " (" & PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & ")", "", _
                        "Compare each with the file beside it and move the one you don't want out of the folder."
                Else
                    errNo = RenameOne(fld & "\" & f, fld & "\" & nw, why)
                    If errNo <> 0 Then RenameFailed f, p, why, errNo
                End If
            End If
        Next i
    End If

    files = ListFolders(fld)
    If Not IsEmpty(files) Then
        For i = 1 To UBound(files)
            If LCase$(files(i)) = "upload" And files(i) <> "Upload" Then
                errNo = RenameOne(fld & "\" & files(i), fld & "\Upload", why)
                If errNo <> 0 Then RenameFailed CStr(files(i)), p, why, errNo
            End If
        Next i
    End If
End Sub

' A file or folder in a packet that couldn't be renamed: one line on CHECK for all of them.
Private Sub RenameFailed(ByVal f As String, ByVal p As Long, ByVal why As String, ByVal errNo As Long)
    AddLookItem "Couldn't rename this, so it keeps its old name:", "Couldn't rename these, so they keep their old names:", _
        f & " (" & PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & ")", why, FixFor(errNo, "Build packets")
End Sub

' A file named after one of the packet's MF-12s under an older name, like "LIC-ACME AUGUST 2026
' MF-12 NOTARIZED.pdf": the same name with the MF-12's new name in front. Blank if it isn't one.
Private Function AfterForm(ByVal p As Long, ByVal f As String, ByVal was As String) As String
    Dim j As Long, k As Long, olds As Variant, s As String, c As String
    For j = 1 To mPkt(p).NForms
        olds = Array(FormStem(p, j, was, False), FormStem(p, j, was, True), _
                     FormStem(p, j, mPkt(p).Prefix & "-" & UCase$(mSup(mPkt(p).Sup).ShortName), True))
        For k = 0 To UBound(olds)
            s = olds(k)
            If Len(f) > Len(s) And UCase$(Left$(f, Len(s))) = UCase$(s) Then
                c = Mid$(f, Len(s) + 1, 1)
                If c = " " Or c = "." Then AfterForm = FormStem(p, j, mPkt(p).PName, False) & Mid$(f, Len(s) + 1): Exit Function
            End If
        Next k
    Next j
End Function

' Renames one file or folder and lists it for CHECK. 0 when it went, else the error number, and
' why says what stopped it.
Private Function RenameOne(ByVal src As String, ByVal dst As String, ByRef why As String) As Long
    On Error GoTo Failed
    Name src As dst
    If Not mRenamed Is Nothing Then mRenamed.Add RelPath(src) & "|" & LastPart(dst)
    Exit Function
Failed:
    RenameOne = Err.Number
    why = Err.Description
    Err.Clear
End Function

' ======================================================================================
'  Kept apart: a problem holds only its own packet or email
' ======================================================================================

' Holds a packet for this click after something went wrong with it. CHECK gets a line saying
' which packet, what it was doing and how to fix it, and the packet gets no MF-12 or draft
' this click. The rest go on. Called from an error catch, while Err still says what happened.
' Packets held the same way, with the same fix, share one line. Only the first stop on a packet
' is listed: it's the one to fix.
Private Sub HoldPacket(ByVal p As Long, ByVal doing As String)
    Dim why As String, errNo As Long
    why = Err.Description: errNo = Err.Number
    Err.Clear
    If StoppedOn(p) Then Exit Sub
    If Len(mPkt(p).Problem) = 0 Then mPkt(p).Problem = "stopped while " & doing
    AddLookItem "Held this click, it stopped partway through", "Held this click, it stopped partway through", _
        PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & " (" & doing & ": " & why & ")", "", FixFor(errNo, "Build packets")
End Sub

' The same for one email: no draft for it this click, and a line on CHECK.
Private Sub HoldMail(ByVal m As Long, ByVal doing As String)
    Dim why As String, errNo As Long
    why = Err.Description: errNo = Err.Number
    Err.Clear
    If Len(mMail(m).Held) = 0 Then mMail(m).Held = "stopped while " & doing
    AddLookItem "No draft made, it stopped partway through", "No drafts made, it stopped partway through", _
        PeriodLabel(mMail(m).PerKey) & " " & mSup(mMail(m).Sup).ShortName & " email (" & doing & ": " & why & ")", "", _
        FixFor(errNo, "Make drafts")
End Sub

' What to do about it, by the kind of error.
Private Function FixFor(ByVal errNo As Long, ByVal button As String) As String
    Select Case errNo
        Case 55, 70, 75       ' the file is open somewhere else, or it's locked
            FixFor = "Close the file in the program that has it open (a PDF viewer, say), then click " & button & " again."
        Case 53, 76           ' the file or folder isn't there
            FixFor = "Check the file or folder is still there, then click " & button & " again."
        Case 52, 68           ' a folder it can't look in, or a drive that isn't there
            FixFor = "Check the folder opens in File Explorer (its drive may not be connected, or you may need access to it), " & _
                "then click " & button & " again."
        Case 58               ' something already has that name
            FixFor = "There's already a file or folder with that name. Move the one you don't want, then click " & button & " again."
        Case vbObjectError + 12
            FixFor = "Shorten it on Corp License Number or Refunds Due so it fits the box, then click " & button & " again."
        Case Else
            FixFor = "Fix that, then click " & button & " again."
    End Select
End Function

' What the box says when it's done, ending with what to do next.
Private Function Finished(ByVal stepNo As Long, ByVal marked As Long) As String
    Dim msg As String, p As Long, m As Long, nSend As Long, nOut As Long, nBack As Long, nFiles As Long, nDraft As Long
    Dim nHeld As Long, made As Long

    For p = 1 To mNPkt
        Select Case mPkt(p).Stage
            Case ST_HELD: nHeld = nHeld + 1
            Case ST_FILES: nFiles = nFiles + 1
            Case ST_SEND: nSend = nSend + 1
            Case ST_DRAFT: nDraft = nDraft + 1
            Case ST_OUT: nOut = nOut + 1
            Case ST_BACK: nBack = nBack + 1
        End Select
    Next p

    If stepNo = 2 Then
        For m = 1 To mNMail
            If mMail(m).Drafted > 0 Then made = made + 1
        Next m
        msg = made & " draft(s) made. They're in Outlook Drafts. Nothing was sent."
    ElseIf stepNo = 3 Then
        msg = "Marked " & marked & " packet(s) filed today. They're off ENTRY."
    ElseIf mNPkt = 0 Then
        msg = "Nothing open on Refunds Due: every row has a Filed Date or a Refund Receive Date."
    Else
        msg = mNPkt & " packet(s) open."
    End If
    If mNPkt > 0 Then
        msg = msg & vbLf & vbLf
        If nHeld > 0 Then msg = msg & "  " & nHeld & " held until something's fixed (CHECK says what)" & vbLf
        If nFiles > 0 Then msg = msg & "  " & nFiles & " waiting on invoices" & vbLf
        If nSend > 0 Then msg = msg & "  " & nSend & " ready for Make drafts" & vbLf
        If nDraft > 0 Then msg = msg & "  " & nDraft & " in Outlook Drafts, not sent yet" & vbLf
        If nOut > 0 Then msg = msg & "  " & nOut & " sent, waiting on the notarized MF-12" & vbLf
        If nBack > 0 Then msg = msg & "  " & nBack & " back notarized, ready to key from ENTRY" & vbLf
    End If
    If mMoved > 0 Then msg = msg & vbLf & "Moved " & mMoved & " paid row(s) to Refunds Received."
    If mShortsFilled > 0 Then msg = msg & vbLf & "Filled in a Short Name for " & mShortsFilled & _
        " supplier(s) from the Corp name, on Corp License Number. It goes in folder and file names; change any you like."
    If mNLook > 0 Then msg = msg & vbLf & mNLook & " thing(s) need you: CHECK lists them under Needs you."

    msg = msg & vbLf & vbLf & "Next: "
    If nFiles > 0 Then
        msg = msg & "save the invoices into INBOX as ""12345 VI"" and ""12345 CI"" (12345 is the BOL), " & _
            "then click Build packets again. INBOX is " & InboxFolder() & "."
    ElseIf nSend > 0 Then
        msg = msg & "click Make drafts."
    ElseIf nDraft > 0 Then
        msg = msg & "look over the drafts in Outlook and send them."
    ElseIf nBack > 0 Then
        msg = msg & "key the ENTRY sheet into the Georgia portal, then click Mark filed."
    ElseIf nOut > 0 Then
        msg = msg & "nothing until a notarized MF-12 comes back. Save it in its packet folder with NOTARIZED at the end of the name."
    ElseIf nHeld > 0 Then
        msg = msg & "fix what CHECK lists, then click Build packets again."
    Else
        msg = msg & "paste the next refund report's rows at the bottom of Refunds Due."
    End If
    Finished = msg
End Function


' ======================================================================================
'  The first click
' ======================================================================================

' The first click's questions, each asked only while its SETUP cell is blank: the template
' (found by itself if it can be), the company on the MF-12, the packets folder, and the oldest
' period to work. False if a question was cancelled or the template can't be found.
Private Function AskSetup() As Boolean
    Dim ws As Worksheet, v As String, q As Variant, f As String, pers As Collection, i As Long, lst As String

    Set ws = SheetNamed(SETUP_SHEET)

    f = SetupValue(ws, "mf-12 template")
    If Not FileExists(f) Then
        f = FindTemplate()
        If Len(f) = 0 Then
            Say "Can't find " & TEMPLATE_NAME & ". It comes with this macro: save it in the same folder as this workbook," & _
                vbLf & mBook.Path & vbLf & "then click Build packets again.", vbExclamation
            Exit Function
        End If
        SetSetup ws, "mf-12 template", f
    End If

    For Each q In Array(Array("purchaser name", "The company's name, the way it goes on the MF-12 as the purchaser:"), _
                        Array("purchaser street", "The company's street address, for the MF-12:"), _
                        Array("purchaser city", "City:"), _
                        Array("purchaser state", "State, 2 letters:"), _
                        Array("purchaser zip", "ZIP code:"), _
                        Array("company short name", "A short name for the company, for the email subject line " & _
                              "(the name suppliers know you by):"))
        If Len(SetupValue(ws, CStr(q(0)))) = 0 Then
            v = Ask(CStr(q(1)) & vbLf & vbLf & "Asked once. It goes on the SETUP tab, where you can change it.", "")
            If Len(v) = 0 Then
                Say "Stopped. Click Build packets again to pick up where you left off.", vbInformation
                Exit Function
            End If
            If CStr(q(0)) = "purchaser state" Then v = UCase$(v)
            SetSetup ws, CStr(q(0)), v
        End If
    Next q

    If Len(SetupValue(ws, "packets folder")) = 0 Then
        Say "Next, pick the folder the GA refund packets go in. Each period gets its own folder inside it " & _
            "(like 2026-08 August), plus an INBOX for the invoices. The macro makes them all.", vbInformation
        v = AskFolder("Pick the folder the GA refund packets go in")
        If Len(v) = 0 Then
            Say "Stopped. Click Build packets again and pick a folder.", vbInformation
            Exit Function
        End If
        If Right$(v, 1) = "\" Then v = Left$(v, Len(v) - 1)
        SetSetup ws, "packets folder", v
    End If

    If Len(PeriodText(ws.Cells(SetupRow(ws, "oldest period"), 2).Value)) = 0 Then
        Set pers = OpenPeriods()
        If pers.Count <= 1 Then
            v = "none"
        Else
            For i = 1 To pers.Count
                If i <= 24 Then lst = lst & IIf(Len(lst) > 0, ", ", "") & PeriodLabel(pers(i))
            Next i
            If pers.Count > 24 Then lst = lst & " and " & (pers.Count - 24) & " more"
            v = Ask("Refunds Due has rows with no Filed Date from these periods:" & vbLf & lst & vbLf & vbLf & _
                "Which is the oldest period to make packets for? Rows from before it are left alone, so a refund " & _
                "you already filed by hand doesn't get a second packet. Type none to take every row." & vbLf & vbLf & _
                "Asked once. It goes on the SETUP tab, where you can change it.", PeriodLabel(pers(pers.Count)))
            If Len(v) = 0 Then
                Say "Stopped. Click Build packets again to pick up where you left off.", vbInformation
                Exit Function
            End If
        End If
        SetSetup ws, "oldest period", v
    End If
    AskSetup = True
End Function

' mf12_template.pdf beside this workbook, in Downloads, or in the packets folder.
Private Function FindTemplate() As String
    Dim c As Variant
    For Each c In Array(mBook.Path, Environ$("USERPROFILE") & "\Downloads", SetupValue(SheetNamed(SETUP_SHEET), "packets folder"))
        If Len(c) > 0 Then
            If FileExists(c & "\" & TEMPLATE_NAME) Then FindTemplate = c & "\" & TEMPLATE_NAME: Exit Function
        End If
    Next c
End Function

Private Function LoadTemplate() As Boolean
    If Not FileExists(mTemplatePath) Then Exit Function
    mTemplate = ReadBytes(mTemplatePath)
    If LenB(mTemplate) < 1000 Then mTemplate = "": Exit Function
    LoadTemplate = InStrB(1, mTemplate, AnsiB("(MF12V:SELLER:"), vbBinaryCompare) > 0 And _
                   InStrB(1, mTemplate, AnsiB("%MF12A:R21.FUEL:"), vbBinaryCompare) > 0 And _
                   InStrB(1, mTemplate, AnsiB("%MF12ID:"), vbBinaryCompare) > 0
End Function


' ======================================================================================
'  Folders and the INBOX
' ======================================================================================

' A period's folder under the packets folder, spelled the way it is on disk whatever its capitals.
Private Function MonthFolder(ByVal perKey As Long) As String
    Dim nm As String, i As Long
    nm = MonthFolderName(perKey)
    If Not IsEmpty(mMonthDirs) Then
        For i = 1 To UBound(mMonthDirs)
            If UCase$(mMonthDirs(i)) = UCase$(nm) Then nm = mMonthDirs(i): Exit For
        Next i
    End If
    MonthFolder = mBase & "\" & nm
End Function

' 2026-08 August: still sorts by date.
Private Function MonthFolderName(ByVal perKey As Long) As String
    MonthFolderName = (perKey \ 100) & "-" & Format$(perKey Mod 100, "00") & " " & MonthName(perKey Mod 100)
End Function

' A filed invoice's name, without .pdf: Acme BOL 123456 Invoice 7890 VI, and a 2 on the end
' of a second copy. kind is VI or CI.
Private Function InvoiceName(ByVal p As Long, ByVal bolShown As String, ByVal invs As String, ByVal kind As String, _
    ByVal sq As String) As String
    InvoiceName = mSup(mPkt(p).Sup).ShortName & " BOL " & SafeName(bolShown) & " Invoice " & SafeName(invs) & " " & kind & _
        IIf(Len(sq) > 0, " " & sq, "")
End Function

' What a filed invoice of this packet should be called now, from the BOL in its name; blank if
' no row of the packet has that BOL.
Private Function FiledName(ByVal p As Long, ByVal bolNo As String, ByVal kind As String, ByVal sq As String) As String
    Dim invs As String, shown As String
    invs = BolInvoices(BolKey(bolNo), p, shown)
    If Len(shown) > 0 Then FiledName = InvoiceName(p, shown, invs, kind, sq)
End Function

' The BOL and kind of copy in a filed invoice's name, the way this version names it (Acme BOL
' 123456 Invoice 7890 VI 2) or the last one did (LIC-ACME 123456 INV7890 VI 2). False if
' it's neither.
Private Function ParseFiledName(ByVal stem As String, ByRef bolNo As String, ByRef kind As String, ByRef sq As String) As Boolean
    Dim t As Variant, i As Long, k As Long
    bolNo = "": kind = "": sq = ""
    stem = Trim$(stem)
    Do While InStr(stem, "  ") > 0: stem = Replace(stem, "  ", " "): Loop
    t = Split(stem, " ")
    k = UBound(t)
    If k >= 0 Then
        If IsAllDigits(CStr(t(k))) Then sq = t(k): k = k - 1
    End If
    If k < 2 Then Exit Function
    If Not IsAllLetters(CStr(t(k))) Then Exit Function
    kind = UCase$(t(k))
    For i = k - 2 To 2 Step -1
        If UCase$(t(i)) = "INVOICE" And UCase$(t(i - 2)) = "BOL" Then bolNo = t(i - 1): Exit For
    Next i
    If Len(bolNo) = 0 And k >= 3 Then
        If UCase$(Left$(t(k - 1), 3)) = "INV" Then bolNo = t(k - 2)
    End If
    If Len(bolNo) = 0 Then sq = "": kind = ""
    ParseFiledName = (Len(bolNo) > 0)
End Function

Private Function InboxFolder() As String
    InboxFolder = mBase & "\INBOX"
End Function

Private Sub MakeFolders()
    Dim p As Long
    MkDirs mBase
    MkDirs InboxFolder()
    For p = 1 To mNPkt
        MakePacketFolders p
    Next p
End Sub

Private Sub MakePacketFolders(ByVal p As Long)
    On Error GoTo Failed
    If Len(mPkt(p).Folder) = 0 Then Exit Sub         ' held before it has a folder
    MkDirs MonthFolder(mPkt(p).PerKey) & "\MFD-33"
    MkDirs mPkt(p).Folder
    MkDirs mPkt(p).Folder & "\Upload"
    Exit Sub
Failed:
    HoldPacket p, "making its folder " & mPkt(p).Folder
End Sub

' Files each "<BOL> VI" / "<BOL> CI" into its packet folder under its full name, whatever period
' it's from. fileThem False (Make drafts, Mark filed) only says on CHECK what Build packets would
' say about each file it can't place, and moves nothing.
Private Sub SortInbox(ByVal fileThem As Boolean)
    Dim files As Variant, i As Long

    files = ListFiles(InboxFolder())
    If IsEmpty(files) Then Exit Sub
    For i = 1 To UBound(files)
        If Not IsJunkFile(CStr(files(i))) Then FileFromInbox CStr(files(i)), fileThem
    Next i
End Sub

' One INBOX file into its packet. A file that won't move (open in another program, say)
' stays in INBOX and holds only its own packet.
Private Sub FileFromInbox(ByVal f As String, ByVal fileIt As Boolean)
    Dim stem As String, ext As String, bolPart As String, kind As String, sq As String
    Dim r As Long, pk As Long, twoPkts As Boolean, waiting As Long, invs As String, bolShown As String, target As String

    On Error GoTo Failed
    SplitExt f, stem, ext
    If Not ParseInboxName(stem, bolPart, kind, sq) Then
        AddLookItem "INBOX: can't read the name", "INBOX: can't read the names", f, "Left in INBOX.", _
            "Name each one like 12345 VI or 12345 CI (12345 is the BOL)."
    Else
        pk = 0: twoPkts = False: waiting = 0
        For r = 1 To mNRow
            If mRow(r).BolK = BolKey(bolPart) Then
                If mRow(r).Pkt = 0 Then
                    waiting = mRow(r).SheetRow
                ElseIf pk = 0 Then
                    pk = mRow(r).Pkt
                ElseIf pk <> mRow(r).Pkt Then
                    twoPkts = True
                End If
            End If
        Next r
        invs = BolInvoices(BolKey(bolPart), 0, bolShown)
        If Len(bolShown) = 0 Then
            AddLookItem "INBOX: no open row has the BOL of", "INBOX: no open row has the BOLs of", f, "Left in INBOX.", _
                "Check the BOL. Its row may be filed already, or from before the Oldest period to work on SETUP."
        ElseIf waiting > 0 Then
            AddLookItem "INBOX: waiting on a row that needs a fix first:", "INBOX: waiting on rows that need a fix first:", _
                f & " (row " & waiting & ")", "Left in INBOX.", "Fix the row (its line is above), then click Build packets again."
        ElseIf twoPkts Then
            AddLook "INBOX: BOL " & bolPart & " is on rows in two packets", f & " stays in INBOX.", "File this one by hand."
        ElseIf Len(mPkt(pk).Folder) = 0 Then
            AddLookItem "INBOX: waiting on its packet, which is held:", "INBOX: waiting on their packets, which are held:", _
                f & " (" & mPkt(pk).PName & ")", "Left in INBOX.", "Fix what holds the packet (its line is above), then click Build packets again."
        Else
            target = mPkt(pk).Folder & "\" & InvoiceName(pk, bolShown, invs, kind, sq) & ext
            If FileExists(target) Or Not FolderExists(mPkt(pk).Folder) Then
                AddLookItem "INBOX: already filed, so the INBOX copy was left alone:", _
                    "INBOX: already filed, so the INBOX copies were left alone:", _
                    f & " (as " & Mid$(target, InStrRev(target, "\") + 1) & ")", "", "Compare the two and delete the one you don't want."
            ElseIf fileIt Then
                Name InboxFolder() & "\" & f As target
            End If
        End If
    End If
    Exit Sub
Failed:
    If pk > 0 Then
        HoldPacket pk, "filing " & f & " from INBOX"
    Else
        AddLookItem "INBOX: it stopped while filing", "INBOX: it stopped while filing", f & " (" & Err.Description & ")", "", _
            FixFor(Err.Number, "Build packets")
    End If
End Sub

' "12345 VI", "12345 VI 2", "12345VI". The BOL, the kind of copy, and an optional number.
Private Function ParseInboxName(ByVal stem As String, ByRef bolPart As String, ByRef kind As String, ByRef sq As String) As Boolean
    Dim t As Variant, s As String
    s = Trim$(Replace(stem, "_", " "))
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    t = Split(s, " ")
    bolPart = "": kind = "": sq = ""
    Select Case UBound(t)
        Case 0
            If Len(s) > 2 And (UCase$(Right$(s, 2)) = "VI" Or UCase$(Right$(s, 2)) = "CI") Then
                bolPart = Left$(s, Len(s) - 2): kind = UCase$(Right$(s, 2))
            End If
        Case 1
            bolPart = t(0): kind = UCase$(t(1))
        Case 2
            If IsAllDigits(CStr(t(2))) Then bolPart = t(0): kind = UCase$(t(1)): sq = t(2)
    End Select
    If Len(bolPart) = 0 Or Len(kind) = 0 Or Len(kind) > 12 Then Exit Function
    If Not IsAllLetters(kind) Then Exit Function
    ParseInboxName = True
End Function


' ======================================================================================
'  What's in each packet folder
' ======================================================================================

Private Sub CheckFiles()
    Dim p As Long
    For p = 1 To mNPkt
        CheckPacket p
    Next p
End Sub

' What's in one packet's folder: its packet.txt, which VIs are there (by the BOL in their names),
' the newest notarized copy, and whether its MF-12s are there (by the fingerprint inside them)
' and still match the rows. A file it can't read holds only this packet.
Private Sub CheckPacket(ByVal p As Long)
    Dim r As Long, files As Variant, i As Long, j As Long, stem As String, ext As String, doing As String
    Dim f As String, stray As String, fp As String, onDisk As Long, bad As String, vis As String
    Dim bolNo As String, kind As String, sq As String, have() As Boolean, hit As Boolean, off As Boolean
    Dim d As Date, printed As Double, owed As Double

    On Error GoTo Failed
    mPkt(p).VIMissing = ""
    mPkt(p).Notarized = 0
    If FolderExists(mPkt(p).Folder) Then
        doing = "writing its " & TAG_FILE
        WriteTag p
    End If
    doing = "looking in its folder"
    files = ListFiles(mPkt(p).Folder)
    vis = "|"
    If Not IsEmpty(files) Then
        For i = 1 To UBound(files)
            SplitExt CStr(files(i)), stem, ext
            If InStr(UCase$(files(i)), "MF-12") = 0 Then
                If ParseFiledName(stem, bolNo, kind, sq) Then
                    If kind = "VI" Then vis = vis & BolKey(bolNo) & "|"
                End If
            End If
        Next i
    End If
    For r = 1 To mNRow
        If mRow(r).Pkt = p Then
            mRow(r).HasVI = (InStr(vis, "|" & mRow(r).BolK & "|") > 0)
            If Not mRow(r).HasVI Then
                If InStr(", " & mPkt(p).VIMissing & ",", ", " & mRow(r).Bol & ",") = 0 Then
                    mPkt(p).VIMissing = mPkt(p).VIMissing & IIf(Len(mPkt(p).VIMissing) > 0, ", ", "") & mRow(r).Bol
                End If
            End If
        End If
    Next r
    ' a VI not in yet is a wait, not a problem: CHECK shows it on the packet and its email

    ' the notarized copies that have come back: the newest one's date
    If Not IsEmpty(files) Then
        For i = 1 To UBound(files)
            If InStr(UCase$(CStr(files(i))), "NOTARIZED") > 0 Then
                doing = "reading " & files(i)
                d = FileDateTime(mPkt(p).Folder & "\" & files(i))
                If d > mPkt(p).Notarized Then mPkt(p).Notarized = d
            End If
        Next i
    End If

    ' the MF-12s, each found by the fingerprint inside it, whatever it's called. One with the
    ' name a form should have but not its fingerprint is out of date; any other is in the way.
    ' Each one found has its printed tax added up and held against its rows.
    onDisk = 0: bad = "": stray = "": off = False
    ReDim have(1 To mPkt(p).NForms)
    If Not IsEmpty(files) Then
        For i = 1 To UBound(files)
            f = files(i)
            If InStr(UCase$(f), "MF-12") > 0 And InStr(UCase$(f), "NOTARIZED") = 0 And LCase$(Right$(f, 4)) = ".pdf" Then
                doing = "reading " & f
                fp = FileFingerprint(mPkt(p).Folder & "\" & f)
                hit = False
                For j = 1 To mPkt(p).NForms
                    If Not have(j) And fp = FormFingerprint(p, j) Then have(j) = True: hit = True: Exit For
                Next j
                If hit Then
                    onDisk = onDisk + 1
                    doing = "adding up the tax on " & f
                    printed = PrintedTax(mPkt(p).Folder & "\" & f)
                    owed = FormTax(p, j)
                    If Abs(printed - owed) > 0.005 Then
                        off = True
                        AddLook PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & ": the tax printed on " & f & " adds up to " & _
                            Format$(printed, "#,##0.00") & ", but its rows add up to " & Format$(owed, "#,##0.00"), "", _
                            "It was changed after it was made. Move it out of the folder and click Build packets to make a fresh one."
                    End If
                ElseIf IsExpectedForm(p, f) Then
                    onDisk = onDisk + 1
                    bad = bad & IIf(Len(bad) > 0, ", ", "") & f
                Else
                    stray = stray & IIf(Len(stray) > 0, ", ", "") & f
                End If
            End If
        Next i
    End If
    If Len(stray) > 0 Then
        mPkt(p).FormState = "OLD MF-12 IN FOLDER"
        AddLook mPkt(p).PName & ": an MF-12 in the folder doesn't fit this run", stray, _
            "The number of forms changed, or it's an old copy. Move it out of the folder and click Build packets."
    ElseIf Len(bad) > 0 Then
        mPkt(p).FormState = "MF-12 OUT OF DATE"
        AddLookItem "This MF-12 doesn't match the rows anymore:", "These MF-12s don't match their rows anymore:", _
            PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & " (" & bad & ")", _
            "The report changed since it was made, or it was edited by hand.", _
            "Move it out of its folder and click Build packets to make a fresh one."
    ElseIf off Then
        mPkt(p).FormState = "TAX DOESN'T ADD UP"
    ElseIf onDisk = mPkt(p).NForms Then
        mPkt(p).FormState = "KEPT"
    ElseIf onDisk = 0 Then
        mPkt(p).FormState = "MISSING"
    Else
        mPkt(p).FormState = "SOME FORMS MISSING"
        AddLook mPkt(p).PName & ": only " & onDisk & " of " & mPkt(p).NForms & " MF-12 forms are there", "The " & mPkt(p).PName & " folder", _
            "Move the MF-12s out of the folder and click Build packets to make them all fresh."
    End If
    Exit Sub
Failed:
    HoldPacket p, doing
    Close
End Sub

Private Function IsExpectedForm(ByVal p As Long, ByVal f As String) As Boolean
    Dim j As Long
    For j = 1 To mPkt(p).NForms
        If UCase$(f) = UCase$(FormFileName(p, j)) Then IsExpectedForm = True: Exit Function
    Next j
End Function

' Acme Licensed MF-12 August 2026 (1 of 2).pdf, with no (1 of 1) when there's one form.
Private Function FormFileName(ByVal p As Long, ByVal formNo As Long) As String
    FormFileName = FormStem(p, formNo, mPkt(p).PName, False) & ".pdf"
End Function

' An MF-12's name without .pdf, for a packet folder called lbl. lastWay gives the name the last
' version used, the month in capitals after the folder's name: LIC-ACME AUGUST 2026 MF-12 (1 of 2).
Private Function FormStem(ByVal p As Long, ByVal formNo As Long, ByVal lbl As String, ByVal lastWay As Boolean) As String
    Dim mo As String, ofN As String
    mo = MonthName(mPkt(p).PerKey Mod 100) & " " & (mPkt(p).PerKey \ 100)
    If mPkt(p).NForms > 1 Then ofN = " (" & formNo & " of " & mPkt(p).NForms & ")"
    If lastWay Then
        FormStem = lbl & " " & UCase$(mo) & " MF-12" & ofN
    Else
        FormStem = lbl & " MF-12 " & mo & ofN
    End If
End Function

' The fingerprint inside a file, blank if it can't be read.
Private Function SafeFingerprint(ByVal path As String) As String
    On Error Resume Next
    SafeFingerprint = FileFingerprint(path)
    If Err.Number <> 0 Then SafeFingerprint = ""
    Close
    On Error GoTo 0
End Function

' The tax an MF-12 shows, added up from its Tax Paid boxes.
Private Function PrintedTax(ByVal path As String) As Double
    Dim s As String, pos As Long, v As Long, e As Long, box As String
    s = ReadBytes(path)
    If LenB(s) = 0 Then Exit Function
    box = AnsiB("/TU (Enter Tax Paid To Seller By Purchaser)")
    pos = InStrB(1, s, box, vbBinaryCompare)
    Do While pos > 0
        v = InStrB(pos, s, AnsiB("/V ("), vbBinaryCompare)
        If v = 0 Then Exit Do
        e = InStrB(v + 4, s, AnsiB(")"), vbBinaryCompare)
        If e = 0 Then Exit Do
        PrintedTax = PrintedTax + Val(StrConv(MidB$(s, v + 4, e - v - 4), vbUnicode))
        pos = InStrB(e, s, box, vbBinaryCompare)
    Loop
End Function

' What the tax on one of a packet's MF-12s should add up to, from its rows.
Private Function FormTax(ByVal p As Long, ByVal formNo As Long) As Double
    Dim ix As Variant, i As Long
    ix = PacketRows(p)
    For i = (formNo - 1) * ROWS_PER_FORM + 1 To formNo * ROWS_PER_FORM
        If i > UBound(ix) Then Exit For
        FormTax = FormTax + Val(PdfNum(mRow(ix(i)).Amount, 2))
    Next i
End Function


' ======================================================================================
'  The MF-12
' ======================================================================================

' Every MF-12 not made yet is filled; saveThem saves them (Build packets), else it's only a
' check that they can be filled.
Private Sub WriteForms(ByVal saveThem As Boolean)
    Dim p As Long
    For p = 1 To mNPkt
        If Len(mPkt(p).Problem) = 0 And mPkt(p).FormState = "MISSING" And FolderExists(mPkt(p).Folder) Then WritePacketForms p, saveThem
    Next p
End Sub

' One packet's MF-12s. Every form is filled before any is saved, so a value too long for its
' box leaves none of them half made, and holds only this packet.
Private Sub WritePacketForms(ByVal p As Long, ByVal saveThem As Boolean)
    Dim j As Long, forms() As String, doing As String
    On Error GoTo Failed
    doing = "filling its MF-12"
    ReDim forms(1 To mPkt(p).NForms)
    For j = 1 To mPkt(p).NForms
        forms(j) = FillForm(p, j)
    Next j
    If Not saveThem Then Exit Sub
    For j = 1 To mPkt(p).NForms
        doing = "saving " & FormFileName(p, j)
        WriteBytes mPkt(p).Folder & "\" & FormFileName(p, j), forms(j)
    Next j
    mPkt(p).FormState = "MADE"
    Exit Sub
Failed:
    HoldPacket p, doing
    Close
End Sub

Private Function FillForm(ByVal p As Long, ByVal formNo As Long) As String
    Dim s As String, v As Object, pos As Long, k As Long, e As Long, n As Long, key As String, rep As String
    Dim w As Double, h As Double, q As Long

    Set v = FormValues(p, formNo)
    s = mTemplate

    ' value slots: (MF12V:KEY:____)
    pos = InStrB(1, s, AnsiB("(MF12V:"), vbBinaryCompare)
    Do While pos > 0
        k = InStrB(pos + 7, s, AnsiB(":"), vbBinaryCompare)
        key = StrConv(MidB$(s, pos + 7, k - pos - 7), vbUnicode)
        e = k + 1
        Do While AscB(MidB$(s, e, 1)) = 95
            e = e + 1
        Loop
        n = e - pos + 1
        rep = AnsiB("(") & PdfText(SlotText(v, key, 0)) & AnsiB(")")
        If LenB(rep) > n Then Err.Raise vbObjectError + 12, , "Too long to fit the MF-12 box " & key & ": " & SlotText(v, key, 0)
        MidB(s, pos, n) = rep & AnsiB(Space$(n - LenB(rep)))
        pos = InStrB(pos + n, s, AnsiB("(MF12V:"), vbBinaryCompare)
    Loop

    ' drawing slots: %MF12A:KEY:W=000.00:H=000.00:Q=0:____
    pos = InStrB(1, s, AnsiB("%MF12A:"), vbBinaryCompare)
    Do While pos > 0
        k = InStrB(pos + 7, s, AnsiB(":"), vbBinaryCompare)
        key = StrConv(MidB$(s, pos + 7, k - pos - 7), vbUnicode)
        w = Val(StrConv(MidB$(s, k + 3, 6), vbUnicode))
        h = Val(StrConv(MidB$(s, k + 12, 6), vbUnicode))
        q = Val(StrConv(MidB$(s, k + 21, 1), vbUnicode))
        e = k + 23
        Do While AscB(MidB$(s, e, 1)) = 95
            e = e + 1
        Loop
        n = e - pos
        rep = Drawing(SlotText(v, key, 1), w, h, q)
        If LenB(rep) > n Then Err.Raise vbObjectError + 12, , "Too long to fit the MF-12 box " & key & ": " & SlotText(v, key, 1)
        MidB(s, pos, n) = rep & AnsiB(Space$(n - LenB(rep)))
        pos = InStrB(pos + n, s, AnsiB("%MF12A:"), vbBinaryCompare)
    Loop

    ' the fingerprint
    pos = InStrB(1, s, AnsiB("%MF12ID:"), vbBinaryCompare)
    e = pos + 8
    Do While AscB(MidB$(s, e, 1)) = 95
        e = e + 1
    Loop
    rep = AnsiB("%MF12ID:" & FormFingerprint(p, formNo))
    MidB(s, pos, e - pos) = rep & AnsiB(Space$(e - pos - LenB(rep)))
    FillForm = s
End Function

Private Function SlotText(v As Object, ByVal key As String, ByVal which As Long) As String
    If v.Exists(key) Then SlotText = v(key)(which)
End Function

' What the box shows: Courier Bold 10 like the form asks, shrunk only if it won't fit.
Private Function Drawing(ByVal shown As String, ByVal w As Double, ByVal h As Double, ByVal q As Long) As String
    Dim size As Double, tw As Double, x As Double, y As Double
    If Len(shown) = 0 Then Exit Function
    size = 10
    tw = Len(shown) * 0.6 * size
    If tw > w - 4 Then
        size = Int((w - 4) / (Len(shown) * 0.6) * 100) / 100
        tw = Len(shown) * 0.6 * size
    End If
    Select Case q
        Case 1: x = (w - tw) / 2
        Case 2: x = w - 2 - tw
        Case Else: x = 2
    End Select
    y = (h - 0.562 * size) / 2
    Drawing = AnsiB("BT /CoBo " & PdfNum(size, 2) & " Tf 0 g " & PdfNum(x, 2) & " " & PdfNum(y, 2) & " Td (") & _
              PdfText(shown) & AnsiB(") Tj ET")
End Function

' Text as PDF bytes: parentheses and backslashes escaped, curly quotes made plain.
Private Function PdfText(ByVal t As String) As String
    Dim i As Long, c As Long, out As String, ch As String
    t = Replace(Replace(t, ChrW(8217), "'"), ChrW(8216), "'")
    t = Replace(Replace(t, ChrW(8220), """"), ChrW(8221), """")
    t = Replace(Replace(t, ChrW(8211), "-"), ChrW(8212), "-")
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        If c = 40 Or c = 41 Or c = 92 Then
            out = out & ChrB(92) & ChrB(c)
        ElseIf c >= 32 And c < 127 Then
            out = out & ChrB(c)
        Else
            c = Asc(ch)
            If c < 32 Or c > 255 Then c = 63
            out = out & ChrB(c)
        End If
    Next i
    PdfText = out
End Function

Private Function FileFingerprint(ByVal path As String) As String
    Dim s As String, pos As Long, e As Long, c As Long
    s = ReadBytes(path)
    If LenB(s) = 0 Then Exit Function
    pos = InStrB(1, s, AnsiB("%MF12ID:"), vbBinaryCompare)
    If pos = 0 Then FileFingerprint = "?": Exit Function
    e = pos + 8
    Do While e <= LenB(s)
        c = AscB(MidB$(s, e, 1))
        If c = 32 Or c = 10 Or c = 13 Then Exit Do
        e = e + 1
    Loop
    FileFingerprint = StrConv(MidB$(s, pos + 8, e - pos - 8), vbUnicode)
End Function

Private Sub WriteBytes(ByVal path As String, ByVal s As String)
    Dim b() As Byte, f As Integer, tmp As String
    If FileExists(path) Then Err.Raise vbObjectError + 13, , "Won't write over " & path
    b = s
    tmp = path & ".partial"
    If FileExists(tmp) Then Kill tmp     ' only ever this macro's own half-written file
    f = FreeFile
    Open tmp For Binary Access Write As #f
    Put #f, , b
    Close #f
    Name tmp As path
End Sub


' ======================================================================================
'  What's been sent, who's ready, and the drafts
' ======================================================================================

' An email has gone once a row of it has MF-12 Sent, or its subject is in Outlook's Sent
' Items. Outlook is opened for Make drafts; any other click only looks if it's already open.
' A draft already waiting in Drafts is noted too, so Make drafts never makes a second one.
' The subjects looked for are the ones its packet.txt files say its draft went under, then
' today's, then today's with the Short Name in capitals the way the last version made it, so a
' Short Name change after the email went doesn't lose it.
Private Sub CheckSent(ByVal openIt As Boolean)
    Dim ol As Object, ns As Object, sentBox As Object, draftBox As Object, m As Long, subj As Variant, found As Object
    If mNMail = 0 Then Exit Sub
    If Len(mTestDraftDir) = 0 Then
        On Error Resume Next
        Set ol = GetObject(, "Outlook.Application")
        If ol Is Nothing And openIt Then Set ol = CreateObject("Outlook.Application")
        If Not ol Is Nothing Then
            Set ns = ol.GetNamespace("MAPI")
            Set sentBox = ns.GetDefaultFolder(5)
            Set draftBox = ns.GetDefaultFolder(16)
        End If
        On Error GoTo 0
        If sentBox Is Nothing Then Exit Sub
    End If
    For m = 1 To mNMail
        For Each subj In Subjects(m)
            If Not mMail(m).IsSent Then
                If Len(mTestDraftDir) > 0 Then
                    If FileExists(mTestDraftDir & "\sent\" & SafeName(CStr(subj)) & ".txt") Then
                        mMail(m).IsSent = True
                        mMail(m).SentOn = FileDateTime(mTestDraftDir & "\sent\" & SafeName(CStr(subj)) & ".txt")
                    End If
                Else
                    Set found = FindMail(sentBox, CStr(subj))
                    If Not found Is Nothing Then
                        mMail(m).IsSent = True
                        mMail(m).SentOn = found.SentOn
                    End If
                End If
            End If
        Next subj
        If Not mMail(m).IsSent Then
            If Len(mTestDraftDir) > 0 Then
                mMail(m).InDrafts = FileExists(TestDraftPath(m))
            Else
                For Each subj In Subjects(m)
                    If Not FindMail(draftBox, CStr(subj)) Is Nothing Then mMail(m).InDrafts = True
                Next subj
            End If
        End If
    Next m
End Sub

' The subjects one email may have gone under, each once.
Private Function Subjects(ByVal m As Long) As Collection
    Dim c As New Collection, p As Long, s As Variant, have As String, pool As String
    For p = 1 To mNPkt
        If mPkt(p).Mail = m And Len(mPkt(p).Subject) > 0 Then pool = pool & mPkt(p).Subject & vbLf
    Next p
    pool = pool & FillWords(mSubject, m) & vbLf & FillWords(mSubject, m, UCase$(mSup(mMail(m).Sup).ShortName))
    For Each s In Split(pool, vbLf)
        If Len(s) > 0 And InStr(have, "|" & s & "|") = 0 Then
            have = have & "|" & s & "|"
            c.Add s
        End If
    Next s
    Set Subjects = c
End Function

Private Sub DecideReady()
    Dim m As Long, s As Long, p As Long, why As String
    ' a notarized copy back means the email went, even if nothing says when
    For p = 1 To mNPkt
        If mPkt(p).Notarized > 0 Then mMail(mPkt(p).Mail).IsSent = True
    Next p
    If mNMail > 0 Then ReDim mWaitOnly(1 To mNMail)
    For m = 1 To mNMail
        s = mMail(m).Sup
        why = mMail(m).Held
        If Len(why) = 0 And Len(mSup(s).Problem) > 0 Then why = mSup(s).Problem
        If Len(why) = 0 And Len(mSup(s).Email) + Len(mSup(s).GroupMail) + Len(mSup(s).Extra) = 0 Then
            why = "no email address on " & CONTACTS_SHEET
            AddLookItem "No email address on " & CONTACTS_SHEET & " for", "No email address on " & CONTACTS_SHEET & " for", _
                mSup(s).ShortName & " (row " & mSup(s).SheetRow & ")", "", "Fill in Email address or Group Email Address."
        End If
        For p = 1 To mNPkt
            If mPkt(p).Mail = m And Len(why) = 0 Then
                If Len(mPkt(p).Problem) > 0 Then
                    why = mPkt(p).PName & ": " & mPkt(p).Problem
                ElseIf Len(mPkt(p).VIMissing) > 0 Then
                    why = mPkt(p).PName & ": waiting on VI for BOL " & mPkt(p).VIMissing
                    mWaitOnly(m) = True
                ElseIf mPkt(p).FormState <> "MADE" And mPkt(p).FormState <> "KEPT" Then
                    why = mPkt(p).PName & ": MF-12 " & LCase$(mPkt(p).FormState)
                    mWaitOnly(m) = (mPkt(p).FormState = "MISSING")
                End If
            End If
        Next p
        mMail(m).Held = why
        ' ready: list what it would attach, so the letter's counts and the size are known
        If Not mMail(m).IsSent And Len(why) = 0 And Not mMail(m).InDrafts Then ListAttachments m
    Next m
    SetStages
End Sub

' Where each packet stands: what the box at the end counts, what ENTRY shows and Mark filed marks.
' CHECK says it in words, by email (Stands).
Private Sub SetStages()
    Dim p As Long, m As Long
    For p = 1 To mNPkt
        m = mPkt(p).Mail
        With mPkt(p)
            If Len(.Problem) > 0 Then
                .Stage = ST_HELD
            ElseIf .Notarized > 0 Then
                .Stage = ST_BACK
            ElseIf mMail(m).IsSent Then
                .Stage = ST_OUT
            ElseIf Len(.VIMissing) > 0 Or (.FormState <> "MADE" And .FormState <> "KEPT") Then
                .Stage = ST_FILES
            ElseIf Len(mMail(m).Held) > 0 Then
                .Stage = ST_HELD
            ElseIf mMail(m).InDrafts Then
                .Stage = ST_DRAFT
            Else
                .Stage = ST_SEND
            End If
        End With
    Next p
End Sub

' Everything in the email's packet folders except the CI copies, notarized copies, packet.txt and the Upload folder.
' MF-12s first. A file it can't size up holds only this email.
Private Sub ListAttachments(ByVal m As Long)
    Dim p As Long, files As Variant, i As Long, pass As Long, f As String, stem As String, ext As String, isForm As Boolean
    Dim doing As String, bolNo As String, kind As String, sq As String, bols As String
    On Error GoTo Failed
    doing = "listing the attachments"
    mMail(m).Files = "": mMail(m).MB = 0
    For p = 1 To mNPkt
        If mPkt(p).Mail = m Then mPkt(p).MailForms = 0: mPkt(p).MailInvoices = 0
    Next p
    For pass = 1 To 2
        For p = 1 To mNPkt
            If mPkt(p).Mail = m Then
                files = ListFiles(mPkt(p).Folder)
                If Not IsEmpty(files) Then
                    SortList files, SORT_NAMES
                    For i = 1 To UBound(files)
                        f = files(i)
                        SplitExt f, stem, ext
                        isForm = IsExpectedForm(p, f)
                        If Not IsJunkFile(f) And InStr(UCase$(f), "NOTARIZED") = 0 And FileKind(stem) <> "CI" And _
                           Right$(LCase$(f), 8) <> ".partial" And LCase$(f) <> TAG_FILE Then
                            If (pass = 1 And isForm) Or (pass = 2 And Not isForm) Then
                                doing = "attaching " & f
                                mMail(m).Files = mMail(m).Files & mPkt(p).Folder & "\" & f & vbLf
                                mMail(m).MB = mMail(m).MB + FileLen(mPkt(p).Folder & "\" & f) / 1048576#
                                If isForm Then
                                    mPkt(p).MailForms = mPkt(p).MailForms + 1
                                ElseIf FileKind(stem) = "VI" Then
                                    ' one invoice per BOL: a second copy (VI 2) goes too, but it's the same invoice
                                    If Not ParseFiledName(stem, bolNo, kind, sq) Then bolNo = stem
                                    If InStr(bols, "|" & p & ":" & BolKey(bolNo) & "|") = 0 Then
                                        bols = bols & "|" & p & ":" & BolKey(bolNo) & "|"
                                        mPkt(p).MailInvoices = mPkt(p).MailInvoices + 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        Next p
    Next pass
    Exit Sub
Failed:
    HoldMail m, doing
End Sub

Private Sub MakeDrafts()
    Dim m As Long, ol As Object

    If Len(mTestDraftDir) = 0 Then
        Set ol = OutlookApp()
        If ol Is Nothing Then Err.Raise vbObjectError + 14, , "Couldn't open Outlook."
    End If
    For m = 1 To mNMail
        If Len(mMail(m).Held) = 0 And Not mMail(m).IsSent And Not mMail(m).InDrafts Then MakeOneDraft ol, m
    Next m
End Sub

' One email's draft, on its own: if it can't be made (Outlook won't take an attachment, say),
' that email is held with a line on CHECK and the others still get theirs.
Private Sub MakeOneDraft(ol As Object, ByVal m As Long)
    Dim subj As String, body As String, doing As String
    On Error GoTo Failed
    doing = "making the draft"
    subj = FillWords(mSubject, m)
    body = FillWords(mBody, m)
    If Len(mTestDraftDir) > 0 Then TestDraft m, subj, body, doing Else OutlookDraft ol, m, subj, body, doing
    mMail(m).Drafted = Now
    mMail(m).InDrafts = True
    RecordSubject m, subj
    Exit Sub
Failed:
    HoldMail m, doing
    Close
End Sub

' The draft's subject goes in each of its packets' packet.txt, so the email is found in Sent
' Items under that subject whatever the Short Name is later. The draft is made either way.
Private Sub RecordSubject(ByVal m As Long, ByVal subj As String)
    Dim p As Long
    On Error Resume Next
    For p = 1 To mNPkt
        If mPkt(p).Mail = m Then
            mPkt(p).Subject = subj
            If FolderExists(mPkt(p).Folder) Then WriteTag p
        End If
    Next p
    Close
    On Error GoTo 0
End Sub

' A draft that stops partway (an attachment Outlook won't take) is thrown away, never saved
' half made; the error goes back to MakeOneDraft.
Private Sub OutlookDraft(ol As Object, ByVal m As Long, ByVal subj As String, ByVal body As String, ByRef doing As String)
    Dim d As Object, insp As Object, f As Variant, html As String, pos As Long, toAddr As String, ccAddr As String
    Dim errNo As Long, errText As String
    On Error GoTo Failed
    Addresses mMail(m).Sup, toAddr, ccAddr
    Set d = ol.CreateItem(0)
    d.To = toAddr
    If Len(ccAddr) > 0 Then d.CC = ccAddr
    d.Subject = subj
    For Each f In Split(mMail(m).Files, vbLf)
        If Len(f) > 0 Then
            doing = "attaching " & Mid$(f, InStrRev(f, "\") + 1)
            d.Attachments.Add CStr(f)
        End If
    Next f
    doing = "putting in the letter"
    ' opening the item's window (without showing it) puts your default signature in the body
    On Error Resume Next
    Set insp = d.GetInspector
    html = d.HTMLBody
    On Error GoTo Failed
    pos = InStr(1, LCase$(html), "<body")
    If pos > 0 Then pos = InStr(pos, html, ">")
    If pos > 0 Then
        d.HTMLBody = Left$(html, pos) & BodyHtml(body) & Mid$(html, pos + 1)
    Else
        d.Body = body & vbCrLf & vbCrLf & d.Body
    End If
    doing = "saving the draft"
    d.Save
    d.Close 0
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    Resume ThrowAway
ThrowAway:
    On Error Resume Next
    If Not d Is Nothing Then d.Close 1          ' close without saving
    On Error GoTo 0
    Err.Raise errNo, , errText
End Sub

Private Function TestDraftPath(ByVal m As Long) As String
    TestDraftPath = mTestDraftDir & "\" & mSup(mMail(m).Sup).ShortName & " " & (mMail(m).PerKey \ 100) & "-" & _
        Format$(mMail(m).PerKey Mod 100, "00") & " draft.txt"
End Function

' A test run's draft, as a text file. Outlook reads each file as it's attached, so this reads
' them too: a file Outlook couldn't take stops a test draft the same way.
Private Sub TestDraft(ByVal m As Long, ByVal subj As String, ByVal body As String, ByRef doing As String)
    Dim f As Integer, toAddr As String, ccAddr As String, a As Variant
    For Each a In Split(mMail(m).Files, vbLf)
        If Len(a) > 0 Then
            doing = "attaching " & Mid$(a, InStrRev(a, "\") + 1)
            f = FreeFile
            Open CStr(a) For Binary Access Read As #f
            Close #f
        End If
    Next a
    doing = "saving the draft"
    Addresses mMail(m).Sup, toAddr, ccAddr
    f = FreeFile
    Open TestDraftPath(m) For Output As #f
    Print #f, "To: " & toAddr
    Print #f, "CC: " & ccAddr
    Print #f, "Subject: " & subj
    Print #f, "Body: " & body
    Print #f, "Attachments:"
    Print #f, mMail(m).Files;
    Close #f
End Sub

Private Function BodyHtml(ByVal t As String) As String
    Dim lines As Variant, i As Long, out As String, l As String
    lines = Split(Replace(t, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        l = Replace(Replace(Replace(lines(i), "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
        If Len(Trim$(l)) = 0 Then l = "&nbsp;"
        out = out & "<p class=MsoNormal>" & l & "</p>"
    Next i
    BodyHtml = out & "<p class=MsoNormal>&nbsp;</p>"
End Function

Private Function OutlookApp() As Object
    If Len(mTestDraftDir) > 0 Then Exit Function    ' test runs never touch Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

Private Function FindMail(fld As Object, ByVal subj As String) As Object
    Dim hits As Object
    On Error Resume Next
    Set hits = fld.Items.Restrict("[Subject] = '" & Replace(subj, "'", "''") & "'")
    If Not hits Is Nothing Then
        If hits.Count > 0 Then Set FindMail = hits(1)
    End If
    On Error GoTo 0
End Function


' ======================================================================================
'  Marking filed
' ======================================================================================

' Stamps today's date as Filed Date on every row of every packet showing on ENTRY, after a
' yes/no box names them. Hands back how many packets it marked.
Private Function MarkFiled() As Long
    Dim p As Long, i As Long, n As Long, names As String, ix As Variant

    ix = PacketOrder()
    For i = 1 To mNPkt
        p = ix(i)
        If mPkt(p).Stage = ST_BACK Then
            n = n + 1
            If n <= 30 Then names = names & "  " & PeriodLabel(mPkt(p).PerKey) & "  " & mPkt(p).PName & vbLf
        End If
    Next i
    If n = 0 Then
        Say "Nothing to mark: no packet is on ENTRY. A packet gets there once its notarized MF-12 is saved in its " & _
            "folder with NOTARIZED at the end of the name.", vbInformation
        Exit Function
    End If
    If n > 30 Then names = names & "  and " & (n - 30) & " more" & vbLf
    If Not AskYes("Stamp today's date as Filed Date on these " & n & " packet(s)?" & vbLf & vbLf & names & vbLf & _
                  "Say yes once they're all keyed into the Georgia portal. They come off ENTRY.") Then Exit Function

    For p = 1 To mNPkt
        If mPkt(p).Stage = ST_BACK Then StampFiled p
    Next p
    MarkFiled = n
End Function


' ======================================================================================
'  The CHECK sheet
' ======================================================================================

' CHECK, to be read in one glance: a bar on top, red with the count when anything needs you and
' green when nothing does; a small line saying when it was built; then Needs you, each problem
' with its fix beside it; then one line per email and one line per packet. No money: that's on
' the report and on ENTRY. Nothing is folded away.
Private Sub WriteCheck(ByVal stepNo As Long)
    Dim ws As Worksheet, r As Long, i As Long, m As Long, p As Long, ix As Variant, clr As Long, v As Variant
    Dim toAddr As String, ccAddr As String, x As Variant, small As String, bar As String

    Set ws = CheckSheet()
    bar = mNPkt & " packet" & IIf(mNPkt = 1, "", "s") & " open."
    If mNLook = 0 Then
        bar = "Nothing needs you.   " & bar
    Else
        bar = mNLook & IIf(mNLook = 1, " thing needs you.   ", " things need you.   ") & bar
    End If
    small = "Built " & Format$(Now, "m/d/yyyy h:nn AM/PM") & " by " & Choose(stepNo, "Build packets", "Make drafts", "Mark filed") & "."
    If Len(mPeriodNote) > 0 Then small = small & "  " & mPeriodNote
    If mLeftAlone > 0 Then small = small & "  Left alone: " & mLeftAlone & " open row(s) from before " & PeriodLabel(mOldest) & _
        ", the Oldest period to work on SETUP."
    If mMoved > 0 Then small = small & "  Moved " & mMoved & " paid row(s) to Refunds Received."
    r = CheckTop(ws, bar, IIf(mNLook = 0, GREEN, RED), small)

    If mNLook > 0 Then
        r = CheckPair(ws, r, "Needs you", "Fix", True)
        For i = 1 To mNLook
            r = CheckPair(ws, r, LookText(i), mLook(i).HowTo, False)
        Next i
        r = r + 1
    End If

    ' one line per supplier per period
    r = TableHead(ws, r, Array("Supplier", "Month", "To", "CC", "Where it stands"))
    ix = MailOrder()
    For i = 1 To mNMail
        m = ix(i)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = mSup(mMail(m).Sup).ShortName
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = PeriodLabel(mMail(m).PerKey)
        Addresses mMail(m).Sup, toAddr, ccAddr
        ws.Cells(r, 3).Value = toAddr
        ws.Cells(r, 4).Value = ccAddr
        ws.Cells(r, 5).Value = Stands(m, clr)
        ws.Cells(r, 5).Interior.Color = clr
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).WrapText = True
        ws.Rows(r).AutoFit
        r = r + 1
    Next i
    If mNMail = 0 Then ws.Cells(r, 1).Value = "No emails open.": r = r + 1
    r = r + 1

    ' one line per packet, oldest period first
    r = TableHead(ws, r, Array("Packet", "Month", "Loads", "MF-12s", "Invoices", "Folder"))
    ix = PacketOrder()
    For i = 1 To mNPkt
        p = ix(i)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = mPkt(p).PName
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = PeriodLabel(mPkt(p).PerKey)
        ws.Cells(r, 3).Value = mPkt(p).NRows
        v = FormsCell(p, clr)
        ws.Cells(r, 4).Value = v
        If clr <> 0 Then ws.Cells(r, 4).Interior.Color = clr
        ws.Cells(r, 5).Value = InvoicesCell(p, clr)
        If clr <> 0 Then ws.Cells(r, 5).Interior.Color = clr
        If FolderExists(mPkt(p).Folder) Then ws.Hyperlinks.Add Anchor:=ws.Cells(r, 6), Address:=mPkt(p).Folder, _
            TextToDisplay:=LastPart(mPkt(p).Folder)
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).WrapText = True
        ws.Rows(r).AutoFit
        r = r + 1
    Next i
    If mNPkt = 0 Then ws.Cells(r, 1).Value = "Nothing open.": r = r + 1

    ' every folder and file renamed this click, to a new Short Name or from the last version's names
    If Not mRenamed Is Nothing Then
        If mRenamed.Count > 0 Then
            r = r + 1
            r = CheckPair(ws, r, "Renamed this click", "Now", True)
            For Each x In mRenamed
                r = CheckPair(ws, r, Split(x, "|")(0), Split(x, "|")(1), False, False)
            Next x
        End If
    End If
    CheckWidths ws, r
End Sub

' CHECK when a click had to stop before it could look at anything: a red bar and what to do.
Private Sub WriteCheckStopped(ByVal stepNo As Long, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo Done
    Set ws = CheckSheet()
    r = CheckTop(ws, "Stopped before anything was checked.", RED, "Drawn " & Format$(Now, "m/d/yyyy h:nn AM/PM") & _
        " by " & Choose(stepNo, "Build packets", "Make drafts", "Mark filed") & ".")
    r = CheckPair(ws, r, "Needs you", "Fix", True)
    r = CheckPair(ws, r, "The click stopped before it could check the packets.", msg, False)
    CheckWidths ws, r
Done:
End Sub

' The CHECK tab, empty, made if it isn't there.
Private Function CheckSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = SheetNamed(CHECK_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = CHECK_SHEET
    End If
    ws.Hyperlinks.Delete
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Size = 10
    ws.Rows.RowHeight = ws.StandardHeight
    Set CheckSheet = ws
End Function

' The bar across the top, the small line under it, and a link to the packets folder. Hands back
' the row the lists start on.
Private Function CheckTop(ws As Worksheet, ByVal bar As String, ByVal clr As Long, ByVal small As String) As Long
    With ws.Range("A1:F1")
        .Merge
        .Value = bar
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = clr
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 28
    ws.Range("A2").Value = small
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = &H595959
    If Len(mBase) > 0 Then
        ws.Hyperlinks.Add Anchor:=ws.Range("A3"), Address:=mBase, TextToDisplay:="Packets folder: " & mBase
        ws.Range("A3").Font.Size = 9
    End If
    CheckTop = 5
End Function

' A line of two halves across CHECK, the first over A to C and the second over D to F: a Needs
' you line and its fix, or a rename. The row is made tall enough for the longer half.
Private Function CheckPair(ws As Worksheet, ByVal r As Long, ByVal a As String, ByVal b As String, ByVal isHead As Boolean, _
    Optional ByVal isProblem As Boolean = True) As Long
    Dim n As Long
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Merge
        .Value = a
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 6))
        .Merge
        .Value = b
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    If isHead Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Font.Bold = True
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = GRAY
    ElseIf isProblem Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = RED
    End If
    n = Int((Len(a) - 1) / 92) + 1
    If Int((Len(b) - 1) / 125) + 1 > n Then n = Int((Len(b) - 1) / 125) + 1
    ws.Rows(r).RowHeight = 13.5 * n
    CheckPair = r + 1
End Function

Private Sub CheckWidths(ws As Worksheet, ByVal r As Long)
    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 9
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 28
    ws.Columns(5).ColumnWidth = 46
    ws.Columns(6).ColumnWidth = 24
    If r > 5 Then
        ws.Range("A5:F" & r).VerticalAlignment = xlTop
    End If
    On Error Resume Next
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' A Needs you line: what's wrong, and what was found.
Private Function LookText(ByVal i As Long) As String
    LookText = mLook(i).What
    If Len(mLook(i).Detail) > 0 Then
        If Right$(LookText, 1) <> "." Then LookText = LookText & "."
        LookText = LookText & " " & mLook(i).Detail
    End If
End Function

' Where one email stands, in the words CHECK shows, and its color: green when it's your move,
' red when something holds it, blue once it's sent and waiting on it back, orange while an
' invoice isn't in yet.
Private Function Stands(ByVal m As Long, ByRef clr As Long) As String
    Dim p As Long, n As Long, back As Long, lastBack As Date, vis As String, why As String, notMade As Boolean
    For p = 1 To mNPkt
        If mPkt(p).Mail = m Then
            n = n + 1
            If Len(mPkt(p).Problem) > 0 And Len(why) = 0 Then why = mPkt(p).PName & ": " & mPkt(p).Problem
            If mPkt(p).Notarized > 0 Then
                back = back + 1
                If mPkt(p).Notarized > lastBack Then lastBack = mPkt(p).Notarized
            End If
            If Len(mPkt(p).VIMissing) > 0 Then vis = vis & IIf(Len(vis) > 0, ", ", "") & mPkt(p).VIMissing
            Select Case mPkt(p).FormState
                Case "MADE", "KEPT"
                Case "MISSING": notMade = True
                Case Else: If Len(why) = 0 Then why = mPkt(p).PName & ": MF-12 " & LCase$(mPkt(p).FormState) & " (see Needs you)"
            End Select
        End If
    Next p
    clr = GREEN
    If Len(why) > 0 Then
        Stands = "Held: " & why: clr = RED
    ElseIf mMail(m).IsSent And back = n Then
        Stands = "Back notarized " & Format$(lastBack, "m/d") & ": key it from ENTRY, then click Mark filed"
    ElseIf mMail(m).IsSent Then
        Stands = "Sent"
        If mMail(m).SentOn > 0 Then Stands = Stands & " " & Format$(mMail(m).SentOn, "m/d") & " (" & DaysAgo(mMail(m).SentOn) & ")"
        If back > 0 Then Stands = Stands & ", " & back & " of " & n & " back" Else Stands = Stands & ", waiting on it back"
        clr = SKY
    ElseIf Len(mMail(m).Held) > 0 And Not WaitOnly(m) Then
        Stands = "Held: " & mMail(m).Held: clr = RED
    ElseIf Len(vis) > 0 Then
        Stands = "Waiting on the VI for BOL " & vis: clr = PEACH
    ElseIf notMade Then
        Stands = "MF-12 not made yet: click Build packets"
    ElseIf mMail(m).Drafted > 0 Then
        Stands = "Drafted " & Format$(mMail(m).Drafted, "m/d h:nn AM/PM") & ": look it over in Outlook and send it"
    ElseIf mMail(m).InDrafts Then
        Stands = "In Outlook Drafts, not sent yet"
    ElseIf mMail(m).MB > BIG_EMAIL_MB Then
        Stands = "Ready: click Make drafts (" & Format$(mMail(m).MB, "0.0") & " MB may be too big for one email)"
    Else
        Stands = "Ready: click Make drafts"
    End If
End Function

' "0 days", "1 day", "23 days"
Private Function DaysAgo(ByVal d As Date) As String
    Dim n As Long
    n = Date - Int(d)
    DaysAgo = n & IIf(n = 1, " day", " days")
End Function

Private Function WaitOnly(ByVal m As Long) As Boolean
    On Error Resume Next
    WaitOnly = mWaitOnly(m)
End Function

' The Files list's MF-12s: how many, and anything to know about them.
Private Function FormsCell(ByVal p As Long, ByRef clr As Long) As Variant
    clr = 0
    With mPkt(p)
        If Len(.Problem) > 0 Then
            FormsCell = "Held: " & .Problem: clr = RED
        ElseIf .FormState = "MADE" Or .FormState = "KEPT" Then
            If .Notarized > 0 Then
                FormsCell = .NForms & ", back notarized " & Format$(.Notarized, "m/d"): clr = GREEN
            Else
                FormsCell = .NForms
            End If
        ElseIf .FormState = "MISSING" Then
            FormsCell = "not made yet"
        Else
            FormsCell = .NForms & ", " & LCase$(.FormState) & " (see Needs you)": clr = RED
        End If
    End With
End Function

' The Files list's Invoices: "3 of 3", or "2 of 3, waiting on BOL 80002".
Private Function InvoicesCell(ByVal p As Long, ByRef clr As Long) As String
    Dim r As Long, have As Long
    clr = 0
    For r = 1 To mNRow
        If mRow(r).Pkt = p And mRow(r).HasVI Then have = have + 1
    Next r
    InvoicesCell = have & " of " & mPkt(p).NRows
    If Len(mPkt(p).VIMissing) > 0 And FolderExists(mPkt(p).Folder) Then
        InvoicesCell = InvoicesCell & ", waiting on BOL " & mPkt(p).VIMissing
        clr = PEACH
    End If
End Function

Private Function TableHead(ws As Worksheet, ByVal r As Long, heads As Variant) As Long
    Dim i As Long
    For i = 0 To UBound(heads)
        ws.Cells(r, i + 1).Value = heads(i)
    Next i
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(heads) + 1))
        .Font.Bold = True
        .Interior.Color = GRAY
    End With
    TableHead = r + 1
End Function


' ======================================================================================
'  The HOW TO tab: three buttons and the steps, at the front of the workbook
' ======================================================================================

' The steps live inside the workbook, with buttons so nothing needs Alt F8: each button with,
' beside it, when to click it and what it does. Redrawn every run; the buttons are made once and
' only moved back into place, so a button you dragged still works.
Private Sub WriteHowTo()
    Dim ws As Worksheet, r As Long

    Set ws = SheetNamed(HOWTO_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(Before:=mBook.Worksheets(1))
        ws.Name = HOWTO_SHEET
    End If
    If ws.Index <> 1 Then ws.Move Before:=mBook.Worksheets(1)
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Rows.RowHeight = ws.StandardHeight
    ws.Cells.Font.Size = 12
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 40
    ws.Columns(3).ColumnWidth = 46
    ws.Columns(4).ColumnWidth = 70
    ws.Rows(1).RowHeight = 30

    ws.Range("B1").Value = "GA refunds: the running list, MF-12 packets, Outlook drafts, portal entry"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 18
    HowToButton ws, 3, "btnBuild", "1.  Build packets", "Refunds1_BuildPackets", _
        "After you paste new refunds, or save invoices in INBOX", _
        "Makes the folders, files the invoices from INBOX, fills the MF-12s, and draws CHECK and ENTRY. " & _
        "Click it any time: a second click changes nothing."
    HowToButton ws, 4, "btnDrafts", "2.  Make drafts", "Refunds2_MakeDrafts", _
        "When CHECK says Ready: click Make drafts", _
        "Makes one Outlook draft per supplier per month, its MF-12s and VIs attached. Nothing is sent."
    HowToButton ws, 5, "btnFiled", "3.  Mark filed", "Refunds3_MarkFiled", _
        "Once ENTRY is keyed into the Georgia portal", _
        "Stamps today's date as Filed Date on every packet on ENTRY, so they drop off. It asks first."
    r = 6
    r = HowToLine(ws, r, "Every click checks everything again, so CHECK is never out of date whichever button you click.", False)
    ws.Cells(r - 1, 2).Font.Italic = True
    r = r + 1

    r = HowToLine(ws, r, "Each month", True)
    r = HowToLine(ws, r, "1. New refunds: paste the refund report's rows at the bottom of Refunds Due, under the rows already there.", False)
    r = HowToLine(ws, r, "2. Invoices: save each load's invoices in INBOX as ""12345 VI"" and ""12345 CI"" (12345 is the BOL).", False)
    r = HowToLine(ws, r, "3. Click Build packets. It files the invoices and fills the MF-12s; CHECK shows where each one stands.", False)
    r = HowToLine(ws, r, "4. When CHECK says Ready, click Make drafts, look each draft over in Outlook, and send it yourself.", False)
    r = HowToLine(ws, r, "5. It comes back: save the notarized MF-12 in its packet folder with NOTARIZED in the name.", False)
    r = HowToLine(ws, r, "6. File it: key ENTRY into the Georgia portal, upload what its banner says, then click Mark filed.", False)
    r = HowToLine(ws, r, "7. The state pays: type the Refund Receive Date; the next click moves those rows to Refunds Received.", False)
    r = r + 1
    r = HowToLine(ws, r, "The tabs", True)
    r = HowToLine(ws, r, "SETUP: the company's details, the packets folder, the template, the oldest period, the refund types and " & _
        "the email words.", False)
    r = HowToLine(ws, r, "CHECK: where everything stands. The bar on top says if anything needs you; then each email and each packet.", False)
    r = HowToLine(ws, r, "ENTRY: the typing sheet for the Georgia portal, one filing per month and refund type. Your x's in Keyed stay.", False)
    r = HowToLine(ws, r, "Corp License Number: the suppliers. The Corp name prints on the MF-12 as typed; a new Short Name renames " & _
        "open folders.", False)
    r = HowToLine(ws, r, "Refunds Due: every refund still out. Refunds Received: the ones the state has paid.", False)
    r = r + 1
    r = HowToLine(ws, r, "Rules it keeps", True)
    r = HowToLine(ws, r, "It never sends an email, never deletes a file, and never writes over one.", False)
    r = HowToLine(ws, r, "A problem holds only its own packet or email. The rest still get built, and CHECK says what to fix.", False)
    r = HowToLine(ws, r, "A date you typed in a status column is never changed, and filed packets keep their names.", False)
    r = r + 1
    r = HowToLine(ws, r, "First time only", True)
    r = HowToLine(ws, r, "1. This goes in the GA tracking workbook, the one with Refunds Due and Corp License Number. " & _
        "It comes as two files, GARefundsReader.bas and GARefundsWorker.bas. Alt F11 opens the macro editor.", False)
    r = HowToLine(ws, r, "2. First take out the old piece(s), so there's only ever one copy: under Modules, right-click each " & _
        "one that holds this macro (GARefunds, GARefundsReader, GARefundsWorker, or a Module1 it was pasted into), " & _
        "click Remove, and No when it asks to export it.", False)
    r = HowToLine(ws, r, "3. Then File, Import File, once for each of the two files. Save the workbook as .xlsm.", False)
    r = HowToLine(ws, r, "4. Save " & TEMPLATE_NAME & " (it comes with this macro) in the same folder as the workbook.", False)
    r = HowToLine(ws, r, "5. You don't make any folders. The first click on Build packets asks for the company's name and " & _
        "address for the MF-12, the folder the packets go in, and which period to start from, once each.", False)
    r = HowToLine(ws, r, "If the buttons ever go missing: Alt F8, run Refunds1_BuildPackets, and this tab redraws with them.", False)
End Sub

' One line of the HOW TO tab, across B to D; hands back the next row.
Private Function HowToLine(ws As Worksheet, ByVal r As Long, ByVal text As String, ByVal isHead As Boolean) As Long
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Merge
        .Value = text
        .WrapText = True
        .VerticalAlignment = xlTop
        If isHead Then
            .Font.Bold = True
            .Font.Size = 14
            .Interior.Color = YELLOW
        End If
    End With
    ws.Rows(r).RowHeight = 16.5 * (Int((Len(text) - 1) / 135) + 1)
    HowToLine = r + 1
End Function

' A big button in column B of row r, when to click it in bold beside it and what it does after
' that. Made once, then only moved back into place.
Private Sub HowToButton(ws As Worksheet, ByVal r As Long, ByVal nm As String, ByVal caption As String, ByVal macroName As String, _
    ByVal whenText As String, ByVal whatText As String)
    Dim b As Object
    ws.Rows(r).RowHeight = 50
    With ws.Cells(r, 3)
        .Value = whenText
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(r, 4)
        .Value = whatText
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    On Error Resume Next
    Set b = ws.Buttons(nm)
    On Error GoTo 0
    If b Is Nothing Then
        Set b = ws.Buttons.Add(0, 0, 100, 30)
        b.Name = nm
    End If
    b.Placement = xlFreeFloating
    b.Left = ws.Cells(r, 2).Left + 4
    b.Top = ws.Cells(r, 2).Top + 4
    b.Width = ws.Cells(r, 2).Width - 8
    b.Height = ws.Cells(r, 2).Height - 8
    b.Caption = caption
    b.OnAction = macroName
    On Error Resume Next
    b.Font.Size = 18
    b.Font.Bold = True
    On Error GoTo 0
End Sub


' ======================================================================================
'  The ENTRY sheet: the typing sheet for the Georgia portal
' ======================================================================================

' The portal takes each invoice typed by hand, one field at a time, no paste. So: every packet
' back notarized and not filed yet (a held one stays on CHECK), one filing per period and refund
' type (one MFD-33 each) with its suppliers inside, the columns in the portal's own order, the
' rows in the order they sit on the form, a running total that starts over at each filing, and a
' total under each block. An x typed in Keyed stays on its row when the sheet is drawn again.
' Values only. Redrawn every run.
Private Sub WriteEntry()
    Dim ws As Worksheet, r As Long, p As Long, i As Long, k As Long, ix As Variant, order As Variant, s As Long, t As Long
    Dim g As Double, a As Double, fg As Double, fa As Double, formNo As Long, slot As Long, head As String
    Dim curT As Long, curPer As Long, bannerRow As Long, tg As Double, ta As Double, tRows As Long, tForms As Long, tPkts As Long
    Dim shown As Long, runTot As Double, marks As Object, seen As Object, key As String

    Set ws = SheetNamed(ENTRY_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = ENTRY_SHEET
    End If
    Set marks = KeyedMarks(ws)
    Set seen = CreateObject("Scripting.Dictionary")
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Rows.RowHeight = ws.StandardHeight
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "Portal entry: every packet back notarized and not filed yet"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Value = "One filing per month and refund type, its suppliers inside, one block per MF-12 packet. " & _
        "Type each row into the Georgia portal left to right, and put an x in Keyed as you go. Exempt Portion stays blank. " & _
        "Tax Paid to Seller is the gross refund on every row, even for a NET supplier; NET is only on its MF-12. " & _
        "The running total starts over at each filing, so if the portal's total is off you can find the row. " & _
        "Held packets aren't here; CHECK says why. When it's all keyed, click Mark filed on the HOW TO tab. " & _
        "Drawn " & Format$(Now, "m/d/yyyy h:nn AM/PM") & "."
    With ws.Range("A2:J2")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(2).RowHeight = 58
    r = 4

    order = PacketOrder()
    curT = 0: curPer = 0
    For i = 1 To mNPkt
        p = order(i)
        If mPkt(p).Stage = ST_BACK Then
            shown = shown + 1
            s = mPkt(p).Sup
            t = mPkt(p).TypeNo
            ' A new period or refund type: close the last filing with its MFD-33 total, open this one
            ' with a banner and the line on what to upload. Both are filled in once the filing is done
            ' (see EntryGroupClose).
            If t <> curT Or mPkt(p).PerKey <> curPer Then
                If curT > 0 Then r = EntryGroupClose(ws, r, curT, curPer, bannerRow, tg, ta, tRows, tForms, tPkts)
                curT = t
                curPer = mPkt(p).PerKey
                bannerRow = r
                ws.Rows(r).RowHeight = 28
                r = r + 3
                tg = 0: ta = 0: tRows = 0: tForms = 0: tPkts = 0: runTot = 0
            End If
            head = mSup(s).ShortName & " (" & mPkt(p).PName & "), " & mPkt(p).NRows & " row" & IIf(mPkt(p).NRows = 1, "", "s")
            If mPkt(p).NForms > 1 Then head = head & " on " & mPkt(p).NForms & " MF-12 forms"
            If mSup(s).IsNet Then head = head & " (its MF-12 shows NET; the figures here are gross, as the state wants)"
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
                .Merge
                .Value = head
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = YELLOW
            End With
            ws.Rows(r).RowHeight = 18
            r = r + 1
            tRows = tRows + mPkt(p).NRows: tForms = tForms + mPkt(p).NForms: tPkts = tPkts + 1
            r = TableHead(ws, r, Split(ENTRY_HEADS, "|"))
            ix = PacketRows(p)
            g = 0: a = 0: fg = 0: fa = 0: formNo = 1: slot = 0
            For k = 1 To UBound(ix)
                With mRow(ix(k))
                    key = LoadKey(.License, .Inv, .InvDate, .Gallons)
                    seen(key) = seen(key) + 1
                    key = key & "|" & seen(key)
                    If marks.Exists(key) Then ws.Cells(r, 1).Value = marks(key)
                    ws.Cells(r, 1).HorizontalAlignment = xlCenter
                    ws.Cells(r, 2).Value = mSup(s).Corp
                    ws.Cells(r, 2).ShrinkToFit = True
                    ws.Cells(r, 3).NumberFormat = "@"
                    ws.Cells(r, 3).Value = .License
                    ws.Cells(r, 4).NumberFormat = "mm/dd/yyyy"
                    ws.Cells(r, 4).Value = .InvDate
                    ws.Cells(r, 5).NumberFormat = "@"
                    ws.Cells(r, 5).Value = .Inv
                    ws.Cells(r, 6).NumberFormat = GalFormat(.Gallons)
                    ws.Cells(r, 6).Value = .Gallons
                    ' Always the gross Refund, never TTL: the state takes gross on the portal and
                    ' the MFD-33; the NET figure is only for the supplier's own MF-12.
                    ws.Cells(r, 7).NumberFormat = "#,##0.00"
                    ws.Cells(r, 7).Value = Round(.Refund, 2)
                    ws.Cells(r, 9).Value = Trim$(.FuelOpt)
                    runTot = Round(runTot + Round(.Refund, 2), 2)
                    ws.Cells(r, 10).NumberFormat = "#,##0.00"
                    ws.Cells(r, 10).Value = runTot
                    g = g + .Gallons: a = a + .Refund
                    fg = fg + .Gallons: fa = fa + .Refund
                End With
                r = r + 1
                slot = slot + 1
                If mPkt(p).NForms > 1 And (slot = ROWS_PER_FORM Or k = UBound(ix)) Then
                    ws.Cells(r, 2).Value = "end of MF-12 (" & formNo & " of " & mPkt(p).NForms & ")"
                    ws.Cells(r, 6).NumberFormat = GalFormat(fg)
                    ws.Cells(r, 6).Value = Round(fg, 3)
                    ws.Cells(r, 7).NumberFormat = "#,##0.00"
                    ws.Cells(r, 7).Value = Round(fa, 2)
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Font.Italic = True
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = GRAY
                    r = r + 1
                    formNo = formNo + 1: slot = 0: fg = 0: fa = 0
                End If
            Next k
            ws.Cells(r, 2).Value = "Total for " & mPkt(p).PName & ", " & mPkt(p).NRows & " row" & IIf(mPkt(p).NRows = 1, "", "s")
            ws.Cells(r, 6).NumberFormat = GalFormat(g)
            ws.Cells(r, 6).Value = Round(g, 3)
            ws.Cells(r, 7).NumberFormat = "#,##0.00"
            ws.Cells(r, 7).Value = Round(a, 2)
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
                .Font.Bold = True
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlDouble
            End With
            tg = tg + g: ta = ta + a
            r = r + 2
        End If
    Next i
    If curT > 0 Then r = EntryGroupClose(ws, r, curT, curPer, bannerRow, tg, ta, tRows, tForms, tPkts)
    If shown = 0 Then
        ws.Cells(r, 1).Value = "Nothing to key yet. A packet shows here once its notarized MF-12 is saved in its folder " & _
            "(NOTARIZED at the end of the name), and drops off when you click Mark filed."
        r = r + 1
    End If

    ' Tight for reading at a desk: the invoice rows close together and the numbers close to each
    ' other, with a clear gap between Date and Invoice Number. The banners, block heads and MFD-33
    ' totals keep their own heights.
    ws.Range("A4:J" & r).VerticalAlignment = xlCenter
    For k = 4 To r
        If ws.Cells(k, 2).Value = Split(ENTRY_HEADS, "|")(1) Then
            ws.Range(ws.Cells(k, 1), ws.Cells(k, 10)).WrapText = True
            ws.Rows(k).RowHeight = 26
        ElseIf ws.Cells(k, 1).Font.Size < 12 And Not ws.Cells(k, 1).MergeCells Then
            ws.Rows(k).RowHeight = 12.75
        End If
    Next k
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 12
    ws.Columns(4).ColumnWidth = 11
    ws.Columns(5).ColumnWidth = 15
    ws.Columns(6).ColumnWidth = 9
    ws.Columns(7).ColumnWidth = 11
    ws.Columns(8).ColumnWidth = 9
    ws.Columns(9).ColumnWidth = 14
    ws.Columns(10).ColumnWidth = 12
    ws.Range("D4:D" & r).HorizontalAlignment = xlLeft
    ws.Range("E4:E" & r).HorizontalAlignment = xlLeft
    ws.Range("E5:E" & r).IndentLevel = 2
    ws.Range("F4:G" & r).HorizontalAlignment = xlRight
    ws.Range("I5:I" & r).IndentLevel = 1
    ws.Range("J4:J" & r).HorizontalAlignment = xlRight
    ' the title stays put while the filings scroll
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' The x's (or anything else) typed in Keyed, by the load on their row: its license, invoice
' number, date and gallons, so the next drawing puts each back beside its load even when the
' order changes. Only loads alike in all four go by which of them it is.
Private Function KeyedMarks(ws As Worksheet) As Object
    Dim d As Object, seen As Object, r As Long, last As Long, key As String, ours As Boolean, heads As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set KeyedMarks = d
    heads = Split(ENTRY_HEADS, "|")
    On Error GoTo Done
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 4 To last
        If CStr(ws.Cells(r, 1).Value) = heads(0) And CStr(ws.Cells(r, 2).Value) = heads(1) Then
            ours = True                ' drawn this way, so Keyed is column A
        ElseIf ours And Len(CStr(ws.Cells(r, 3).Value)) > 0 And Len(CStr(ws.Cells(r, 5).Value)) > 0 And _
               Not ws.Cells(r, 1).MergeCells Then
            key = LoadKey(CStr(ws.Cells(r, 3).Value), CStr(ws.Cells(r, 5).Value), ws.Cells(r, 4).Value, ws.Cells(r, 6).Value)
            seen(key) = seen(key) + 1
            If Len(CStr(ws.Cells(r, 1).Value)) > 0 Then d(key & "|" & seen(key)) = ws.Cells(r, 1).Value
        End If
    Next r
Done:
End Function

' One load on ENTRY, the way the Keyed x's find it again: license, invoice number, date, gallons.
Private Function LoadKey(ByVal lic As String, ByVal inv As String, ByVal d As Variant, ByVal gal As Variant) As String
    LoadKey = UCase$(lic) & "|" & inv & "|"
    If IsDate(d) Then LoadKey = LoadKey & Format$(d, "yyyymmdd") Else LoadKey = LoadKey & CStr(d)
    If IsNumeric(gal) Then LoadKey = LoadKey & "|" & Format$(gal, "0.000") Else LoadKey = LoadKey & "|" & CStr(gal)
End Function

' Closes a filing on ENTRY: the MFD-33 total line, two blank rows, and the banner above the
' filing filled in with what it holds, with the line on what to upload under it. Hands back the
' next row.
Private Function EntryGroupClose(ws As Worksheet, ByVal r As Long, ByVal t As Long, ByVal perKey As Long, ByVal bannerRow As Long, _
    ByVal tg As Double, ByVal ta As Double, ByVal tRows As Long, ByVal tForms As Long, ByVal tPkts As Long) As Long
    Dim txt As String, when As String
    when = MonthName(perKey Mod 100) & " " & (perKey \ 100)
    ws.Cells(r, 2).Value = "MFD-33 total, " & when & " " & mTypePhrase(t) & ", " & tRows & " row" & IIf(tRows = 1, "", "s") & _
        " on " & tForms & " MF-12" & IIf(tForms = 1, "", "s")
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Merge
    ws.Cells(r, 6).NumberFormat = GalFormat(tg)
    ws.Cells(r, 6).Value = Round(tg, 3)
    ws.Cells(r, 7).NumberFormat = "#,##0.00"
    ws.Cells(r, 7).Value = Round(ta, 2)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = BLUE_PALE
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    ws.Rows(r).RowHeight = 18

    txt = when & ", " & mTypePhrase(t) & ": one filing, " & tRows & " row" & IIf(tRows = 1, "", "s") & " on " & _
        tForms & " MF-12" & IIf(tForms = 1, "", "s") & ", tax " & Format$(Round(ta, 2), "#,##0.00")
    With ws.Range(ws.Cells(bannerRow, 1), ws.Cells(bannerRow, 10))
        .Merge
        .Value = txt
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = vbWhite
        .Interior.Color = BLUE_DARK
        .VerticalAlignment = xlCenter
    End With
    txt = "Upload with it: the " & IIf(tForms = 1, "notarized MF-12", tForms & " notarized MF-12s") & _
        " as MF-12 refund waiver, the VI+CI combos from the Upload folder of " & _
        IIf(tPkts = 1, "its packet", "each of its " & tPkts & " packets") & ", and the signed MFD-33 as refund attachment."
    With ws.Range(ws.Cells(bannerRow + 1, 1), ws.Cells(bannerRow + 1, 10))
        .Merge
        .Value = txt
        .Font.Size = 11
        .Interior.Color = BLUE_PALE
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(bannerRow + 1).RowHeight = 30
    EntryGroupClose = r + 3
End Function

' Whole gallons show whole; anything else shows its three places.
Private Function GalFormat(ByVal x As Double) As String
    If Abs(x - Round(x, 0)) < 0.0005 Then GalFormat = "#,##0" Else GalFormat = "#,##0.000"
End Function


' ======================================================================================
'  Questions and boxes
' ======================================================================================

Private Sub Say(ByVal msg As String, ByVal icon As VbMsgBoxStyle)
    mLastSaid = msg
    If mQuiet Then
        mNotes = mNotes & msg & vbLf
    Else
        MsgBox msg, icon, "GA refunds"
    End If
End Sub

' A question in a box, or the next answer handed in by a test run. Blank when cancelled.
Private Function Ask(ByVal prompt As String, ByVal guess As String) As String
    If mQuiet Then
        mNotes = mNotes & "asked: " & Replace(prompt, vbLf, " / ") & vbLf
        Ask = NextAnswer()
    Else
        Ask = Trim$(InputBox(prompt, "GA refunds", guess))
    End If
End Function

Private Function AskFolder(ByVal title As String) As String
    Dim fd As Object
    If mQuiet Then AskFolder = NextAnswer(): Exit Function
    Set fd = Application.FileDialog(4)          ' a folder picker
    fd.Title = title
    fd.AllowMultiSelect = False
    fd.InitialFileName = mBook.Path & "\"
    If fd.Show = -1 Then AskFolder = fd.SelectedItems(1)
End Function

Private Function AskYes(ByVal prompt As String) As Boolean
    If mQuiet Then
        mNotes = mNotes & "asked: " & Replace(prompt, vbLf, " / ") & vbLf
        AskYes = (LCase$(NextAnswer()) = "yes")
    Else
        AskYes = (MsgBox(prompt, vbYesNo + vbQuestion, "GA refunds") = vbYes)
    End If
End Function

Private Function NextAnswer() As String
    If mFeed Is Nothing Then Exit Function
    If mFeed.Count = 0 Then Exit Function
    NextAnswer = mFeed(1)
    mFeed.Remove 1
End Function


' ======================================================================================
'  Files and folders
' ======================================================================================

Private Sub SplitExt(ByVal f As String, ByRef stem As String, ByRef ext As String)
    Dim d As Long
    d = InStrRev(f, ".")
    If d > 1 Then
        stem = Left$(f, d - 1): ext = Mid$(f, d)
    Else
        stem = f: ext = ""
    End If
End Sub

' The kind of copy a file is (VI, CI, ...): the last word of the name, skipping a trailing number.
Private Function FileKind(ByVal stem As String) As String
    Dim t As Variant, i As Long
    t = Split(Trim$(stem), " ")
    For i = UBound(t) To 0 Step -1
        If Not IsAllDigits(CStr(t(i))) Then FileKind = UCase$(t(i)): Exit Function
    Next i
End Function

Private Function IsJunkFile(ByVal f As String) As Boolean
    f = LCase$(f)
    IsJunkFile = (Left$(f, 2) = "~$" Or f = "thumbs.db" Or f = "desktop.ini")
End Function

' Every file in a folder (not subfolders), as an array from 1, or Empty.
Private Function ListFiles(ByVal folder As String) As Variant
    Dim found() As String, n As Long, f As String
    If Not FolderExists(folder) Then Exit Function
    f = Dir(folder & "\*", vbNormal Or vbReadOnly)
    Do While Len(f) > 0
        n = n + 1
        ReDim Preserve found(1 To n)
        found(n) = f
        f = Dir()
    Loop
    If n > 0 Then ListFiles = found
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then FolderExists = False
    On Error GoTo 0
End Function

Private Sub MkDirs(ByVal p As String)
    Dim parts As Variant, i As Long, cur As String, start As Long
    If FolderExists(p) Then Exit Sub
    parts = Split(p, "\")
    If Left$(p, 2) = "\\" Then
        cur = "\\" & parts(2) & "\" & parts(3)
        start = 4
    Else
        cur = parts(0)
        start = 1
    End If
    For i = start To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If Not FolderExists(cur) Then MkDir cur
        End If
    Next i
End Sub

' Every folder in a folder, as an array from 1, or Empty.
Private Function ListFolders(ByVal fld As String) As Variant
    Dim found() As String, n As Long, f As String
    If Not FolderExists(fld) Then Exit Function
    f = Dir(fld & "\*", vbDirectory)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            If FolderExists(fld & "\" & f) Then
                n = n + 1
                ReDim Preserve found(1 To n)
                found(n) = f
            End If
        End If
        f = Dir()
    Loop
    If n > 0 Then ListFolders = found
End Function

' The last part of a path: the folder or file's own name.
Private Function LastPart(ByVal path As String) As String
    LastPart = Mid$(path, InStrRev(path, "\") + 1)
End Function

' A path from inside the packets folder, the way CHECK shows it.
Private Function RelPath(ByVal path As String) As String
    If UCase$(Left$(path, Len(mBase) + 1)) = UCase$(mBase & "\") Then RelPath = Mid$(path, Len(mBase) + 2) Else RelPath = path
End Function

' A whole text file, as it is.
Private Function ReadText(ByVal path As String) As String
    ReadText = StrConv(ReadBytes(path), vbUnicode)
End Function

' A whole file's bytes, one to a byte of the string, the way the PDF code reads them. Blank for an
' empty file. The one place a file is read.
Private Function ReadBytes(ByVal path As String) As String
    Dim f As Integer, b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
        ReadBytes = b
    End If
    Close #f
End Function

' packet.txt: whose packet this folder is, so it's found whatever the folder is called. A plain
' text file, written only when what it should say has changed, never attached to an email.
Private Sub WriteTag(ByVal p As Long)
    Dim path As String, txt As String, f As Integer
    path = mPkt(p).Folder & "\" & TAG_FILE
    txt = TagText(p)
    If FileExists(path) Then
        If ReadText(path) = txt Then Exit Sub
    End If
    f = FreeFile
    Open path For Output As #f
    Print #f, txt;
    Close #f
End Sub

' What packet.txt says. While the folder is still under its old name (a rename that couldn't
' happen), it keeps the Short Name it had, so the next click knows what the old name was.
Private Function TagText(ByVal p As Long) As String
    Dim sh As String
    sh = mSup(mPkt(p).Sup).ShortName
    If LastPart(mPkt(p).Folder) <> mPkt(p).PName And Len(mPkt(p).TagShort) > 0 Then sh = mPkt(p).TagShort
    TagText = "This folder is one GA refund packet. The refunds macro finds it by the lines below, whatever the folder " & _
        "is called, so leave this file in it." & vbCrLf & _
        "Supplier: " & sh & vbCrLf & _
        "License: " & mSup(mPkt(p).Sup).License & vbCrLf & _
        "Refund type: " & mPkt(p).Word & vbCrLf & _
        "Period: " & PeriodLabel(mPkt(p).PerKey) & vbCrLf
    If Len(mPkt(p).Subject) > 0 Then TagText = TagText & "Subject: " & mPkt(p).Subject & vbCrLf
End Function

Private Function AnsiB(ByVal s As String) As String
    AnsiB = StrConv(s, vbFromUnicode)
End Function
