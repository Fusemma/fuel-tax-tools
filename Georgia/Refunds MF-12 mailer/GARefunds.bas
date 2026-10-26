' GA refunds: works the running list of GA refunds on Refunds Due, from the invoices to the
' MF-12 packets, the Outlook drafts, the portal entry and the refund check.
'
' Lives in the GA tracking workbook. The HOW TO tab it puts at the front has the three buttons
' and the steps. In short:
'   - Refunds Due is the running list: every refund still out, whatever its period. Paste each
'     new refund report's rows at the bottom.
'   - Build packets works every row with no Filed Date: a folder for each period, refund type
'     and supplier, the invoices filed into them from INBOX, the MF-12s filled, and CHECK
'     showing what each packet is waiting on.
'   - Make drafts makes one Outlook draft per supplier per period, its MF-12s and VIs attached.
'     It never sends. Look each one over in Outlook and send it yourself.
'   - Mark filed stamps today's date as Filed Date on every packet showing on ENTRY, once
'     you've keyed them into the Georgia portal.
'   - Every click stamps MF-12 Sent (the email is in Sent Items) and MF-12 Signed (a NOTARIZED
'     copy is in the packet folder), and moves rows with a Refund Receive Date to the bottom
'     of Refunds Received.
'
' The first click asks for what it needs, once each: the company's name and address for the
' MF-12, the folder the packets go in, and the oldest period to work. It finds
' mf12_template.pdf itself beside the workbook. It makes every folder itself.
'
' Suppliers are matched on the license number. Anything it can't place goes on the LOOK
' list on the CHECK sheet, and that supplier gets no draft until it's fixed.
' Every run also writes the Corp name from Corp License Number onto each matched row of
' Refunds Due, so the report's own spelling of a supplier never shows anywhere you work from.
' It never deletes your files and never writes over a file that's already there.
' Each MF-12 it makes carries a fingerprint of the rows on it, so if the report changes
' after a form is made, CHECK says so instead of letting the old form go out.
'
' Paste into an empty module in the tracking workbook and save it as .xlsm.

Option Explicit

Private Const ROWS_PER_FORM As Long = 21       ' invoice lines on one MF-12; more spill onto a second form
Private Const BIG_EMAIL_MB As Double = 20      ' a draft bigger than this gets a warning on CHECK
Private Const SETUP_SHEET As String = "SETUP"
Private Const CHECK_SHEET As String = "CHECK"
Private Const ENTRY_SHEET As String = "ENTRY"
Private Const HOWTO_SHEET As String = "HOW TO"
Private Const CONTACTS_SHEET As String = "Corp License Number"
Private Const RECEIVED_SHEET As String = "Refunds Received"
Private Const TEMPLATE_NAME As String = "mf12_template.pdf"

' Column names it looks for, spellings separated by |. Found by name, so moved columns are fine.
Private Const H_TYPE As String = "refund type|schedule"
Private Const H_PERIOD As String = "period"
Private Const H_CORP As String = "corp"
Private Const H_LICENSE As String = "license number|license"
Private Const H_BOL As String = "bol"
Private Const H_INV As String = "inv|invoice"
Private Const H_DATE As String = "inv date|invoice date"
Private Const H_GALLONS As String = "gallons"
Private Const H_REFUND As String = "refund"
Private Const H_ALLOW As String = "allowance"
Private Const H_TTL As String = "ttl refund|total refund"
Private Const H_FUEL As String = "fuel"
Private Const H_SENT As String = "mf-12 sent|mf12 sent"
Private Const H_SIGNED As String = "mf-12 signed|mf12 signed"
Private Const H_FILED As String = "filed date|date filed"
Private Const H_RECEIVED As String = "refund receive date|refund received date|refund received|date received"
Private Const H_CONTACT As String = "contact name|contact"
Private Const H_EMAIL As String = "email address|email"
Private Const H_GROUP As String = "group email address|group email"
Private Const H_EXTRA As String = "additional email|additional email address|additional emails"
Private Const H_SHORT As String = "short name"
Private Const H_NET As String = "net"

' Where a packet stands, in the order it gets there
Private Const ST_HELD As Long = 1        ' a row or the supplier needs a fix first
Private Const ST_FILES As Long = 2       ' waiting on a VI, or its MF-12 isn't right
Private Const ST_SEND As Long = 3        ' ready for Make drafts
Private Const ST_DRAFT As Long = 4       ' in Outlook Drafts, not sent yet
Private Const ST_OUT As Long = 5         ' sent, waiting on the notarized MF-12
Private Const ST_BACK As Long = 6        ' back notarized: key it from ENTRY, then Mark filed

' Colors, written back to front the way Excel wants them (blue, green, red)
Private Const GREEN As Long = &HCEEFC6
Private Const RED As Long = &HCEC7FF
Private Const YELLOW As Long = &HCCF2FF
Private Const GRAY As Long = &HD9D9D9
Private Const BLUE_DARK As Long = &H794E1F
Private Const BLUE_PALE As Long = &HF2E6DD

Private Type RRow
    SheetRow As Long
    TypeText As String
    PerKey As Long           ' 202608 for Aug-26
    Corp As String
    License As String
    Bol As String
    BolK As String
    Inv As String
    InvDate As Date
    Gallons As Double
    Refund As Double
    Allow As Double
    Ttl As Double
    FuelText As String
    FuelOpt As String        ' the form's own spelling of the fuel, padding and all
    Prefix As String         ' LIC, GOV, DIV
    Sup As Long              ' which supplier, 0 if none
    Pkt As Long              ' which packet, 0 if none
    Amount As Double         ' what goes on the MF-12: TTL Refund for NET suppliers, else Refund
    Problem As String
    HasVI As Boolean
    HasSent As Boolean       ' MF-12 Sent already has something in it
    SentOn As Date
    HasSigned As Boolean     ' MF-12 Signed already has something in it
End Type

Private Type RSup
    License As String
    LicKey As String
    Corp As String
    ShortName As String
    IsNet As Boolean
    Contact As String
    Email As String
    GroupMail As String
    Extra As String          ' Additional Email
    SheetRow As Long
    Problem As String
    NRows As Long
End Type

' One email: a supplier's packets for one period
Private Type RMail
    Sup As Long
    PerKey As Long
    NRows As Long
    Held As String           ' why no draft, blank when ready
    Status As String
    IsSent As Boolean
    SentOn As Date
    InDrafts As Boolean
    Files As String
    FileCount As Long
    MB As Double
End Type

Private Type RPkt
    Key As String            ' 202608|LIC-MANSFIELD
    PName As String          ' LIC-MANSFIELD
    Prefix As String
    PerKey As Long
    Sup As Long
    Mail As Long
    Folder As String
    NRows As Long
    Gallons As Double
    Refund As Double
    Allow As Double
    Ttl As Double
    Amount As Double
    NForms As Long
    Problem As String        ' blank when its rows are all good
    VIMissing As String
    FormState As String      ' what's on disk: MADE, KEPT, MISSING, or what's wrong
    Notarized As Date        ' the newest NOTARIZED copy's date, 0 if none
    Stage As Long
    Waiting As String
End Type

Private Type RLook
    What As String
    Detail As String
    HowTo As String
End Type

Private mRow() As RRow, mNRow As Long
Private mSup() As RSup, mNSup As Long
Private mMail() As RMail, mNMail As Long
Private mPkt() As RPkt, mNPkt As Long
Private mLook() As RLook, mNLook As Long
Private mTypeText() As String, mTypePrefix() As String, mTypePhrase() As String, mNType As Long

Private mBook As Workbook
Private mInboxLeft As Long
Private mLeftAlone As Long       ' open rows from before the oldest period, not worked
Private mMoved As Long           ' paid rows moved to Refunds Received this click
Private mShortsFilled As Long    ' Short Names filled in from the Corp name this click

' from SETUP
Private mBuyer As String, mAddress As String, mCity As String, mState As String, mZip As String
Private mUs As String, mBase As String, mTemplatePath As String, mOldest As Long
Private mSubject As String, mBody As String

Private mTemplate As String      ' the template's bytes
Private mStage As String         ' what it was doing, for the error message
Private mQuiet As Boolean        ' test runs: no popups
Private mFeed As Collection      ' test runs: the answers to the questions, in order
Private mTestDraftDir As String  ' test runs: drafts go to text files here instead of Outlook
Private mNotes As String

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
    Dim msg As String, marked As Long

    On Error GoTo Failed
    mStage = "getting started"
    ' Excel keeps these between runs, so start every run from nothing
    Erase mTypeText, mTypePrefix, mTypePhrase
    mNType = 0: mTemplate = "": mMoved = 0: mShortsFilled = 0
    ClearLists
    Set mBook = Book()
    If mBook Is Nothing Then
        Say "Can't find the Refunds Due sheet. Paste this macro into the GA tracking workbook (the one with " & _
            "Refunds Due and Corp License Number) and click the button there.", vbExclamation
        Exit Sub
    End If
    If Len(mBook.Path) = 0 Then
        Say "Save this workbook first: File, Save As, and pick Excel Macro-Enabled Workbook (.xlsm). " & _
            "Then click Build packets again.", vbExclamation
        Exit Sub
    End If

    mStage = "writing the HOW TO tab"
    WriteHowTo

    mStage = "checking the setup"
    msg = EnsureSetup()
    If Len(msg) > 0 Then Say msg, vbExclamation: Exit Sub
    If Not AskSetup() Then Exit Sub
    msg = ReadSetup()
    If Len(msg) > 0 Then Say "Fill these in on the SETUP sheet first:" & vbLf & msg, vbExclamation: Exit Sub
    If Not LoadTemplate() Then
        Say "The MF-12 template on SETUP isn't there, or isn't the right file:" & vbLf & mTemplatePath & vbLf & _
            "It's " & TEMPLATE_NAME & ", which comes with this macro. Save it beside this workbook, clear the " & _
            "MF-12 template file cell on SETUP, and click Build packets again.", vbExclamation
        Exit Sub
    End If

    mStage = "moving paid rows to Refunds Received"
    mMoved = MoveReceived()

    Application.ScreenUpdating = False
    LoadAll
    If stepNo = 1 Then
        mStage = "making folders under " & mBase
        MakeFolders
        mStage = "filing the INBOX"
        SortInbox
    Else
        CountInbox
    End If

    mStage = "checking the packet folders"
    CheckFiles
    If stepNo = 1 Then
        mStage = "filling the MF-12s"
        WriteForms
    End If
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
            ClearLists
            LoadAll
            CountInbox
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
    Application.ScreenUpdating = True
    Say "Stopped with an error while " & mStage & ":" & vbLf & Err.Description, vbCritical
End Sub

Private Sub ClearLists()
    Erase mRow, mSup, mMail, mPkt, mLook
    mNRow = 0: mNSup = 0: mNMail = 0: mNPkt = 0: mNLook = 0: mInboxLeft = 0: mLeftAlone = 0
End Sub

' Reads the contacts and the open rows, and sorts the rows into packets and emails.
Private Sub LoadAll()
    mStage = "reading Corp License Number"
    LoadContacts
    mStage = "reading Refunds Due"
    LoadRows
    mStage = "matching rows to suppliers"
    AssignPackets
    mStage = "writing the Corp names onto Refunds Due"
    WriteCorpNames
End Sub

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
            If Left$(mMail(m).Status, 7) = "DRAFTED" Then made = made + 1
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
    If mNLook > 0 Then msg = msg & vbLf & mNLook & " thing(s) on the CHECK sheet need a look."

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
'  Setup
' ======================================================================================

' The tracking workbook: the one this code lives in, or if it lives somewhere else
' (PERSONAL, say), the workbook in front.
Private Function Book() As Workbook
    If Not FindRefundsSheet(ThisWorkbook) Is Nothing Then
        Set Book = ThisWorkbook
    ElseIf Not ActiveWorkbook Is Nothing Then
        If Not FindRefundsSheet(ActiveWorkbook) Is Nothing Then Set Book = ActiveWorkbook
    End If
End Function

' Adds what it needs the first time, without stopping: the SETUP sheet, Short Name and NET on
' Corp License Number, and any of MF-12 Sent, MF-12 Signed, Filed Date and Refund Receive Date
' that Refunds Due is missing. Hands back a reason to stop, blank when it can go on.
Private Function EnsureSetup() As String
    Dim ws As Worksheet, cs As Worksheet, rd As Worksheet, hr As Long, lastCol As Long, h As Variant

    Set cs = FindContactsSheet(mBook)
    If cs Is Nothing Then
        EnsureSetup = "Can't find the Corp License Number sheet (the one with License Number and Email address columns)."
        Exit Function
    End If
    hr = HeaderRow(cs, Array(H_LICENSE, H_EMAIL))
    lastCol = cs.Cells(hr, cs.Columns.Count).End(xlToLeft).Column
    If FindCol(cs, hr, H_SHORT, lastCol) = 0 Then lastCol = AddHeader(cs, hr, lastCol, "Short Name")
    If FindCol(cs, hr, H_NET, lastCol) = 0 Then lastCol = AddHeader(cs, hr, lastCol, "NET")

    Set rd = FindRefundsSheet(mBook)
    hr = HeaderRow(rd, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(rd, hr)
    For Each h In Array(Array(H_SENT, "MF-12 Sent"), Array(H_SIGNED, "MF-12 Signed"), Array(H_FILED, "Filed Date"), _
                        Array(H_RECEIVED, "Refund Receive Date"))
        If FindCol(rd, hr, CStr(h(0)), lastCol) = 0 Then lastCol = AddHeader(rd, hr, lastCol, CStr(h(1)))
    Next h

    Set ws = SheetNamed(SETUP_SHEET)
    If ws Is Nothing Then MakeSetupSheet Else UpgradeSetupSheet ws
End Function

' A new header at the end of a header row, dressed like the one before it. Hands back its column.
Private Function AddHeader(ws As Worksheet, ByVal hr As Long, ByVal lastCol As Long, ByVal nm As String) As Long
    ws.Cells(hr, lastCol + 1).Value = nm
    LookLike ws.Cells(hr, lastCol), ws.Cells(hr, lastCol + 1)
    AddHeader = lastCol + 1
End Function

Private Sub MakeSetupSheet()
    Dim ws As Worksheet, r As Long
    Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
    ws.Name = SETUP_SHEET
    ws.Columns(1).ColumnWidth = 44
    ws.Columns(2).ColumnWidth = 60
    ws.Columns(3).ColumnWidth = 40
    ws.Range("B8").NumberFormat = "@"     ' a ZIP like 02134 keeps its zero
    ws.Range("B13").NumberFormat = "@"    ' Aug-26 stays a period, not August 26
    ws.Range("A1").Value = "GA refunds setup"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "The first click asks for these and fills them in. Change any here. The macro reads them " & _
        "by the words in column A, so keep those."
    SetupLine ws, 4, "Purchaser name (as it goes on the MF-12)", ""
    SetupLine ws, 5, "Purchaser street address", ""
    SetupLine ws, 6, "Purchaser city", ""
    SetupLine ws, 7, "Purchaser state (2 letters)", ""
    SetupLine ws, 8, "Purchaser ZIP", ""
    SetupLine ws, 9, "Company short name for the email subject", ""
    SetupLine ws, 11, "Packets folder", ""
    ws.Range("C11").Value = "Each period gets its own folder inside it, like 2026-08 AUGUST, and INBOX sits here too."
    SetupLine ws, 12, "MF-12 template file", ""
    ws.Range("C12").Value = "Full path to " & TEMPLATE_NAME & ". Found by itself when it's beside this workbook."
    SetupLine ws, 13, "Oldest period to work", ""
    ws.Range("C13").Value = "Rows from before this period are left alone. Type none to work every row with no Filed Date."
    SetupLine ws, 15, "Email subject", ">MONTH< >YEAR< GA MF-12 Waiver - >US< / >SUPPLIER<"
    SetupLine ws, 16, "Email text", "Attached is our Georgia MF-12 waiver covering >TYPES< for >MONTH< >YEAR<, " & _
        "along with copies of the supporting invoices." & vbLf & _
        "Please sign, notarize, and return each attached MF-12 at your earliest convenience." & vbLf & vbLf & _
        "Thank you very much."
    ws.Range("B16").WrapText = True
    ws.Range("C16").Value = "Alt+Enter starts a new line inside the cell."
    ws.Range("C15").Value = "These get filled in per supplier: >MONTH< >YEAR< >SUPPLIER< >US< >TYPES< >CONTACT<"
    r = 18
    ws.Cells(r, 1).Value = "Refund type (as the report says it)"
    ws.Cells(r, 2).Value = "Folder prefix"
    ws.Cells(r, 3).Value = "Words for the email"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = GRAY
    ws.Cells(r + 1, 1).Value = "6 - Disbursements: Licensed - Tax Unpaid"
    ws.Cells(r + 1, 2).Value = "LIC"
    ws.Cells(r + 1, 3).Value = "licensed customer sales"
    ws.Cells(r + 2, 1).Value = "8 - Disbursements: Federal Government"
    ws.Cells(r + 2, 2).Value = "GOV"
    ws.Cells(r + 2, 3).Value = "government customer sales"
    ws.Range(ws.Cells(r + 3, 1), ws.Cells(r + 6, 3)).Interior.Color = YELLOW
    ws.Cells(r + 8, 1).Value = "A refund type the list doesn't have lands on LOOK. Add it on a yellow line: the words exactly as the report has them, " & _
        "a short prefix for the folders (like DIV), and the words for the email (like diversions)."
    ws.Range("A1:C31").VerticalAlignment = xlTop
    ws.Rows(16).AutoFit
End Sub

' A SETUP sheet from before the running list: its Month folder line becomes Packets folder,
' and an Oldest period line goes in under the template line.
Private Sub UpgradeSetupSheet(ws As Worksheet)
    Dim r As Long, v As String
    If SetupRow(ws, "packets folder") = 0 Then
        r = SetupRow(ws, "month folder")
        If r > 0 Then
            v = CStr(ws.Cells(r, 2).Value)
            ws.Cells(r, 1).Value = "Packets folder"
            ' a pattern like \\server\share\Refunds\{YYYY}-{MM} {MONTH} keeps the part before the first {
            If InStr(v, "{") > 0 Then
                v = Left$(v, InStr(v, "{") - 1)
                If InStrRev(v, "\") > 0 Then v = Left$(v, InStrRev(v, "\") - 1) Else v = ""
                ws.Cells(r, 2).Value = v
            End If
            ws.Cells(r, 3).Value = "Each period gets its own folder inside it, like 2026-08 AUGUST, and INBOX sits here too."
        End If
    End If
    If SetupRow(ws, "oldest period") = 0 Then
        r = SetupRow(ws, "mf-12 template")
        If r > 0 Then
            ws.Rows(r + 1).Insert
            SetupLine ws, r + 1, "Oldest period to work", ""
            ws.Cells(r + 1, 2).NumberFormat = "@"
            ws.Cells(r + 1, 3).Value = "Rows from before this period are left alone. Type none to work every row with no Filed Date."
        End If
    End If
End Sub

' A new header cell dressed like the one beside it.
Private Sub LookLike(src As Range, dst As Range)
    dst.Font.Bold = src.Font.Bold
    dst.Font.Color = src.Font.Color
    If src.Interior.ColorIndex <> xlColorIndexNone Then dst.Interior.Color = src.Interior.Color
    dst.HorizontalAlignment = src.HorizontalAlignment
    dst.WrapText = src.WrapText
    dst.EntireColumn.ColumnWidth = 16
End Sub

Private Sub SetupLine(ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = value
    ws.Cells(r, 2).Interior.Color = YELLOW
End Sub

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
            "(like 2026-08 AUGUST), plus an INBOX for the invoices. The macro makes them all.", vbInformation
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

Private Sub SetSetup(ws As Worksheet, ByVal labelStart As String, ByVal value As String)
    Dim r As Long
    r = SetupRow(ws, labelStart)
    If r = 0 Then Exit Sub
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = value
End Sub

' Returns a list of what's missing, blank when all good.
Private Function ReadSetup() As String
    Dim ws As Worksheet, miss As String, r As Long, t As String, m As Long, y As Long

    Set ws = SheetNamed(SETUP_SHEET)
    mBuyer = SetupValue(ws, "purchaser name")
    mAddress = SetupValue(ws, "purchaser street")
    mCity = SetupValue(ws, "purchaser city")
    mState = UCase$(SetupValue(ws, "purchaser state"))
    mZip = SetupValue(ws, "purchaser zip")
    If IsAllDigits(mZip) And Len(mZip) < 5 Then mZip = Right$("0000" & mZip, 5)
    mUs = SetupValue(ws, "company short name")
    mBase = SetupValue(ws, "packets folder")
    mTemplatePath = SetupValue(ws, "mf-12 template")
    mSubject = SetupValue(ws, "email subject")
    mBody = SetupValue(ws, "email text")
    If Right$(mBase, 1) = "\" Then mBase = Left$(mBase, Len(mBase) - 1)
    mOldest = 0
    t = PeriodText(ws.Cells(SetupRow(ws, "oldest period"), 2).Value)

    If Len(mBuyer) = 0 Then miss = miss & "  Purchaser name" & vbLf
    If Len(mAddress) = 0 Then miss = miss & "  Purchaser street address" & vbLf
    If Len(mCity) = 0 Then miss = miss & "  Purchaser city" & vbLf
    If Len(mState) <> 2 Then miss = miss & "  Purchaser state, 2 letters" & vbLf
    If Not (mZip Like "#####" Or mZip Like "#####-####") Then miss = miss & "  Purchaser ZIP, like 12345 or 12345-6789" & vbLf
    If Len(mUs) = 0 Then miss = miss & "  Company short name for the email subject" & vbLf
    If InStr(mBase, "\") = 0 Then miss = miss & "  Packets folder, a full path like C:\GA Refunds or \\server\share\GA Refunds" & vbLf
    If Len(mTemplatePath) = 0 Then miss = miss & "  MF-12 template file" & vbLf
    If LCase$(t) <> "none" Then
        If ParsePeriod(t, m, y) Then mOldest = y * 100 + m Else miss = miss & "  Oldest period to work, a period like Aug-26, or none" & vbLf
    End If
    If Len(mSubject) = 0 Then miss = miss & "  Email subject" & vbLf
    If Len(mBody) = 0 Then miss = miss & "  Email text" & vbLf

    ' the refund type list
    r = SetupRow(ws, "refund type")
    If r > 0 Then
        r = r + 1
        Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
            t = Trim$(CStr(ws.Cells(r, 2).Value))
            mNType = mNType + 1
            ReDim Preserve mTypeText(1 To mNType): ReDim Preserve mTypePrefix(1 To mNType): ReDim Preserve mTypePhrase(1 To mNType)
            mTypeText(mNType) = Squash(CStr(ws.Cells(r, 1).Value))
            mTypePrefix(mNType) = UCase$(SafeName(t))
            mTypePhrase(mNType) = Trim$(CStr(ws.Cells(r, 3).Value))
            If Len(mTypePrefix(mNType)) = 0 Or Len(mTypePhrase(mNType)) = 0 Then
                miss = miss & "  Refund type list, row " & r & ": needs a folder prefix and words for the email" & vbLf
            End If
            r = r + 1
        Loop
    End If
    If mNType = 0 Then miss = miss & "  The refund type list (under 'Refund type (as the report says it)')" & vbLf
    ReadSetup = miss
End Function

Private Function SetupRow(ws As Worksheet, ByVal labelStart As String) As Long
    Dim r As Long
    For r = 1 To 60
        If Left$(Squash(CStr(ws.Cells(r, 1).Value)), Len(labelStart)) = labelStart Then SetupRow = r: Exit Function
    Next r
End Function

Private Function SetupValue(ws As Worksheet, ByVal labelStart As String) As String
    Dim r As Long
    r = SetupRow(ws, labelStart)
    If r > 0 Then SetupValue = Trim$(CStr(ws.Cells(r, 2).Value))
End Function

Private Function LoadTemplate() As Boolean
    Dim f As Integer, b() As Byte
    If Not FileExists(mTemplatePath) Then Exit Function
    f = FreeFile
    Open mTemplatePath For Binary Access Read As #f
    If LOF(f) < 1000 Then Close #f: Exit Function
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    mTemplate = b
    LoadTemplate = InStrB(1, mTemplate, AnsiB("(MF12V:SELLER:"), vbBinaryCompare) > 0 And _
                   InStrB(1, mTemplate, AnsiB("%MF12A:R21.FUEL:"), vbBinaryCompare) > 0 And _
                   InStrB(1, mTemplate, AnsiB("%MF12ID:"), vbBinaryCompare) > 0
End Function

' ======================================================================================
'  Reading the sheets
' ======================================================================================

Private Function FindRefundsSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(Squash(ws.Name), 11) = "refunds due" Then
            If HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE)) > 0 Then Set FindRefundsSheet = ws: Exit Function
        End If
    Next ws
    For Each ws In wb.Worksheets
        If ws.Name <> SETUP_SHEET And ws.Name <> CHECK_SHEET And ws.Name <> ENTRY_SHEET And ws.Name <> HOWTO_SHEET And _
           Left$(Squash(ws.Name), 16) <> "refunds received" Then
            If HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE, H_PERIOD)) > 0 Then Set FindRefundsSheet = ws: Exit Function
        End If
    Next ws
End Function

Private Function FindReceivedSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In mBook.Worksheets
        If Left$(Squash(ws.Name), 16) = "refunds received" Then Set FindReceivedSheet = ws: Exit Function
    Next ws
End Function

Private Function FindContactsSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetNamedIn(wb, CONTACTS_SHEET)
    If Not ws Is Nothing Then
        If HeaderRow(ws, Array(H_LICENSE, H_EMAIL)) > 0 Then Set FindContactsSheet = ws: Exit Function
    End If
    For Each ws In wb.Worksheets
        If HeaderRow(ws, Array(H_LICENSE, H_EMAIL)) > 0 Then
            If HeaderRow(ws, Array(H_BOL)) = 0 Then Set FindContactsSheet = ws: Exit Function
        End If
    Next ws
End Function

Private Sub LoadContacts()
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long
    Dim cLic As Long, cCorp As Long, cContact As Long, cEmail As Long, cGroup As Long, cShort As Long, cNet As Long
    Dim cExtra As Long, lic As String, i As Long, j As Long, n As String, taken As String

    Set ws = FindContactsSheet(mBook)
    hr = HeaderRow(ws, Array(H_LICENSE, H_EMAIL))
    lastCol = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    cLic = FindCol(ws, hr, H_LICENSE, lastCol)
    cCorp = FindCol(ws, hr, H_CORP, lastCol)
    cContact = FindCol(ws, hr, H_CONTACT, lastCol)
    cEmail = FindCol(ws, hr, H_EMAIL, lastCol)
    cGroup = FindCol(ws, hr, H_GROUP, lastCol)
    cExtra = FindCol(ws, hr, H_EXTRA, lastCol)
    cShort = FindCol(ws, hr, H_SHORT, lastCol)
    cNet = FindCol(ws, hr, H_NET, lastCol)
    lastRow = ws.Cells(ws.Rows.Count, cLic).End(xlUp).Row

    For r = hr + 1 To lastRow
        lic = CellStr(ws.Cells(r, cLic).Value)
        If Len(lic) > 0 Then
            i = FindSup(LicenseKey(lic))
            If i > 0 Then
                AddLook "License " & lic & " is on " & ws.Name & " twice", "Rows " & mSup(i).SheetRow & " and " & r & _
                    ". Only row " & mSup(i).SheetRow & " is used.", "Delete the extra row, or merge them."
            Else
                mNSup = mNSup + 1
                ReDim Preserve mSup(1 To mNSup)
                With mSup(mNSup)
                    .License = lic
                    .LicKey = LicenseKey(lic)
                    .SheetRow = r
                    If cCorp > 0 Then .Corp = CellStr(ws.Cells(r, cCorp).Value)
                    If cContact > 0 Then .Contact = CellStr(ws.Cells(r, cContact).Value)
                    If cEmail > 0 Then .Email = CellStr(ws.Cells(r, cEmail).Value)
                    If cGroup > 0 Then .GroupMail = CellStr(ws.Cells(r, cGroup).Value)
                    If cExtra > 0 Then .Extra = CellStr(ws.Cells(r, cExtra).Value)
                    If cShort > 0 Then .ShortName = UCase$(SafeName(CellStr(ws.Cells(r, cShort).Value)))
                    If cNet > 0 Then
                        n = LCase$(CellStr(ws.Cells(r, cNet).Value))
                        .IsNet = (n = "y" Or n = "yes" Or n = "x" Or n = "net" Or n = "true")
                    End If
                    If Len(.ShortName) > 0 Then taken = taken & "|" & .ShortName & "|"
                End With
            End If
        End If
    Next r

    ' a blank Short Name gets one from the Corp name, written back onto the sheet
    For i = 1 To mNSup
        With mSup(i)
            If Len(.ShortName) = 0 And Len(.Corp) > 0 And cShort > 0 Then
                .ShortName = GuessShort(.Corp, taken)
                If Len(.ShortName) > 0 Then
                    ws.Cells(.SheetRow, cShort).Value = .ShortName
                    taken = taken & "|" & .ShortName & "|"
                    mShortsFilled = mShortsFilled + 1
                End If
            End If
            If Len(.ShortName) = 0 Then .Problem = "No Short Name on " & ws.Name & " row " & .SheetRow
            If Len(.Corp) = 0 Then .Problem = "No Corp name on " & ws.Name & " row " & .SheetRow & " (it goes on the MF-12 as the seller)"
        End With
    Next i

    ' two licenses can't share a short name, or their folders would collide
    For i = 1 To mNSup
        For j = i + 1 To mNSup
            If Len(mSup(i).ShortName) > 0 And mSup(i).ShortName = mSup(j).ShortName Then
                mSup(i).Problem = "Short Name " & mSup(i).ShortName & " is also on row " & mSup(j).SheetRow
                mSup(j).Problem = "Short Name " & mSup(j).ShortName & " is also on row " & mSup(i).SheetRow
            End If
        Next j
    Next i
End Sub

' MANSFIELD from Mansfield Oil Company: the Corp name's first word, or its first two if another
' supplier already has the one word, then a number on the end if those are taken too.
Private Function GuessShort(ByVal corp As String, ByVal taken As String) As String
    Dim w As Variant, s As String, i As Long, c As String, n As Long
    For i = 1 To Len(corp)
        c = UCase$(Mid$(corp, i, 1))
        If c Like "[A-Z0-9]" Then s = s & c Else s = s & " "
    Next i
    s = Trim$(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    If Len(s) = 0 Then Exit Function
    w = Split(s, " ")
    i = 0
    If UBound(w) > 0 And w(0) = "THE" Then i = 1
    s = w(i)
    If InStr(taken, "|" & s & "|") > 0 And UBound(w) > i Then s = s & " " & w(i + 1)
    If InStr(taken, "|" & s & "|") > 0 Then
        n = 2
        Do While InStr(taken, "|" & s & " " & n & "|") > 0: n = n + 1: Loop
        s = s & " " & n
    End If
    GuessShort = s
End Function

' Every row on Refunds Due that's still open: no Filed Date, no Refund Receive Date, and not
' from before the oldest period on SETUP.
Private Sub LoadRows()
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long, data As Variant, ok As Boolean
    Dim cType As Long, cPer As Long, cCorp As Long, cLic As Long, cBol As Long, cInv As Long, cInvDate As Long
    Dim cGal As Long, cRef As Long, cAllow As Long, cTtl As Long, cFuel As Long, cSent As Long, cSigned As Long
    Dim cFiled As Long, cRecv As Long, key As Long, odd As Object, k As Variant, d As Date, per As String

    Set ws = FindRefundsSheet(mBook)
    hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(ws, hr)
    cType = FindCol(ws, hr, H_TYPE, lastCol)
    cPer = FindCol(ws, hr, H_PERIOD, lastCol)
    cCorp = FindCol(ws, hr, H_CORP, lastCol)
    cLic = FindCol(ws, hr, H_LICENSE, lastCol)
    cBol = FindCol(ws, hr, H_BOL, lastCol)
    cInv = FindCol(ws, hr, H_INV, lastCol)
    cInvDate = FindCol(ws, hr, H_DATE, lastCol)
    cGal = FindCol(ws, hr, H_GALLONS, lastCol)
    cRef = FindCol(ws, hr, H_REFUND, lastCol)
    cAllow = FindCol(ws, hr, H_ALLOW, lastCol)
    cTtl = FindCol(ws, hr, H_TTL, lastCol)
    cFuel = FindCol(ws, hr, H_FUEL, lastCol)
    cSent = FindCol(ws, hr, H_SENT, lastCol)
    cSigned = FindCol(ws, hr, H_SIGNED, lastCol)
    cFiled = FindCol(ws, hr, H_FILED, lastCol)
    cRecv = FindCol(ws, hr, H_RECEIVED, lastCol)
    If cType = 0 Or cPer = 0 Or cCorp = 0 Or cLic = 0 Or cBol = 0 Or cInv = 0 Or cInvDate = 0 Or cGal = 0 Or _
       cRef = 0 Or cAllow = 0 Or cTtl = 0 Or cFuel = 0 Then
        Err.Raise vbObjectError + 10, , "Refunds Due is missing a column. It needs: Refund Type (or Schedule), Period, Corp, " & _
            "License Number, BOL, INV, Inv Date, Gallons, Refund, Allowance, TTL Refund, Fuel."
    End If
    lastRow = LastDataRow(ws, hr, lastCol)
    If lastRow <= hr Then Exit Sub
    data = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lastRow, lastCol)).Value
    Set odd = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(data, 1)
        ' a blank row or a note row has no BOL and no license; a filed or paid one is done
        If Len(CellStr(data(r, cBol))) + Len(CellStr(data(r, cLic))) > 0 And _
           Len(CellStr(data(r, cFiled))) + Len(CellStr(data(r, cRecv))) = 0 Then
            per = PeriodText(data(r, cPer))
            key = PeriodKey(per)
            If key = 0 Then
                d = ToDate(data(r, cInvDate), ok)
                If ok Then
                    key = Year(d) * 100 + Month(d)
                    If odd.Exists(per) Then odd(per) = odd(per) & ", " & (hr + r) Else odd.Add per, CStr(hr + r)
                End If
            End If
            If mOldest > 0 And key > 0 And key < mOldest Then
                mLeftAlone = mLeftAlone + 1
            Else
                mNRow = mNRow + 1
                ReDim Preserve mRow(1 To mNRow)
                With mRow(mNRow)
                    .SheetRow = hr + r
                    .PerKey = key
                    .TypeText = CellStr(data(r, cType))
                    .Corp = CellStr(data(r, cCorp))
                    .License = CellStr(data(r, cLic))
                    .Bol = CellStr(data(r, cBol))
                    .BolK = BolKey(.Bol)
                    .Inv = CellStr(data(r, cInv))
                    .InvDate = ToDate(data(r, cInvDate), ok)
                    If Not ok Then .Problem = "Inv Date isn't a date"
                    .Gallons = ToNum(data(r, cGal), ok)
                    If Not ok Or .Gallons <= 0 Then .Problem = "Gallons is blank or not a number"
                    .Refund = ToNum(data(r, cRef), ok)
                    If Not ok Or .Refund <= 0 Then .Problem = "Refund is blank or not a number"
                    .Allow = ToNum(data(r, cAllow), ok)
                    .Ttl = ToNum(data(r, cTtl), ok)
                    .FuelText = CellStr(data(r, cFuel))
                    .FuelOpt = FuelOption(.FuelText)
                    If Len(.FuelOpt) = 0 Then .Problem = "Fuel '" & .FuelText & "' isn't one of the MF-12's fuel choices"
                    If Len(.Inv) = 0 Then .Problem = "No INV (supplier invoice number)"
                    If Len(.Bol) = 0 Then .Problem = "No BOL"
                    If key = 0 Then .Problem = "Period isn't a month like Aug-26, and Inv Date isn't a date"
                    If cSent > 0 Then
                        .HasSent = Len(CellStr(data(r, cSent))) > 0
                        .SentOn = ToDate(data(r, cSent), ok)
                    End If
                    If cSigned > 0 Then .HasSigned = Len(CellStr(data(r, cSigned))) > 0
                End With
            End If
        End If
    Next r
    For Each k In odd.Keys
        AddLook "Couldn't read the period " & IIf(Len(k) > 0, "'" & k & "'", "(it's blank)"), "Row(s) " & odd(k) & _
            ". Went by each row's invoice date.", _
            "Nothing to do if that's right. If not, fix Period on those rows."
    Next k
End Sub

' The last row with anything in a header column, for reading a sheet down to its end.
Private Function LastDataRow(ws As Worksheet, ByVal hr As Long, ByVal lastCol As Long) As Long
    Dim c As Long, r As Long
    LastDataRow = hr
    For c = 1 To lastCol
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > LastDataRow Then LastDataRow = r
    Next c
End Function

' The periods of the rows still open on Refunds Due, oldest first, for the first click's question.
Private Function OpenPeriods() As Collection
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long, data As Variant
    Dim cPer As Long, cBol As Long, cLic As Long, cFiled As Long, cRecv As Long, key As Long, i As Long, done As Boolean
    Dim out As New Collection

    Set OpenPeriods = out
    Set ws = FindRefundsSheet(mBook)
    hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(ws, hr)
    cPer = FindCol(ws, hr, H_PERIOD, lastCol)
    cBol = FindCol(ws, hr, H_BOL, lastCol)
    cLic = FindCol(ws, hr, H_LICENSE, lastCol)
    cFiled = FindCol(ws, hr, H_FILED, lastCol)
    cRecv = FindCol(ws, hr, H_RECEIVED, lastCol)
    lastRow = LastDataRow(ws, hr, lastCol)
    If cPer = 0 Or cFiled = 0 Or cRecv = 0 Or lastRow <= hr Then Exit Function
    data = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lastRow, lastCol)).Value
    For r = 1 To UBound(data, 1)
        If Len(CellStr(data(r, cBol))) + Len(CellStr(data(r, cLic))) > 0 And _
           Len(CellStr(data(r, cFiled))) + Len(CellStr(data(r, cRecv))) = 0 Then
            key = PeriodKey(PeriodText(data(r, cPer)))
            If key > 0 Then
                done = False
                For i = 1 To out.Count
                    If out(i) = key Then done = True: Exit For
                    If out(i) > key Then out.Add key, Before:=i: done = True: Exit For
                Next i
                If Not done Then out.Add key
            End If
        End If
    Next r
End Function

Private Sub AssignPackets()
    Dim i As Long, t As Long, s As Long, p As Long, m As Long, key As String, probs As Object, k As Variant

    Set probs = CreateObject("Scripting.Dictionary")
    For i = 1 To mNRow
        With mRow(i)
            If Len(.License) = 0 Then
                .Problem = "No License Number on the row"
            Else
                .Sup = FindSup(LicenseKey(.License))
                If .Sup = 0 Then
                    .Problem = "License " & .License & " (" & .Corp & ") isn't on " & CONTACTS_SHEET
                Else
                    mSup(.Sup).NRows = mSup(.Sup).NRows + 1
                    If mSup(.Sup).IsNet Then .Amount = .Ttl Else .Amount = .Refund
                    If mSup(.Sup).IsNet And Abs(.Refund - .Allow - .Ttl) > 0.015 Then
                        .Problem = "TTL Refund isn't Refund minus Allowance, and this supplier is NET"
                    End If
                End If
            End If
            t = FindType(.TypeText)
            If t = 0 Then
                .Problem = "Refund type '" & .TypeText & "' isn't on the SETUP list"
            Else
                .Prefix = mTypePrefix(t)
            End If

            ' one email per supplier per period
            m = 0
            If .Sup > 0 And .PerKey > 0 Then
                m = FindMailing(.Sup, .PerKey)
                If m = 0 Then
                    mNMail = mNMail + 1
                    ReDim Preserve mMail(1 To mNMail)
                    m = mNMail
                    mMail(m).Sup = .Sup
                    mMail(m).PerKey = .PerKey
                End If
                mMail(m).NRows = mMail(m).NRows + 1
                If .HasSent Then
                    mMail(m).IsSent = True
                    If .SentOn > mMail(m).SentOn Then mMail(m).SentOn = .SentOn
                End If
            End If

            ' a packet is a period, a refund type and a supplier with a short name
            If m > 0 And t > 0 Then
                If Len(mSup(.Sup).ShortName) > 0 Then
                    key = .PerKey & "|" & .Prefix & "-" & mSup(.Sup).ShortName
                    p = FindPkt(key)
                    If p = 0 Then
                        mNPkt = mNPkt + 1
                        ReDim Preserve mPkt(1 To mNPkt)
                        p = mNPkt
                        mPkt(p).Key = key
                        mPkt(p).PName = .Prefix & "-" & mSup(.Sup).ShortName
                        mPkt(p).Prefix = .Prefix
                        mPkt(p).PerKey = .PerKey
                        mPkt(p).Sup = .Sup
                        mPkt(p).Mail = m
                        mPkt(p).Folder = MonthFolder(.PerKey) & "\" & mPkt(p).PName
                    End If
                    .Pkt = p
                    mPkt(p).NRows = mPkt(p).NRows + 1
                    mPkt(p).Gallons = mPkt(p).Gallons + .Gallons
                    mPkt(p).Refund = mPkt(p).Refund + .Refund
                    mPkt(p).Allow = mPkt(p).Allow + .Allow
                    mPkt(p).Ttl = mPkt(p).Ttl + .Ttl
                    mPkt(p).Amount = mPkt(p).Amount + .Amount
                End If
            End If

            If Len(.Problem) > 0 Then
                If .Pkt > 0 Then
                    If Len(mPkt(.Pkt).Problem) = 0 Then mPkt(.Pkt).Problem = "row " & .SheetRow & ": " & .Problem
                End If
                If m > 0 Then
                    If Len(mMail(m).Held) = 0 Then mMail(m).Held = "a row needs a look (row " & .SheetRow & ")"
                End If
                If probs.Exists(.Problem) Then
                    probs(.Problem) = probs(.Problem) & ", " & .SheetRow
                Else
                    probs.Add .Problem, CStr(.SheetRow)
                End If
            End If
        End With
    Next i

    For Each k In probs.Keys
        AddLook "Refunds Due: " & k, "Row(s) " & probs(k), RowFix(CStr(k))
    Next k
    For s = 1 To mNSup
        If mSup(s).NRows > 0 And Len(mSup(s).Problem) > 0 Then
            AddLook mSup(s).Problem, "License " & mSup(s).License & ", " & mSup(s).NRows & " open row(s)", "Fill it in on " & CONTACTS_SHEET & "."
        End If
    Next s
    For p = 1 To mNPkt
        mPkt(p).NForms = Int((mPkt(p).NRows + ROWS_PER_FORM - 1) / ROWS_PER_FORM)
    Next p
End Sub

Private Function RowFix(ByVal problem As String) As String
    If InStr(problem, "isn't on " & CONTACTS_SHEET) > 0 Then
        RowFix = "Add the license to " & CONTACTS_SHEET & " with its contact, Short Name and NET."
    ElseIf InStr(problem, "SETUP list") > 0 Then
        RowFix = "Add it to the refund type list on SETUP (prefix like DIV, words like diversions)."
    ElseIf InStr(problem, "fuel choices") > 0 Then
        RowFix = "Change the Fuel cell to one of: Aviation, Bio Diesel, Fuel Oil Clear, Fuel Oil Dyed, Gasohol, Gasoline, Jet Fuel, LP, Special Fuel."
    Else
        RowFix = "Fix the row on Refunds Due, then run it again."
    End If
End Function

' The Corp name as typed on Corp License Number goes onto every open matched row of Refunds
' Due, so the report's own spelling of a supplier never shows anywhere you work from. A row
' whose license isn't on Corp License Number keeps what was pasted.
Private Sub WriteCorpNames()
    Dim ws As Worksheet, hr As Long, c As Long, i As Long, nm As String
    Set ws = FindRefundsSheet(mBook)
    If ws Is Nothing Then Exit Sub
    hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
    c = FindCol(ws, hr, H_CORP, LastHeaderCol(ws, hr))
    If c = 0 Then Exit Sub
    For i = 1 To mNRow
        If mRow(i).Sup > 0 Then
            nm = mSup(mRow(i).Sup).Corp
            If Len(nm) > 0 And mRow(i).Corp <> nm Then
                ws.Cells(mRow(i).SheetRow, c).Value = nm
                mRow(i).Corp = nm
            End If
        End If
    Next i
End Sub

' ======================================================================================
'  Rows that are paid, and rows that are filed
' ======================================================================================

' Rows with a Refund Receive Date move to the bottom of Refunds Received, matched column by
' column on the headers, and come off Refunds Due. Hands back how many moved.
Private Function MoveReceived() As Long
    Dim rd As Worksheet, rr As Worksheet, hr As Long, rhr As Long, lastCol As Long, rLast As Long, cRecv As Long
    Dim lastRow As Long, r As Long, c As Long, dest As Long, map() As Long, moving() As Long, n As Long, i As Long, nm As String

    Set rd = FindRefundsSheet(mBook)
    hr = HeaderRow(rd, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(rd, hr)
    cRecv = FindCol(rd, hr, H_RECEIVED, lastCol)
    If cRecv = 0 Then Exit Function
    lastRow = LastDataRow(rd, hr, lastCol)
    For r = hr + 1 To lastRow
        If Len(CellStr(rd.Cells(r, cRecv).Value)) > 0 Then
            n = n + 1
            ReDim Preserve moving(1 To n)
            moving(n) = r
        End If
    Next r
    If n = 0 Then Exit Function

    Set rr = FindReceivedSheet()
    If rr Is Nothing Then
        Set rr = mBook.Worksheets.Add(After:=rd)
        rr.Name = RECEIVED_SHEET
        rd.Rows(hr).Copy rr.Rows(1)
        rhr = 1
    Else
        rhr = HeaderRow(rr, Array(H_BOL))
        If rhr = 0 Then
            If Application.WorksheetFunction.CountA(rr.UsedRange) = 0 Then rhr = 1 Else rhr = rr.UsedRange.Row + rr.UsedRange.Rows.Count + 1
            rd.Rows(hr).Copy rr.Rows(rhr)
        End If
    End If
    ' a column Refunds Received doesn't have (FUEL, say) isn't carried over; its layout stays yours
    rLast = LastHeaderCol(rr, rhr)
    ReDim map(1 To lastCol)
    For c = 1 To lastCol
        nm = Squash(CellStr(rd.Cells(hr, c).Value))
        If Len(nm) > 0 Then map(c) = FindCol(rr, rhr, nm, rLast)
    Next c
    dest = LastDataRow(rr, rhr, rLast) + 1
    For i = 1 To n
        For c = 1 To lastCol
            If map(c) > 0 Then
                rr.Cells(dest, map(c)).NumberFormat = rd.Cells(moving(i), c).NumberFormat
                rr.Cells(dest, map(c)).Value = rd.Cells(moving(i), c).Value
            End If
        Next c
        dest = dest + 1
    Next i
    For i = n To 1 Step -1
        rd.Rows(moving(i)).Delete
    Next i
    MoveReceived = n
End Function

' Stamps today's date as Filed Date on every row of every packet showing on ENTRY, after a
' yes/no box names them. Hands back how many packets it marked.
Private Function MarkFiled() As Long
    Dim ws As Worksheet, hr As Long, cFiled As Long, p As Long, i As Long, n As Long, names As String, ix As Variant

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

    Set ws = FindRefundsSheet(mBook)
    hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
    cFiled = FindCol(ws, hr, H_FILED, LastHeaderCol(ws, hr))
    For i = 1 To mNRow
        p = mRow(i).Pkt
        If p > 0 Then
            If mPkt(p).Stage = ST_BACK Then StampCell ws.Cells(mRow(i).SheetRow, cFiled), Date
        End If
    Next i
    MarkFiled = n
End Function

' MF-12 Sent from the email in Sent Items, MF-12 Signed from the NOTARIZED copy's date, on
' each row where it's still blank. A date you typed is never changed.
Private Sub WriteStamps()
    Dim ws As Worksheet, hr As Long, lastCol As Long, cSent As Long, cSigned As Long, i As Long, m As Long, p As Long
    Set ws = FindRefundsSheet(mBook)
    hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(ws, hr)
    cSent = FindCol(ws, hr, H_SENT, lastCol)
    cSigned = FindCol(ws, hr, H_SIGNED, lastCol)
    For i = 1 To mNRow
        p = mRow(i).Pkt
        If p > 0 Then
            m = mPkt(p).Mail
            If cSent > 0 And Not mRow(i).HasSent And mMail(m).SentOn > 0 Then
                StampCell ws.Cells(mRow(i).SheetRow, cSent), mMail(m).SentOn
                mRow(i).HasSent = True
            End If
            If cSigned > 0 And Not mRow(i).HasSigned And mPkt(p).Notarized > 0 Then
                StampCell ws.Cells(mRow(i).SheetRow, cSigned), mPkt(p).Notarized
                mRow(i).HasSigned = True
            End If
        End If
    Next i
End Sub

Private Sub StampCell(c As Range, ByVal d As Date)
    c.NumberFormat = "m/d/yyyy"
    c.Value = DateSerial(Year(d), Month(d), Day(d))
End Sub

' ======================================================================================
'  Folders and the INBOX
' ======================================================================================

' A period's folder under the packets folder: 2026-08 AUGUST.
Private Function MonthFolder(ByVal perKey As Long) As String
    MonthFolder = mBase & "\" & (perKey \ 100) & "-" & Format$(perKey Mod 100, "00") & " " & UCase$(MonthName(perKey Mod 100))
End Function

Private Function InboxFolder() As String
    InboxFolder = mBase & "\INBOX"
End Function

Private Sub MakeFolders()
    Dim p As Long
    MkDirs mBase
    MkDirs InboxFolder()
    For p = 1 To mNPkt
        MkDirs MonthFolder(mPkt(p).PerKey) & "\MFD-33"
        MkDirs mPkt(p).Folder
        MkDirs mPkt(p).Folder & "\UPLOAD"
    Next p
End Sub

' Files each "<BOL> VI" / "<BOL> CI" into its packet folder under its full name, whatever period it's from.
Private Sub SortInbox()
    Dim files As Variant, i As Long, f As String, stem As String, ext As String, bolPart As String, tag As String, sq As String
    Dim r As Long, pk As Long, twoPkts As Boolean, waiting As Long, invs As String, bolShown As String, target As String

    files = ListFiles(InboxFolder())
    If IsEmpty(files) Then Exit Sub
    For i = 1 To UBound(files)
        f = files(i)
        If Not IsJunkFile(f) Then
            SplitExt f, stem, ext
            If Not ParseInboxName(stem, bolPart, tag, sq) Then
                AddLook "INBOX: can't read the name " & f, "It stays in INBOX.", "Name it like 12345 VI or 12345 CI."
                mInboxLeft = mInboxLeft + 1
            Else
                pk = 0: twoPkts = False: waiting = 0: invs = "": bolShown = ""
                For r = 1 To mNRow
                    If mRow(r).BolK = BolKey(bolPart) Then
                        If mRow(r).Pkt = 0 Then
                            waiting = mRow(r).SheetRow
                        ElseIf pk = 0 Then
                            pk = mRow(r).Pkt
                        ElseIf pk <> mRow(r).Pkt Then
                            twoPkts = True
                        End If
                        If InStr("+" & invs & "+", "+" & mRow(r).Inv & "+") = 0 Then invs = invs & IIf(Len(invs) > 0, "+", "") & mRow(r).Inv
                        bolShown = mRow(r).Bol
                    End If
                Next r
                mInboxLeft = mInboxLeft + 1
                If Len(bolShown) = 0 Then
                    AddLook "INBOX: no open row has BOL " & bolPart, f & " stays in INBOX.", _
                        "Check the BOL. Its row may be filed already, or from before the Oldest period to work on SETUP."
                ElseIf waiting > 0 Then
                    AddLook "INBOX: " & f & " waits for row " & waiting, "That row needs a look first (see above).", _
                        "Fix the row and run it again."
                ElseIf twoPkts Then
                    AddLook "INBOX: BOL " & bolPart & " is on rows in two packets", f & " stays in INBOX.", "File this one by hand."
                Else
                    target = mPkt(pk).Folder & "\" & mPkt(pk).PName & " " & SafeName(bolShown) & " INV" & SafeName(invs) & _
                             " " & tag & IIf(Len(sq) > 0, " " & sq, "") & ext
                    If FileExists(target) Or Not FolderExists(mPkt(pk).Folder) Then
                        AddLook "INBOX: " & f & " is already filed", "As " & Mid$(target, InStrRev(target, "\") + 1) & ". The INBOX copy was left alone.", _
                            "Compare the two and delete the one you don't want."
                    Else
                        Name InboxFolder() & "\" & f As target
                        mInboxLeft = mInboxLeft - 1
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub CountInbox()
    Dim files As Variant, i As Long
    files = ListFiles(InboxFolder())
    If IsEmpty(files) Then Exit Sub
    For i = 1 To UBound(files)
        If Not IsJunkFile(CStr(files(i))) Then mInboxLeft = mInboxLeft + 1
    Next i
    If mInboxLeft > 0 Then AddLook "INBOX still has " & mInboxLeft & " file(s)", InboxFolder(), _
        "Click Build packets to file them."
End Sub

' "12345 VI", "12345 VI 2", "12345VI". The BOL, the kind of copy, and an optional number.
Private Function ParseInboxName(ByVal stem As String, ByRef bolPart As String, ByRef tag As String, ByRef sq As String) As Boolean
    Dim t As Variant, s As String
    s = Trim$(Replace(stem, "_", " "))
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    t = Split(s, " ")
    bolPart = "": tag = "": sq = ""
    Select Case UBound(t)
        Case 0
            If Len(s) > 2 And (UCase$(Right$(s, 2)) = "VI" Or UCase$(Right$(s, 2)) = "CI") Then
                bolPart = Left$(s, Len(s) - 2): tag = UCase$(Right$(s, 2))
            End If
        Case 1
            bolPart = t(0): tag = UCase$(t(1))
        Case 2
            If IsAllDigits(CStr(t(2))) Then bolPart = t(0): tag = UCase$(t(1)): sq = t(2)
    End Select
    If Len(bolPart) = 0 Or Len(tag) = 0 Or Len(tag) > 12 Then Exit Function
    If Not IsAllLetters(tag) Then Exit Function
    ParseInboxName = True
End Function

' ======================================================================================
'  What's in each packet folder
' ======================================================================================

Private Sub CheckFiles()
    Dim p As Long, r As Long, files As Variant, i As Long, j As Long, stem As String, ext As String
    Dim want As String, stray As String, formName As String, fp As String, onDisk As Long, bad As String, pre As String
    Dim d As Date

    For p = 1 To mNPkt
        files = ListFiles(mPkt(p).Folder)
        mPkt(p).VIMissing = ""
        mPkt(p).Notarized = 0
        For r = 1 To mNRow
            If mRow(r).Pkt = p Then
                mRow(r).HasVI = False
                pre = UCase$(mPkt(p).PName & " " & SafeName(mRow(r).Bol) & " ")
                If Not IsEmpty(files) Then
                    For i = 1 To UBound(files)
                        SplitExt CStr(files(i)), stem, ext
                        If UCase$(Left$(stem, Len(pre))) = pre Then
                            If FileTag(stem) = "VI" Then mRow(r).HasVI = True
                        End If
                    Next i
                End If
                If Not mRow(r).HasVI Then
                    If InStr(", " & mPkt(p).VIMissing & ",", ", " & mRow(r).Bol & ",") = 0 Then
                        mPkt(p).VIMissing = mPkt(p).VIMissing & IIf(Len(mPkt(p).VIMissing) > 0, ", ", "") & mRow(r).Bol
                    End If
                End If
            End If
        Next r
        If Len(mPkt(p).VIMissing) > 0 And FolderExists(mPkt(p).Folder) Then
            AddLook PeriodLabel(mPkt(p).PerKey) & " " & mPkt(p).PName & ": no VI yet for BOL " & mPkt(p).VIMissing, "The " & mPkt(p).PName & " folder", _
                "Pull the vendor invoice, save it in INBOX as '<BOL> VI', click Build packets."
        End If

        ' the notarized copies that have come back: the newest one's date
        If Not IsEmpty(files) Then
            For i = 1 To UBound(files)
                If InStr(UCase$(CStr(files(i))), "NOTARIZED") > 0 Then
                    d = FileDateTime(mPkt(p).Folder & "\" & files(i))
                    If d > mPkt(p).Notarized Then mPkt(p).Notarized = d
                End If
            Next i
        End If

        ' the MF-12s: the ones it expects, and any other MF-12 lying there
        onDisk = 0: bad = "": stray = ""
        For j = 1 To mPkt(p).NForms
            formName = FormFileName(p, j)
            If FileExists(mPkt(p).Folder & "\" & formName) Then
                onDisk = onDisk + 1
                fp = FileFingerprint(mPkt(p).Folder & "\" & formName)
                If fp <> FormFingerprint(p, j) Then bad = bad & IIf(Len(bad) > 0, ", ", "") & formName
            End If
        Next j
        If Not IsEmpty(files) Then
            For i = 1 To UBound(files)
                want = UCase$(CStr(files(i)))
                If InStr(want, " MF-12") > 0 And InStr(want, "NOTARIZED") = 0 And Right$(want, 4) = ".PDF" Then
                    If Not IsExpectedForm(p, CStr(files(i))) Then stray = stray & IIf(Len(stray) > 0, ", ", "") & files(i)
                End If
            Next i
        End If
        If Len(stray) > 0 Then
            mPkt(p).FormState = "OLD MF-12 IN FOLDER"
            AddLook mPkt(p).PName & ": an MF-12 in the folder doesn't fit this run", stray, _
                "The number of forms changed, or it's an old copy. Move it out of the folder and click Build packets."
        ElseIf Len(bad) > 0 Then
            mPkt(p).FormState = "MF-12 OUT OF DATE"
            AddLook mPkt(p).PName & ": the MF-12 doesn't match the rows anymore", bad, _
                "The report changed since it was made (or it was edited by hand). Move it out of the folder and click Build packets to make a fresh one."
        ElseIf onDisk = mPkt(p).NForms Then
            mPkt(p).FormState = "KEPT"
        ElseIf onDisk = 0 Then
            mPkt(p).FormState = "MISSING"
        Else
            mPkt(p).FormState = "SOME FORMS MISSING"
            AddLook mPkt(p).PName & ": only " & onDisk & " of " & mPkt(p).NForms & " MF-12 forms are there", "The " & mPkt(p).PName & " folder", _
                "Move the MF-12s out of the folder and click Build packets to make them all fresh."
        End If
    Next p
End Sub

Private Function IsExpectedForm(ByVal p As Long, ByVal f As String) As Boolean
    Dim j As Long
    For j = 1 To mPkt(p).NForms
        If UCase$(f) = UCase$(FormFileName(p, j)) Then IsExpectedForm = True: Exit Function
    Next j
End Function

Private Function FormFileName(ByVal p As Long, ByVal formNo As Long) As String
    FormFileName = mPkt(p).PName & " " & UCase$(MonthName(mPkt(p).PerKey Mod 100)) & " " & (mPkt(p).PerKey \ 100) & " MF-12" & _
        IIf(mPkt(p).NForms > 1, " (" & formNo & " of " & mPkt(p).NForms & ")", "") & ".pdf"
End Function

' ======================================================================================
'  The MF-12
' ======================================================================================

Private Sub WriteForms()
    Dim p As Long, j As Long
    For p = 1 To mNPkt
        If Len(mPkt(p).Problem) = 0 And mPkt(p).FormState = "MISSING" And FolderExists(mPkt(p).Folder) Then
            For j = 1 To mPkt(p).NForms
                mStage = "filling the MF-12 for " & mPkt(p).PName
                WriteBytes mPkt(p).Folder & "\" & FormFileName(p, j), FillForm(p, j)
            Next j
            mPkt(p).FormState = "MADE"
        End If
    Next p
End Sub

' The packet's rows in the order they go on the form: by invoice date, then BOL.
Private Function PacketRows(ByVal p As Long) As Variant
    Dim idx() As Long, n As Long, i As Long, j As Long, t As Long
    For i = 1 To mNRow
        If mRow(i).Pkt = p Then n = n + 1: ReDim Preserve idx(1 To n): idx(n) = i
    Next i
    For i = 2 To n
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Not RowBefore(t, idx(j)) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    PacketRows = idx
End Function

Private Function RowBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If mRow(a).InvDate <> mRow(b).InvDate Then
        RowBefore = mRow(a).InvDate < mRow(b).InvDate
    ElseIf mRow(a).BolK <> mRow(b).BolK Then
        RowBefore = mRow(a).BolK < mRow(b).BolK
    Else
        RowBefore = mRow(a).SheetRow < mRow(b).SheetRow
    End If
End Function

Private Function FormValues(ByVal p As Long, ByVal formNo As Long) As Object
    Dim v As Object, ix As Variant, i As Long, slot As Long, k As String, s As Long
    Set v = CreateObject("Scripting.Dictionary")
    s = mPkt(p).Sup
    v("SELLER") = Array(UCase$(mSup(s).Corp), UCase$(mSup(s).Corp))
    v("LICENSE") = Array(UCase$(mSup(s).License), UCase$(mSup(s).License))
    v("BUYER") = Array(UCase$(mBuyer), UCase$(mBuyer))
    v("ADDRESS") = Array(UCase$(mAddress), UCase$(mAddress))
    v("CITY") = Array(UCase$(mCity), UCase$(mCity))
    v("STATE") = Array(mState, mState)
    v("ZIP") = Array(mZip, mZip)
    ix = PacketRows(p)
    For i = (formNo - 1) * ROWS_PER_FORM + 1 To formNo * ROWS_PER_FORM
        If i > UBound(ix) Then Exit For
        slot = i - (formNo - 1) * ROWS_PER_FORM
        k = "R" & Format$(slot, "00") & "."
        With mRow(ix(i))
            v(k & "DATE") = Array(Format$(.InvDate, "mm\/dd\/yyyy"), Format$(.InvDate, "mm\/dd\/yyyy"))
            v(k & "INV") = Array(.Inv, .Inv)
            v(k & "GAL") = Array(PdfNum(.Gallons, 3), Format$(.Gallons, "#,##0"))
            v(k & "TAX") = Array(PdfNum(.Amount, 2), Format$(.Amount, "#,##0.00"))
            v(k & "FUEL") = Array(.FuelOpt, .FuelOpt)
        End With
    Next i
    Set FormValues = v
End Function

' A short tag for everything that goes on one form, so a later run can tell if it changed.
Private Function FormFingerprint(ByVal p As Long, ByVal formNo As Long) As String
    Dim v As Object, k As Variant, h As Double, t As String, i As Long, n As Long
    Set v = FormValues(p, formNo)
    For Each k In v.Keys
        If Left$(k, 1) = "R" Then n = n + 1
        t = k & "=" & v(k)(0) & ";"
        For i = 1 To Len(t)
            h = h * 131 + AscW(Mid$(t, i, 1))
            h = h - Int(h / 4294967291#) * 4294967291#
        Next i
    Next k
    FormFingerprint = "v1-" & (n \ 5) & "-" & Format$(h, "0")
End Function

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

' A number the way PDF wants it: dot for the decimal, no thousands commas, no trailing zeros past what's asked.
Private Function PdfNum(ByVal x As Double, ByVal places As Long) As String
    Dim s As String
    s = Replace(Format$(x, "0." & String$(places, "0")), ",", ".")
    If places = 3 Then
        Do While Right$(s, 1) = "0": s = Left$(s, Len(s) - 1): Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If
    PdfNum = s
End Function

Private Function FileFingerprint(ByVal path As String) As String
    Dim f As Integer, b() As Byte, s As String, pos As Long, e As Long, c As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) = 0 Then Close #f: Exit Function
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    s = b
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
Private Sub CheckSent(ByVal openIt As Boolean)
    Dim ol As Object, ns As Object, sentBox As Object, draftBox As Object, m As Long, subj As String, found As Object
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
        subj = FillWords(mSubject, m)
        If Not mMail(m).IsSent Then
            If Len(mTestDraftDir) > 0 Then
                If FileExists(mTestDraftDir & "\sent\" & SafeName(subj) & ".txt") Then
                    mMail(m).IsSent = True
                    mMail(m).SentOn = FileDateTime(mTestDraftDir & "\sent\" & SafeName(subj) & ".txt")
                End If
            Else
                Set found = FindMail(sentBox, subj)
                If Not found Is Nothing Then
                    mMail(m).IsSent = True
                    mMail(m).SentOn = found.SentOn
                End If
            End If
        End If
        If Not mMail(m).IsSent Then
            If Len(mTestDraftDir) > 0 Then
                mMail(m).InDrafts = FileExists(TestDraftPath(m))
            Else
                mMail(m).InDrafts = Not FindMail(draftBox, subj) Is Nothing
            End If
        End If
    Next m
End Sub

Private Sub DecideReady()
    Dim m As Long, s As Long, p As Long, why As String
    ' a notarized copy back means the email went, even if nothing says when
    For p = 1 To mNPkt
        If mPkt(p).Notarized > 0 Then mMail(mPkt(p).Mail).IsSent = True
    Next p
    For m = 1 To mNMail
        s = mMail(m).Sup
        why = mMail(m).Held
        If Len(why) = 0 And Len(mSup(s).Problem) > 0 Then why = mSup(s).Problem
        If Len(why) = 0 And Len(mSup(s).Email) + Len(mSup(s).GroupMail) + Len(mSup(s).Extra) = 0 Then
            why = "no email address on " & CONTACTS_SHEET
            AddLook mSup(s).ShortName & ": no email address", CONTACTS_SHEET & " row " & mSup(s).SheetRow, "Fill in Email address or Group Email Address."
        End If
        For p = 1 To mNPkt
            If mPkt(p).Mail = m And Len(why) = 0 Then
                If Len(mPkt(p).Problem) > 0 Then
                    why = mPkt(p).PName & ": " & mPkt(p).Problem
                ElseIf Len(mPkt(p).VIMissing) > 0 Then
                    why = mPkt(p).PName & ": waiting on VI for BOL " & mPkt(p).VIMissing
                ElseIf mPkt(p).FormState <> "MADE" And mPkt(p).FormState <> "KEPT" Then
                    why = mPkt(p).PName & ": MF-12 " & LCase$(mPkt(p).FormState)
                End If
            End If
        Next p
        mMail(m).Held = why
        If mMail(m).IsSent Then
            mMail(m).Status = "SENT" & IIf(mMail(m).SentOn > 0, " " & Format$(mMail(m).SentOn, "m/d/yyyy"), "")
        ElseIf Len(why) > 0 Then
            mMail(m).Status = "HELD: " & why
        ElseIf mMail(m).InDrafts Then
            mMail(m).Status = "IN DRAFTS, not sent yet"
        Else
            mMail(m).Status = "READY"
            ListAttachments m
            If mMail(m).MB > BIG_EMAIL_MB Then mMail(m).Status = "READY, but " & Format$(mMail(m).MB, "0.0") & " MB may be too big for one email"
        End If
    Next m
    SetStages
End Sub

' Where each packet stands, and the words CHECK shows for it.
Private Sub SetStages()
    Dim p As Long, m As Long
    For p = 1 To mNPkt
        m = mPkt(p).Mail
        With mPkt(p)
            If Len(.Problem) > 0 Then
                .Stage = ST_HELD: .Waiting = "HELD: " & .Problem
            ElseIf .Notarized > 0 Then
                .Stage = ST_BACK: .Waiting = "Back notarized " & Format$(.Notarized, "m/d") & ": key it from ENTRY, then click Mark filed"
            ElseIf mMail(m).IsSent Then
                .Stage = ST_OUT: .Waiting = "Sent" & IIf(mMail(m).SentOn > 0, " " & Format$(mMail(m).SentOn, "m/d"), "") & _
                    ", waiting on the notarized MF-12"
            ElseIf Len(.VIMissing) > 0 Then
                .Stage = ST_FILES: .Waiting = "Waiting on the VI for BOL " & .VIMissing
            ElseIf .FormState <> "MADE" And .FormState <> "KEPT" Then
                .Stage = ST_FILES: .Waiting = "MF-12 " & LCase$(.FormState)
            ElseIf Len(mMail(m).Held) > 0 Then
                .Stage = ST_HELD: .Waiting = "HELD: " & mMail(m).Held
            ElseIf mMail(m).InDrafts Then
                .Stage = ST_DRAFT: .Waiting = "In Outlook Drafts, not sent yet"
            Else
                .Stage = ST_SEND: .Waiting = "READY: click Make drafts"
            End If
        End With
    Next p
End Sub

' Everything in the email's packet folders except the CI copies, notarized copies and the UPLOAD folder.
' MF-12s first.
Private Sub ListAttachments(ByVal m As Long)
    Dim p As Long, files As Variant, i As Long, pass As Long, f As String, stem As String, ext As String, isForm As Boolean
    mMail(m).Files = "": mMail(m).FileCount = 0: mMail(m).MB = 0
    For pass = 1 To 2
        For p = 1 To mNPkt
            If mPkt(p).Mail = m Then
                files = ListFiles(mPkt(p).Folder)
                If Not IsEmpty(files) Then
                    SortStrings files
                    For i = 1 To UBound(files)
                        f = files(i)
                        SplitExt f, stem, ext
                        isForm = IsExpectedForm(p, f)
                        If Not IsJunkFile(f) And InStr(UCase$(f), "NOTARIZED") = 0 And FileTag(stem) <> "CI" And _
                           Right$(LCase$(f), 8) <> ".partial" Then
                            If (pass = 1 And isForm) Or (pass = 2 And Not isForm) Then
                                mMail(m).Files = mMail(m).Files & mPkt(p).Folder & "\" & f & vbLf
                                mMail(m).FileCount = mMail(m).FileCount + 1
                                mMail(m).MB = mMail(m).MB + FileLen(mPkt(p).Folder & "\" & f) / 1048576#
                            End If
                        End If
                    Next i
                End If
            End If
        Next p
    Next pass
End Sub

Private Sub MakeDrafts()
    Dim m As Long, ol As Object, subj As String, body As String

    If Len(mTestDraftDir) = 0 Then
        Set ol = OutlookApp()
        If ol Is Nothing Then Err.Raise vbObjectError + 14, , "Couldn't open Outlook."
    End If
    For m = 1 To mNMail
        If Len(mMail(m).Held) = 0 And Not mMail(m).IsSent Then
            mStage = "making the draft for " & mSup(mMail(m).Sup).ShortName & " " & PeriodLabel(mMail(m).PerKey)
            If mMail(m).InDrafts Then
                mMail(m).Status = "ALREADY IN DRAFTS"
            Else
                subj = FillWords(mSubject, m)
                body = FillWords(mBody, m)
                If Len(mTestDraftDir) > 0 Then TestDraft m, subj, body Else OutlookDraft ol, m, subj, body
                mMail(m).Status = "DRAFTED " & Format$(Now, "m/d h:nn AM/PM")
                mMail(m).InDrafts = True
            End If
        End If
    Next m
End Sub

Private Sub OutlookDraft(ol As Object, ByVal m As Long, ByVal subj As String, ByVal body As String)
    Dim d As Object, insp As Object, f As Variant, html As String, pos As Long, toAddr As String, ccAddr As String
    Addresses mMail(m).Sup, toAddr, ccAddr
    Set d = ol.CreateItem(0)
    d.To = toAddr
    If Len(ccAddr) > 0 Then d.CC = ccAddr
    d.Subject = subj
    For Each f In Split(mMail(m).Files, vbLf)
        If Len(f) > 0 Then d.Attachments.Add CStr(f)
    Next f
    ' opening the item's window (without showing it) puts your default signature in the body
    On Error Resume Next
    Set insp = d.GetInspector
    html = d.HTMLBody
    On Error GoTo 0
    pos = InStr(1, LCase$(html), "<body")
    If pos > 0 Then pos = InStr(pos, html, ">")
    If pos > 0 Then
        d.HTMLBody = Left$(html, pos) & BodyHtml(body) & Mid$(html, pos + 1)
    Else
        d.Body = body & vbCrLf & vbCrLf & d.Body
    End If
    d.Save
    d.Close 0
End Sub

' Who a supplier's email goes to: Email address, else Group Email Address, else Additional
' Email. The others go on CC.
Private Sub Addresses(ByVal s As Long, ByRef toAddr As String, ByRef ccAddr As String)
    Dim a As Variant, i As Long
    a = Array(mSup(s).Email, mSup(s).GroupMail, mSup(s).Extra)
    toAddr = "": ccAddr = ""
    For i = 0 To 2
        If Len(a(i)) > 0 Then
            If Len(toAddr) = 0 Then
                toAddr = a(i)
            Else
                ccAddr = ccAddr & IIf(Len(ccAddr) > 0, "; ", "") & a(i)
            End If
        End If
    Next i
End Sub

Private Function TestDraftPath(ByVal m As Long) As String
    TestDraftPath = mTestDraftDir & "\" & mSup(mMail(m).Sup).ShortName & " " & (mMail(m).PerKey \ 100) & "-" & _
        Format$(mMail(m).PerKey Mod 100, "00") & " draft.txt"
End Function

Private Sub TestDraft(ByVal m As Long, ByVal subj As String, ByVal body As String)
    Dim f As Integer, toAddr As String, ccAddr As String
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

Private Function FillWords(ByVal t As String, ByVal m As Long) As String
    t = Replace(t, ">MONTH<", MonthName(mMail(m).PerKey Mod 100))
    t = Replace(t, ">YEAR<", CStr(mMail(m).PerKey \ 100))
    t = Replace(t, ">SUPPLIER<", mSup(mMail(m).Sup).ShortName)
    t = Replace(t, ">US<", mUs)
    t = Replace(t, ">CONTACT<", mSup(mMail(m).Sup).Contact)
    t = Replace(t, ">TYPES<", TypeWords(m))
    FillWords = t
End Function

' "licensed customer sales and government customer sales", in the SETUP list's order.
Private Function TypeWords(ByVal m As Long) As String
    Dim t As Long, p As Long, words() As String, n As Long, i As Long, has As Boolean
    For t = 1 To mNType
        has = False
        For p = 1 To mNPkt
            If mPkt(p).Mail = m And mPkt(p).Prefix = mTypePrefix(t) Then has = True
        Next p
        If has Then
            For i = 1 To n
                If words(i) = mTypePhrase(t) Then has = False
            Next i
        End If
        If has Then n = n + 1: ReDim Preserve words(1 To n): words(n) = mTypePhrase(t)
    Next t
    For i = 1 To n
        If i = 1 Then
            TypeWords = words(i)
        ElseIf i = n Then
            TypeWords = TypeWords & " and " & words(i)
        Else
            TypeWords = TypeWords & ", " & words(i)
        End If
    Next i
End Function

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
'  The CHECK sheet
' ======================================================================================

Private Sub WriteCheck(ByVal stepNo As Long)
    Dim ws As Worksheet, r As Long, i As Long, m As Long, p As Long, t As Long, first As Long, ix As Variant, s As Long
    Dim n As Long, g As Double, rf As Double, al As Double, tt As Double, k As Long, keys As Collection
    Dim toAddr As String, ccAddr As String

    Set ws = SheetNamed(CHECK_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = CHECK_SHEET
    End If
    ws.Hyperlinks.Delete
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "Every refund still open on Refunds Due"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("E1").Value = Choose(stepNo, "Packets built ", "Drafts made ", "Marked filed ") & Format$(Now, "m/d/yyyy h:nn AM/PM")
    With ws.Range("A2:L2")
        .Merge
        .Font.Bold = True
        .Font.Size = 12
        If mNLook = 0 Then
            .Value = "All clear. " & mNPkt & " packet(s) open."
            .Interior.Color = GREEN
        Else
            .Value = "LOOK: " & mNLook & " thing(s) below need a look. " & mNPkt & " packet(s) open."
            .Interior.Color = RED
        End If
    End With
    ws.Range("A3").Value = "Packets folder"
    ws.Hyperlinks.Add Anchor:=ws.Range("B3"), Address:=mBase, TextToDisplay:=mBase
    r = 4
    If mLeftAlone > 0 Then
        ws.Cells(r, 1).Value = "Left alone: " & mLeftAlone & " open row(s) from before " & PeriodLabel(mOldest) & _
            ", the Oldest period to work on SETUP."
        r = r + 1
    End If
    If mMoved > 0 Then
        ws.Cells(r, 1).Value = "Moved " & mMoved & " paid row(s) to Refunds Received."
        r = r + 1
    End If
    r = r + 1

    If mNLook > 0 Then
        r = TableHead(ws, r, Array("Needs a look", "Details", "How to fix"))
        For i = 1 To mNLook
            ws.Cells(r, 1).Value = mLook(i).What
            ws.Cells(r, 2).Value = mLook(i).Detail
            ws.Cells(r, 3).Value = mLook(i).HowTo
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = RED
            r = r + 1
        Next i
        r = r + 1
    End If

    ' every open packet, oldest period first, with what it's waiting on
    first = r
    r = TableHead(ws, r, Array("Period", "Packet", "Rows", "Gallons", "Refund", "Allowance", "TTL Refund", "On the MF-12", _
                               "MF-12 forms", "VIs", "Waiting on", "Folder"))
    ix = PacketOrder()
    For i = 1 To mNPkt
        p = ix(i)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = PeriodLabel(mPkt(p).PerKey)
        ws.Cells(r, 2).Value = mPkt(p).PName
        ws.Cells(r, 3).Value = mPkt(p).NRows
        ws.Cells(r, 4).Value = Round(mPkt(p).Gallons, 3)
        ws.Cells(r, 5).Value = Round(mPkt(p).Refund, 2)
        ws.Cells(r, 6).Value = Round(mPkt(p).Allow, 2)
        ws.Cells(r, 7).Value = Round(mPkt(p).Ttl, 2)
        ws.Cells(r, 8).Value = Round(mPkt(p).Amount, 2)
        ws.Cells(r, 9).Value = mPkt(p).NForms & " " & IIf(Len(mPkt(p).Problem) > 0, "HELD", mPkt(p).FormState)
        If Len(mPkt(p).VIMissing) = 0 Then ws.Cells(r, 10).Value = "all there" Else ws.Cells(r, 10).Value = "missing " & mPkt(p).VIMissing
        ws.Cells(r, 11).Value = mPkt(p).Waiting
        Select Case mPkt(p).Stage
            Case ST_HELD, ST_FILES: ws.Cells(r, 11).Interior.Color = RED
            Case ST_SEND, ST_BACK: ws.Cells(r, 11).Interior.Color = GREEN
            Case Else: ws.Cells(r, 11).Interior.Color = YELLOW
        End Select
        If FolderExists(mPkt(p).Folder) Then ws.Hyperlinks.Add Anchor:=ws.Cells(r, 12), Address:=mPkt(p).Folder, TextToDisplay:=mPkt(p).PName
        r = r + 1
    Next i
    If mNPkt = 0 Then ws.Cells(r, 1).Value = "Nothing open.": r = r + 1
    ws.Range(ws.Cells(first, 4), ws.Cells(r, 4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 5), ws.Cells(r, 8)).NumberFormat = "#,##0.00"
    r = r + 1

    ' one email per supplier per period
    r = TableHead(ws, r, Array("Period", "Supplier", "Seller on the MF-12", "License", "Types", "Rows", "Amount on MF-12s", _
                               "NET", "To", "CC", "Attachments", "Email"))
    ix = MailOrder()
    For i = 1 To mNMail
        m = ix(i)
        s = mMail(m).Sup
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = PeriodLabel(mMail(m).PerKey)
        ws.Cells(r, 2).Value = mSup(s).ShortName
        ws.Cells(r, 3).Value = mSup(s).Corp
        ws.Cells(r, 4).NumberFormat = "@"
        ws.Cells(r, 4).Value = mSup(s).License
        ws.Cells(r, 5).Value = TypeWords(m)
        ws.Cells(r, 6).Value = mMail(m).NRows
        g = 0
        For p = 1 To mNPkt
            If mPkt(p).Mail = m Then g = g + mPkt(p).Amount
        Next p
        ws.Cells(r, 7).Value = Round(g, 2)
        ws.Cells(r, 7).NumberFormat = "#,##0.00"
        ws.Cells(r, 8).Value = IIf(mSup(s).IsNet, "NET", "gross")
        Addresses s, toAddr, ccAddr
        ws.Cells(r, 9).Value = toAddr
        ws.Cells(r, 10).Value = ccAddr
        If mMail(m).FileCount > 0 Then ws.Cells(r, 11).Value = AttachSummary(m)
        ws.Cells(r, 12).Value = mMail(m).Status
        If Len(mMail(m).Held) = 0 Or mMail(m).IsSent Then ws.Cells(r, 12).Interior.Color = GREEN Else ws.Cells(r, 12).Interior.Color = RED
        r = r + 1
    Next i
    ' rows whose supplier isn't on Corp License Number
    For i = 1 To mNRow
        If mRow(i).Sup = 0 Then
            ws.Cells(r, 1).NumberFormat = "@"
            ws.Cells(r, 1).Value = PeriodLabel(mRow(i).PerKey)
            ws.Cells(r, 2).Value = "(no match)"
            ws.Cells(r, 3).Value = mRow(i).Corp
            ws.Cells(r, 4).NumberFormat = "@"
            ws.Cells(r, 4).Value = mRow(i).License
            ws.Cells(r, 12).Value = "HELD: row " & mRow(i).SheetRow & ", " & mRow(i).Problem
            ws.Cells(r, 12).Interior.Color = RED
            r = r + 1
        End If
    Next i
    r = r + 1

    ' totals by period and refund type, gross, the way each MFD-33 wants them
    first = r
    r = TableHead(ws, r, Array("Totals by period and refund type (gross, for the MFD-33)", "Rows", "Gallons", "Refund", "Allowance", "TTL Refund"))
    Set keys = New Collection
    For p = 1 To mNPkt
        ix = Empty
        For k = 1 To keys.Count
            If keys(k) = mPkt(p).PerKey Then ix = k
        Next k
        If IsEmpty(ix) Then
            For k = 1 To keys.Count
                If keys(k) > mPkt(p).PerKey Then keys.Add mPkt(p).PerKey, Before:=k: ix = k: Exit For
            Next k
            If IsEmpty(ix) Then keys.Add mPkt(p).PerKey
        End If
    Next p
    For k = 1 To keys.Count
        For t = 1 To mNType
            n = 0: g = 0: rf = 0: al = 0: tt = 0
            For i = 1 To mNRow
                If mRow(i).PerKey = keys(k) And Squash(mRow(i).TypeText) = mTypeText(t) Then
                    n = n + 1: g = g + mRow(i).Gallons: rf = rf + mRow(i).Refund: al = al + mRow(i).Allow: tt = tt + mRow(i).Ttl
                End If
            Next i
            If n > 0 Then
                ws.Cells(r, 1).Value = PeriodLabel(keys(k)) & " " & mTypePrefix(t) & ": " & mTypePhrase(t)
                ws.Cells(r, 2).Value = n: ws.Cells(r, 3).Value = Round(g, 3): ws.Cells(r, 4).Value = Round(rf, 2)
                ws.Cells(r, 5).Value = Round(al, 2): ws.Cells(r, 6).Value = Round(tt, 2)
                r = r + 1
            End If
        Next t
    Next k
    ws.Range(ws.Cells(first, 3), ws.Cells(r, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 4), ws.Cells(r, 6)).NumberFormat = "#,##0.00"

    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 30
    ws.Columns("F:H").ColumnWidth = 14
    ws.Columns("I:J").ColumnWidth = 24
    ws.Columns(11).ColumnWidth = 44
    ws.Columns(12).ColumnWidth = 44
    ws.Range("A4:L" & r).VerticalAlignment = xlTop
    ws.Range("A4:L" & r).WrapText = True
    On Error Resume Next
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

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

' "31 files, 1.3 MB: 3 MF-12, 28 VI"
Private Function AttachSummary(ByVal m As Long) As String
    Dim f As Variant, nm As String, stem As String, ext As String, forms As Long, vis As Long, other As Long
    For Each f In Split(mMail(m).Files, vbLf)
        If Len(f) > 0 Then
            nm = Mid$(f, InStrRev(f, "\") + 1)
            SplitExt nm, stem, ext
            If InStr(UCase$(nm), " MF-12") > 0 Then
                forms = forms + 1
            ElseIf FileTag(stem) = "VI" Then
                vis = vis + 1
            Else
                other = other + 1
            End If
        End If
    Next f
    AttachSummary = mMail(m).FileCount & " files, " & Format$(mMail(m).MB, "0.0") & " MB: " & forms & " MF-12, " & vis & " VI" & _
        IIf(other > 0, ", " & other & " other", "")
End Function

' ======================================================================================
'  The HOW TO tab: three buttons and the steps, at the front of the workbook
' ======================================================================================

' The steps live inside the workbook, with
' buttons so nothing needs Alt F8. Redrawn every run; the buttons are made once and only
' moved back into place, so a button you dragged still works.
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
    ws.Cells.Font.Size = 12
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 100
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 48

    ws.Range("B1").Value = "GA refunds: the running list, MF-12 packets, Outlook drafts, portal entry"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 18
    HowToButton ws, "btnBuild", "1.  Build packets", "Refunds1_BuildPackets", 0
    HowToButton ws, "btnDrafts", "2.  Make drafts", "Refunds2_MakeDrafts", 1
    HowToButton ws, "btnFiled", "3.  Mark filed", "Refunds3_MarkFiled", 2

    r = 4
    r = HowToLine(ws, r, "Before you start (the first time only)", True)
    r = HowToLine(ws, r, "1. This macro goes in the GA tracking workbook, the one with Refunds Due and Corp License Number. " & _
        "Save the workbook as .xlsm.", False)
    r = HowToLine(ws, r, "2. Save " & TEMPLATE_NAME & " (it comes with this macro) in the same folder as the workbook.", False)
    r = HowToLine(ws, r, "That's all. You don't make any folders. The first click on Build packets asks for the company's " & _
        "name and address for the MF-12, then for the folder the packets go in, then which period to start from. " & _
        "It asks each one once and makes every folder itself.", False)
    r = r + 1
    r = HowToLine(ws, r, "Each time", True)
    r = HowToLine(ws, r, "1. New refunds: paste the refund report's rows (the Excel listing the refund loads) at the bottom of " & _
        "Refunds Due, under the rows already there. Leave the old rows. Refunds Due is the running list of every refund still out.", False)
    r = HowToLine(ws, r, "2. Click Build packets. CHECK lists every open packet, any period, and what each one is waiting on.", False)
    r = HowToLine(ws, r, "3. Save each load's invoices into the INBOX folder as ""12345 VI"" and ""12345 CI"" (12345 is the BOL). " & _
        "Then click Build packets again. It files them into their packets and fills the MF-12s.", False)
    r = HowToLine(ws, r, "4. When CHECK says READY, click Make drafts: one Outlook draft per supplier per period. " & _
        "Nothing is sent; look each one over and send it yourself. The next click stamps MF-12 Sent.", False)
    r = HowToLine(ws, r, "5. When a notarized MF-12 comes back, save it in its packet folder with NOTARIZED at the end of the " & _
        "name. The next click stamps MF-12 Signed and puts the packet on ENTRY.", False)
    r = HowToLine(ws, r, "6. ENTRY is the typing sheet for the Georgia portal: one group per period and refund type, which " & _
        "is one MFD-33 filing. When it's keyed, click Mark filed. It stamps today's date as Filed Date and those packets drop off.", False)
    r = HowToLine(ws, r, "7. When the state pays, type the Refund Receive Date on those rows. The next click moves them to " & _
        "the bottom of Refunds Received.", False)
    r = r + 1
    r = HowToLine(ws, r, "The tabs", True)
    r = HowToLine(ws, r, "SETUP holds the company's details, the packets folder, the template, the oldest period to work, " & _
        "the refund types and the email words. The first click fills it; change anything there.", False)
    r = HowToLine(ws, r, "CHECK is where everything stands: each open packet and what it's waiting on, each email, and " & _
        "anything the macro could not place.", False)
    r = HowToLine(ws, r, "ENTRY is the typing sheet for the Georgia portal.", False)
    r = HowToLine(ws, r, "Corp License Number is the contacts list. The Corp name there prints on every MF-12 exactly " & _
        "as typed, and every click copies it onto the matching rows of Refunds Due, so trim any extra text " & _
        "off it once, like "" - ST (COMPANY)"". Short Name and NET are the two columns the macro added; a blank Short " & _
        "Name gets the Corp name's first word.", False)
    r = r + 1
    r = HowToLine(ws, r, "Rules it keeps", True)
    r = HowToLine(ws, r, "It never sends an email, never deletes a file, and never writes over one. A packet with " & _
        "anything on CHECK gets no draft until it is fixed. A date you typed in a status column is never changed.", False)
    r = HowToLine(ws, r, "If the buttons ever go missing: Alt F8, run Refunds1_BuildPackets, and this tab redraws " & _
        "with them.", False)
    ws.Range("B4:B" & r).WrapText = True
    ws.Range("B4:B" & r).VerticalAlignment = xlTop
End Sub

' One line of the HOW TO tab; hands back the next row.
Private Function HowToLine(ws As Worksheet, ByVal r As Long, ByVal text As String, ByVal isHead As Boolean) As Long
    ws.Cells(r, 2).Value = text
    If isHead Then
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, 2).Font.Size = 14
        ws.Cells(r, 2).Interior.Color = YELLOW
    End If
    HowToLine = r + 1
End Function

' A big button on row 2. Made once, then only moved back into place.
Private Sub HowToButton(ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal macroName As String, ByVal slot As Long)
    Dim b As Object
    On Error Resume Next
    Set b = ws.Buttons(nm)
    On Error GoTo 0
    If b Is Nothing Then
        Set b = ws.Buttons.Add(0, 0, 100, 30)
        b.Name = nm
    End If
    b.Placement = xlFreeFloating
    b.Left = ws.Range("B2").Left + 4 + slot * 236
    b.Top = ws.Range("B2").Top + 4
    b.Width = 224
    b.Height = ws.Range("B2").Height - 8
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
' back notarized and not filed yet, one group per period and refund type (one MFD-33 filing
' each), inside it one block per MF-12 packet, the columns in the portal's own order, the rows
' in the order they sit on the form, and a total under each block to check the portal's running
' total against. Values only. Redrawn every run.
Private Sub WriteEntry()
    Dim ws As Worksheet, r As Long, p As Long, i As Long, k As Long, ix As Variant, order As Variant, s As Long, t As Long
    Dim g As Double, a As Double, fg As Double, fa As Double, formNo As Long, slot As Long, head As String
    Dim curT As Long, curPer As Long, bannerRow As Long, tg As Double, ta As Double, tRows As Long, tForms As Long, tHeld As Long
    Dim shown As Long

    Set ws = SheetNamed(ENTRY_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = ENTRY_SHEET
    End If
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "Portal entry: every packet back notarized and not filed yet"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Value = "One group per period and refund type, which is one MFD-33 filing; inside it, one block per MF-12. " & _
        "Type each row into the Georgia portal left to right, top to bottom. " & _
        "Exempt Portion in the portal stays blank, so it is not on this sheet. " & _
        "Tax Paid to Seller is the gross refund on every row, even for a NET supplier; NET is only on its MF-12. " & _
        "When a block is in, the portal's running total should match the total under it; the MFD-33 total closes each group. " & _
        "When it's all keyed, click Mark filed on the HOW TO tab. Drawn " & Format$(Now, "m/d/yyyy h:nn AM/PM") & "."
    With ws.Range("A2:G2")
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
        If mPkt(p).Notarized > 0 Then
            shown = shown + 1
            s = mPkt(p).Sup
            t = TypeIndex(mPkt(p).Prefix)
            ' A new period or refund type: close the last group with its MFD-33 total, open this one
            ' with a banner. The banner's counts are filled in once the group is done (see EntryGroupClose).
            If t <> curT Or mPkt(p).PerKey <> curPer Then
                If curT > 0 Then r = EntryGroupClose(ws, r, curT, curPer, bannerRow, tg, ta, tRows, tForms, tHeld)
                curT = t
                curPer = mPkt(p).PerKey
                bannerRow = r
                ws.Rows(r).RowHeight = 28
                r = r + 1
                tg = 0: ta = 0: tRows = 0: tForms = 0: tHeld = 0
            End If
            head = mSup(s).ShortName & " (" & mPkt(p).PName & "), " & mPkt(p).NRows & " row" & IIf(mPkt(p).NRows = 1, "", "s")
            If mPkt(p).NForms > 1 Then head = head & " on " & mPkt(p).NForms & " MF-12 forms"
            If mSup(s).IsNet Then head = head & " (its MF-12 shows NET; the figures here are gross, as the state wants)"
            If Len(mPkt(p).Problem) > 0 Then head = head & ". HELD, see CHECK: nothing to key yet."
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
                .Merge
                .Value = head
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = IIf(Len(mPkt(p).Problem) > 0, RED, YELLOW)
            End With
            ws.Rows(r).RowHeight = 18
            r = r + 1
            If Len(mPkt(p).Problem) > 0 Then
                tHeld = tHeld + 1
            Else
                tRows = tRows + mPkt(p).NRows: tForms = tForms + mPkt(p).NForms
                r = TableHead(ws, r, Array("Seller Name", "Seller License Number", "Date", "Invoice Number", "Gallons", _
                    "Tax Paid to Seller", "Type of Fuel"))
                ix = PacketRows(p)
                g = 0: a = 0: fg = 0: fa = 0: formNo = 1: slot = 0
                For k = 1 To UBound(ix)
                    With mRow(ix(k))
                        ws.Cells(r, 1).Value = mSup(s).Corp
                        ws.Cells(r, 2).NumberFormat = "@"
                        ws.Cells(r, 2).Value = .License
                        ws.Cells(r, 3).NumberFormat = "mm/dd/yyyy"
                        ws.Cells(r, 3).Value = .InvDate
                        ws.Cells(r, 4).NumberFormat = "@"
                        ws.Cells(r, 4).Value = .Inv
                        ws.Cells(r, 5).NumberFormat = GalFormat(.Gallons)
                        ws.Cells(r, 5).Value = .Gallons
                        ' Always the gross Refund, never TTL: the state takes gross on the portal and
                        ' the MFD-33; the NET figure is only for the supplier's own MF-12.
                        ws.Cells(r, 6).NumberFormat = "#,##0.00"
                        ws.Cells(r, 6).Value = Round(.Refund, 2)
                        ws.Cells(r, 7).Value = Trim$(.FuelOpt)
                        g = g + .Gallons: a = a + .Refund
                        fg = fg + .Gallons: fa = fa + .Refund
                    End With
                    r = r + 1
                    slot = slot + 1
                    If mPkt(p).NForms > 1 And (slot = ROWS_PER_FORM Or k = UBound(ix)) Then
                        ws.Cells(r, 1).Value = "end of MF-12 (" & formNo & " of " & mPkt(p).NForms & ")"
                        ws.Cells(r, 5).NumberFormat = GalFormat(fg)
                        ws.Cells(r, 5).Value = Round(fg, 3)
                        ws.Cells(r, 6).NumberFormat = "#,##0.00"
                        ws.Cells(r, 6).Value = Round(fa, 2)
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Italic = True
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = GRAY
                        r = r + 1
                        formNo = formNo + 1: slot = 0: fg = 0: fa = 0
                    End If
                Next k
                ws.Cells(r, 1).Value = "Total for " & mPkt(p).PName & ", " & mPkt(p).NRows & " row" & IIf(mPkt(p).NRows = 1, "", "s")
                ws.Cells(r, 5).NumberFormat = GalFormat(g)
                ws.Cells(r, 5).Value = Round(g, 3)
                ws.Cells(r, 6).NumberFormat = "#,##0.00"
                ws.Cells(r, 6).Value = Round(a, 2)
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
                    .Font.Bold = True
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlDouble
                End With
                tg = tg + g: ta = ta + a
                r = r + 1
            End If
            r = r + 1
        End If
    Next i
    If curT > 0 Then r = EntryGroupClose(ws, r, curT, curPer, bannerRow, tg, ta, tRows, tForms, tHeld)
    If shown = 0 Then
        ws.Cells(r, 1).Value = "Nothing to key yet. A packet shows here once its notarized MF-12 is saved in its folder " & _
            "(NOTARIZED at the end of the name), and drops off when you click Mark filed."
        r = r + 1
    End If

    ' Tight for reading at a desk: the invoice rows close together and the
    ' numbers close to each other, with a clear gap between Date and Invoice Number. The banners,
    ' block heads and MFD-33 totals keep their own heights.
    ws.Range("A4:G" & r).VerticalAlignment = xlCenter
    For k = 4 To r
        If ws.Cells(k, 1).Value = "Seller Name" Then
            ws.Range(ws.Cells(k, 1), ws.Cells(k, 7)).WrapText = True
            ws.Rows(k).RowHeight = 26
        ElseIf ws.Cells(k, 1).Font.Size < 12 Then
            ws.Rows(k).RowHeight = 12.75
        End If
    Next k
    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 11
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 9
    ws.Columns(6).ColumnWidth = 11
    ws.Columns(7).ColumnWidth = 14
    ws.Range("A4:A" & r).ShrinkToFit = True
    ws.Range("C4:C" & r).HorizontalAlignment = xlLeft
    ws.Range("D4:D" & r).HorizontalAlignment = xlLeft
    ws.Range("D5:D" & r).IndentLevel = 2
    ws.Range("E4:F" & r).HorizontalAlignment = xlRight
    ws.Range("G5:G" & r).IndentLevel = 1
    ' the title stays put while the blocks scroll
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Closes a period and refund-type group on ENTRY: the MFD-33 total line, two blank rows, and the
' banner above the group filled in with what it holds. Hands back the next row.
Private Function EntryGroupClose(ws As Worksheet, ByVal r As Long, ByVal t As Long, ByVal perKey As Long, ByVal bannerRow As Long, _
    ByVal tg As Double, ByVal ta As Double, ByVal tRows As Long, ByVal tForms As Long, ByVal tHeld As Long) As Long
    Dim txt As String, when As String
    when = MonthName(perKey Mod 100) & " " & (perKey \ 100)
    ws.Cells(r, 1).Value = "MFD-33 total, " & when & " " & mTypePhrase(t) & ", " & tRows & " row" & IIf(tRows = 1, "", "s") & _
        " on " & tForms & " MF-12" & IIf(tForms = 1, "", "s")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Merge
    ws.Cells(r, 5).NumberFormat = GalFormat(tg)
    ws.Cells(r, 5).Value = Round(tg, 3)
    ws.Cells(r, 6).NumberFormat = "#,##0.00"
    ws.Cells(r, 6).Value = Round(ta, 2)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = BLUE_PALE
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    ws.Rows(r).RowHeight = 18
    txt = UCase$(when) & ", " & UCase$(mTypePhrase(t)) & ": one MFD-33 filing, " & tForms & " MF-12" & IIf(tForms = 1, "", "s") & _
        " to key, " & tRows & " row" & IIf(tRows = 1, "", "s")
    If tHeld > 0 Then txt = txt & ", plus " & tHeld & " held"
    With ws.Range(ws.Cells(bannerRow, 1), ws.Cells(bannerRow, 7))
        .Merge
        .Value = txt
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = vbWhite
        .Interior.Color = BLUE_DARK
        .VerticalAlignment = xlCenter
    End With
    EntryGroupClose = r + 3
End Function

' Whole gallons show whole; anything else shows its three places.
Private Function GalFormat(ByVal x As Double) As String
    If Abs(x - Round(x, 0)) < 0.0005 Then GalFormat = "#,##0" Else GalFormat = "#,##0.000"
End Function

' Packets in the order to key them: oldest period first, then by refund type in the SETUP list's
' order (one period's type is one MFD-33 filing, so its MF-12s sit together), then by supplier
' short name inside the type.
Private Function PacketOrder() As Variant
    Dim idx() As Long, i As Long, j As Long, t As Long
    If mNPkt = 0 Then PacketOrder = Array(): Exit Function
    ReDim idx(1 To mNPkt)
    For i = 1 To mNPkt
        idx(i) = i
    Next i
    For i = 2 To mNPkt
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Not PktBefore(t, idx(j)) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    PacketOrder = idx
End Function

Private Function PktBefore(ByVal a As Long, ByVal b As Long) As Boolean
    Dim na As String, nb As String, ta As Long, tb As Long
    If mPkt(a).PerKey <> mPkt(b).PerKey Then PktBefore = (mPkt(a).PerKey < mPkt(b).PerKey): Exit Function
    ta = TypeIndex(mPkt(a).Prefix): tb = TypeIndex(mPkt(b).Prefix)
    If ta <> tb Then
        PktBefore = (ta < tb)
    Else
        na = mSup(mPkt(a).Sup).ShortName: nb = mSup(mPkt(b).Sup).ShortName
        PktBefore = (na < nb)
    End If
End Function

' Emails oldest period first, then by supplier short name.
Private Function MailOrder() As Variant
    Dim idx() As Long, i As Long, j As Long, t As Long
    If mNMail = 0 Then MailOrder = Array(): Exit Function
    ReDim idx(1 To mNMail)
    For i = 1 To mNMail
        idx(i) = i
    Next i
    For i = 2 To mNMail
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Not MailBefore(t, idx(j)) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    MailOrder = idx
End Function

Private Function MailBefore(ByVal a As Long, ByVal b As Long) As Boolean
    If mMail(a).PerKey <> mMail(b).PerKey Then
        MailBefore = (mMail(a).PerKey < mMail(b).PerKey)
    Else
        MailBefore = (mSup(mMail(a).Sup).ShortName < mSup(mMail(b).Sup).ShortName)
    End If
End Function

' Where a folder prefix sits on the SETUP list (the first type carrying that prefix).
Private Function TypeIndex(ByVal prefix As String) As Long
    Dim t As Long
    For t = 1 To mNType
        If mTypePrefix(t) = prefix Then TypeIndex = t: Exit Function
    Next t
    TypeIndex = mNType
End Function

' ======================================================================================
'  Small helpers
' ======================================================================================

Private Sub Say(ByVal msg As String, ByVal icon As VbMsgBoxStyle)
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

Private Sub AddLook(ByVal what As String, ByVal detail As String, ByVal howTo As String)
    Dim i As Long
    For i = 1 To mNLook
        If mLook(i).What = what And mLook(i).Detail = detail Then Exit Sub
    Next i
    mNLook = mNLook + 1
    ReDim Preserve mLook(1 To mNLook)
    mLook(mNLook).What = what
    mLook(mNLook).Detail = detail
    mLook(mNLook).HowTo = howTo
End Sub

Private Function SheetNamed(ByVal nm As String) As Worksheet
    Set SheetNamed = SheetNamedIn(mBook, nm)
End Function

Private Function SheetNamedIn(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = LCase$(nm) Then Set SheetNamedIn = ws: Exit Function
    Next ws
End Function

' The first row in the top 20 that has every column asked for. 0 if none.
Private Function HeaderRow(ws As Worksheet, specs As Variant) As Long
    Dim r As Long, i As Long, lastCol As Long, ok As Boolean
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastCol > 60 Then lastCol = 60
    For r = 1 To 20
        ok = True
        For i = 0 To UBound(specs)
            If FindCol(ws, r, CStr(specs(i)), lastCol) = 0 Then ok = False: Exit For
        Next i
        If ok Then HeaderRow = r: Exit Function
    Next r
End Function

Private Function LastHeaderCol(ws As Worksheet, ByVal hr As Long) As Long
    LastHeaderCol = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function FindCol(ws As Worksheet, ByVal r As Long, ByVal spec As String, ByVal lastCol As Long) As Long
    Dim alts As Variant, i As Long, c As Long
    alts = Split(spec, "|")
    For i = 0 To UBound(alts)
        For c = 1 To lastCol
            If Squash(CellStr(ws.Cells(r, c).Value)) = alts(i) Then FindCol = c: Exit Function
        Next c
    Next i
End Function

Private Function Squash(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squash = LCase$(Trim$(s))
End Function

Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger, vbDecimal
            If v = Int(v) And Abs(v) < 1E+15 Then CellStr = Format$(v, "0") Else CellStr = CStr(v)
        Case Else
            CellStr = Trim$(CStr(v))
    End Select
End Function

Private Function PeriodText(ByVal v As Variant) As String
    If VarType(v) = vbDate Then PeriodText = Format$(v, "mmm-yy") Else PeriodText = CellStr(v)
End Function

' 202608 from Aug-26, 08/2026, 2026-08 and the like. 0 if it isn't a month.
Private Function PeriodKey(ByVal t As String) As Long
    Dim m As Long, y As Long
    If ParsePeriod(t, m, y) Then PeriodKey = y * 100 + m
End Function

' Aug-26 from 202608.
Private Function PeriodLabel(ByVal key As Long) As String
    If key = 0 Then PeriodLabel = "(no period)": Exit Function
    PeriodLabel = Left$(MonthName(key Mod 100), 3) & "-" & Format$((key \ 100) Mod 100, "00")
End Function

Private Function ToNum(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim t As String
    ok = False
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        t = Replace(Replace(Replace(Trim$(v), ",", ""), "$", ""), " ", "")
        If Len(t) = 0 Or Not IsNumeric(t) Then Exit Function
        ToNum = CDbl(t)
        ok = True
    ElseIf IsNumeric(v) Then
        ToNum = CDbl(v)
        ok = True
    End If
End Function

Private Function ToDate(ByVal v As Variant, ByRef ok As Boolean) As Date
    ok = False
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToDate = v: ok = True
    ElseIf VarType(v) = vbDouble Then
        If v > 30000 And v < 80000 Then ToDate = CDate(v): ok = True
    ElseIf VarType(v) = vbString Then
        If IsDate(Trim$(v)) Then ToDate = CDate(Trim$(v)): ok = True
    End If
End Function

Private Function ParsePeriod(ByVal t As String, ByRef m As Long, ByRef y As Long) As Boolean
    Dim parts As Variant, a As String, b As String, i As Long
    t = Trim$(Replace(Replace(Replace(t, "-", " "), "/", " "), ".", " "))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    parts = Split(t, " ")
    m = 0: y = 0
    If UBound(parts) <> 1 Then Exit Function
    a = parts(0): b = parts(1)
    If Not IsAllDigits(b) Then Exit Function
    If IsAllDigits(a) Then
        If Len(a) = 4 Then
            y = CLng(a): m = CLng(b)
        Else
            m = CLng(a): y = CLng(b)
        End If
    Else
        For i = 1 To 12
            If Len(a) >= 3 And LCase$(Left$(a, 3)) = LCase$(Left$(MonthName(i), 3)) Then m = i
        Next i
        y = CLng(b)
    End If
    If y >= 0 And y < 100 Then y = y + 2000
    ParsePeriod = (m >= 1 And m <= 12 And y >= 2000 And y <= 2099)
End Function

' The form's own fuel choices, padding and all, so the dropdown shows the right one.
Private Function FuelOption(ByVal t As String) As String
    Dim opts As Variant, i As Long
    opts = Array(Space$(6) & "Aviation", Space$(5) & "Bio Diesel", Space$(5) & "Fuel Oil Clear", Space$(5) & "Fuel Oil Dyed", _
                 Space$(7) & "Gasohol", Space$(6) & "Gasoline", Space$(6) & "Jet Fuel", Space$(9) & "LP", Space$(5) & "Special Fuel")
    t = Replace(Squash(t), " ", "")
    If Len(t) = 0 Then Exit Function
    For i = 0 To UBound(opts)
        If Replace(Squash(CStr(opts(i))), " ", "") = t Then FuelOption = opts(i): Exit Function
    Next i
End Function

Private Function FindSup(ByVal licKey As String) As Long
    Dim i As Long
    For i = 1 To mNSup
        If mSup(i).LicKey = licKey Then FindSup = i: Exit Function
    Next i
End Function

Private Function FindMailing(ByVal s As Long, ByVal perKey As Long) As Long
    Dim i As Long
    For i = 1 To mNMail
        If mMail(i).Sup = s And mMail(i).PerKey = perKey Then FindMailing = i: Exit Function
    Next i
End Function

Private Function FindPkt(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mNPkt
        If mPkt(i).Key = key Then FindPkt = i: Exit Function
    Next i
End Function

Private Function FindType(ByVal t As String) As Long
    Dim i As Long
    t = Squash(t)
    For i = 1 To mNType
        If mTypeText(i) = t Then FindType = i: Exit Function
    Next i
End Function

Private Function LicenseKey(ByVal s As String) As String
    LicenseKey = UCase$(Replace(Replace(Trim$(s), " ", ""), "-", ""))
End Function

' BOLs compared without spaces or case, and without leading zeros when they're all digits.
Private Function BolKey(ByVal s As String) As String
    s = UCase$(Replace(Trim$(s), " ", ""))
    If IsAllDigits(s) Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If
    BolKey = s
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "#" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsAllLetters(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[A-Za-z]" Then Exit Function
    Next i
    IsAllLetters = True
End Function

' Drops the characters Windows won't take in a file name.
Private Function SafeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "")
    Next c
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SafeName = Trim$(s)
End Function

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
Private Function FileTag(ByVal stem As String) As String
    Dim t As Variant, i As Long
    t = Split(Trim$(stem), " ")
    For i = UBound(t) To 0 Step -1
        If Not IsAllDigits(CStr(t(i))) Then FileTag = UCase$(t(i)): Exit Function
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

Private Sub SortStrings(a As Variant)
    Dim i As Long, j As Long, t As String
    For i = LBound(a) + 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If UCase$(a(j)) <= UCase$(t) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

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

Private Function AnsiB(ByVal s As String) As String
    AnsiB = StrConv(s, vbFromUnicode)
End Function
