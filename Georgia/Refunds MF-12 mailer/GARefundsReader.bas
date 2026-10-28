Attribute VB_Name = "GARefundsReader"
' GA refunds, the reader: reads SETUP, Refunds Due, Corp License Number and Refunds Received
' into rows, suppliers, packets and emails, and writes back only to those tabs. It never
' touches a folder, a file or Outlook; GARefundsWorker does that. A packet is a license, a
' refund type and a period; the Short Name only labels it.

Option Explicit
Option Private Module

Public Const ROWS_PER_FORM As Long = 21       ' invoice lines on one MF-12; more spill onto a second form
Public Const SETUP_SHEET As String = "SETUP"
Public Const CHECK_SHEET As String = "CHECK"
Public Const ENTRY_SHEET As String = "ENTRY"
Public Const HOWTO_SHEET As String = "HOW TO"
Public Const CONTACTS_SHEET As String = "Corp License Number"
Private Const RECEIVED_SHEET As String = "Refunds Received"
Public Const TEMPLATE_NAME As String = "mf12_template.pdf"

' Words on SETUP
Private Const PACKETS_HINT As String = "Each period gets its own folder inside it, like 2026-08 August, and INBOX sits here too."
Private Const MONTH_NOTE As String = "Only this month's rows are worked. Every button asks for it, with this filled in."
Private Const TYPE_WORD_HEAD As String = "Name in folders and files"
Private Const OLD_PREFIX_HEAD As String = "Old folder prefix"
Private Const TYPE_NOTE As String = "A refund type the list doesn't have lands on CHECK. Add it on a yellow line: the words " & _
    "exactly as the report has them, its name for the folders and files (like Diversion), and the words for the email (like diversions)."
Private Const FILL_NOTE As String = "These get filled in per supplier: >MONTH< >YEAR< >SUPPLIER< >US< >CONTACT< >GREETING< >COUNTS< >TYPES<"
Private Const OLD_FILL_NOTE As String = "These get filled in per supplier: >MONTH< >YEAR< >SUPPLIER< >US< >TYPES< >CONTACT<"
Private Const LETTER_NOTE As String = "Alt+Enter starts a new line inside the cell. >GREETING< is Hi and the contact's first " & _
    "name, or Hello with no contact name. >COUNTS< is a line per refund type with its MF-12s and invoices; with one type " & _
    "it ends the sentence instead."
Private Const OLD_LETTER_NOTE As String = "Alt+Enter starts a new line inside the cell."

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

' Colors, written back to front the way Excel wants them (blue, green, red)
Public Const GREEN As Long = &HCEEFC6
Public Const RED As Long = &HCEC7FF
Public Const YELLOW As Long = &HCCF2FF
Public Const GRAY As Long = &HD9D9D9
Public Const BLUE_DARK As Long = &H794E1F
Public Const BLUE_PALE As Long = &HF2E6DD
Public Const SKY As Long = &HF7EBDD          ' on CHECK: sent, waiting on it back
Public Const PEACH As Long = &HD6E4FC        ' on CHECK: waiting on an invoice

Public Type RRow
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
    TypeNo As Long           ' which line of SETUP's refund type list, 0 if none
    Sup As Long              ' which supplier, 0 if none
    Pkt As Long              ' which packet, 0 if none
    Amount As Double         ' what goes on the MF-12: TTL Refund for NET suppliers, else Refund
    Problem As String
    HasVI As Boolean
    HasSent As Boolean       ' MF-12 Sent already has something in it
    SentOn As Date
    HasSigned As Boolean     ' MF-12 Signed already has something in it
End Type

Public Type RSup
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
Public Type RMail
    Sup As Long
    PerKey As Long
    Held As String           ' why no draft, blank when ready
    IsSent As Boolean
    SentOn As Date
    InDrafts As Boolean
    Drafted As Date          ' when this click made its draft, 0 if it didn't
    Files As String
    MB As Double
End Type

Public Type RPkt
    Key As String            ' 202608|12345GR|LICENSED: the period, the license and the refund type
    PName As String          ' Acme Licensed
    Word As String           ' Licensed, the refund type's name in folders and files
    TypeNo As Long           ' the first line of SETUP's refund type list with that name
    Prefix As String         ' LIC, what the last version put in front of its folder names
    PerKey As Long
    Sup As Long
    Mail As Long
    Folder As String         ' filled in by the worker, which finds the folders
    Subject As String        ' the subject its packet.txt says its draft went under, blank before one
    TagShort As String       ' the Short Name its packet.txt had when the worker found it
    NRows As Long
    NForms As Long
    Problem As String        ' blank when its rows are all good and nothing went wrong with it
    VIMissing As String
    FormState As String      ' what's on disk: MADE, KEPT, MISSING, or what's wrong
    Notarized As Date        ' the newest NOTARIZED copy's date, 0 if none
    Stage As Long
    MailForms As Long        ' filled in by the worker: the MF-12s and invoices its email attaches
    MailInvoices As Long
End Type

Public Type RLook
    What As String
    Detail As String
    HowTo As String
    Starts As String         ' lines that start alike are one line: how it starts for one thing
    Many As String           ' and for more than one
    Items As String          ' each thing it names, one to a line
End Type

Public mRow() As RRow, mNRow As Long
Public mSup() As RSup, mNSup As Long
Public mMail() As RMail, mNMail As Long
Public mPkt() As RPkt, mNPkt As Long
Public mLook() As RLook, mNLook As Long
Public mTypeText() As String, mTypeWord() As String, mTypePrefix() As String, mTypePhrase() As String, mNType As Long

Public mBook As Workbook
Public mByHand As Long          ' this month's open rows sent by hand, with no packet from this macro, left alone
Public mHandRows As Object      ' their sheet rows, found once the click has looked in the folders
Public mOtherBols As Object     ' the BOLs of other months' open rows, so their INBOX files wait for that month
Public mShortsFilled As Long    ' Short Names filled in from the Corp name this click
Public mPeriodNote As String    ' periods it couldn't read and went by the invoice date for, for under CHECK's bar
Private mPaidStuck As String    ' paid rows Refunds Received couldn't take this click, one to a line
Public mStage As String         ' what it's doing, for the line on CHECK if something goes wrong

' The kinds of list SortList sorts
Public Const SORT_ROWS As Long = 1       ' a packet's rows: by invoice date, then BOL
Public Const SORT_PACKETS As Long = 2    ' period, then refund type, then supplier
Public Const SORT_MAILS As Long = 3      ' period, then supplier
Public Const SORT_NAMES As Long = 4      ' plain names, any capitals

' from SETUP
Private mBuyer As String, mAddress As String, mCity As String, mState As String, mZip As String
Public mUs As String, mBase As String, mTemplatePath As String, mMonth As Long
Public mSubject As String, mBody As String

' ======================================================================================
'  The workbook and the SETUP tab
' ======================================================================================

' The tracking workbook: the one this code lives in, or if it lives somewhere else
' (PERSONAL, say), the workbook in front.
Public Function Book() As Workbook
    If Not FindRefundsSheet(ThisWorkbook) Is Nothing Then
        Set Book = ThisWorkbook
    ElseIf Not ActiveWorkbook Is Nothing Then
        If Not FindRefundsSheet(ActiveWorkbook) Is Nothing Then Set Book = ActiveWorkbook
    End If
End Function

' Adds what it needs the first time, without stopping: the SETUP sheet, Short Name and NET on
' Corp License Number, and any of MF-12 Sent, MF-12 Signed, Filed Date and Refund Receive Date
' that Refunds Due is missing. Hands back a reason to stop, blank when it can go on.
Public Function EnsureSetup() As String
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

    Set rd = DueSheet(hr, lastCol)
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
    ws.Range("B13").NumberFormat = "@"    ' Aug-26 stays a month, not August 26
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
    ws.Range("C11").Value = PACKETS_HINT
    SetupLine ws, 12, "MF-12 template file", ""
    ws.Range("C12").Value = "Full path to " & TEMPLATE_NAME & ". Found by itself when it's beside this workbook."
    SetupLine ws, 13, "Month to work", ""
    ws.Range("C13").Value = MONTH_NOTE
    SetupLine ws, 15, "Email subject", ">MONTH< >YEAR< GA MF-12 Waiver - >US< / >SUPPLIER<"
    SetupLine ws, 16, "Email text", NewLetter()
    ws.Range("B16").WrapText = True
    ws.Range("C16").Value = LETTER_NOTE
    ws.Range("C16").WrapText = True
    ws.Range("C15").Value = FILL_NOTE
    r = 18
    ws.Cells(r, 1).Value = "Refund type (as the report says it)"
    ws.Cells(r, 2).Value = TYPE_WORD_HEAD
    ws.Cells(r, 3).Value = "Words for the email"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = GRAY
    ws.Cells(r + 1, 1).Value = "6 - Disbursements: Licensed - Tax Unpaid"
    ws.Cells(r + 1, 2).Value = "Licensed"
    ws.Cells(r + 1, 3).Value = "licensed customer sales"
    ws.Cells(r + 2, 1).Value = "8 - Disbursements: Federal Government"
    ws.Cells(r + 2, 2).Value = "Government"
    ws.Cells(r + 2, 3).Value = "government customer sales"
    ws.Range(ws.Cells(r + 3, 1), ws.Cells(r + 6, 3)).Interior.Color = YELLOW
    ws.Cells(r + 8, 1).Value = TYPE_NOTE
    ws.Range("A1:C31").VerticalAlignment = xlTop
    ws.Rows(16).AutoFit
End Sub

' A SETUP sheet from before the running list: its Month folder line becomes Packets folder,
' and a Month to work line goes in under the template line. The last version's Oldest period
' line becomes Month to work, blank, so the click asks. One from before the new names gets
' them in its refund type list.
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
            ws.Cells(r, 3).Value = PACKETS_HINT
        End If
    End If
    r = SetupRow(ws, "packets folder")
    If r > 0 Then
        If InStr(CStr(ws.Cells(r, 3).Value), "2026-08 AUGUST") > 0 Then ws.Cells(r, 3).Value = PACKETS_HINT
    End If
    UpgradeTypeList ws
    UpgradeLetter ws
    If SetupRow(ws, "month to work") = 0 Then
        r = SetupRow(ws, "oldest period")
        If r > 0 Then
            ws.Cells(r, 1).Value = "Month to work"
            ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 2).Value = ""
            ws.Cells(r, 3).Value = MONTH_NOTE
        Else
            r = SetupRow(ws, "mf-12 template")
            If r > 0 Then
                ws.Rows(r + 1).Insert
                SetupLine ws, r + 1, "Month to work", ""
                ws.Cells(r + 1, 2).NumberFormat = "@"
                ws.Cells(r + 1, 3).Value = MONTH_NOTE
            End If
        End If
    End If
End Sub

' A refund type list from before the new names: its Folder prefix column becomes the name in
' folders and files. LIC is Licensed and GOV is Government; any other prefix moves to an Old
' folder prefix column, so its old folders are still found, and its name is left for you to
' type (CHECK says so until it's there).
Private Sub UpgradeTypeList(ws As Worksheet)
    Dim r As Long, v As String
    r = SetupRow(ws, "refund type")
    If r = 0 Then Exit Sub
    If Squash(CStr(ws.Cells(r, 2).Value)) <> "folder prefix" Then Exit Sub
    ws.Cells(r, 2).Value = TYPE_WORD_HEAD
    r = r + 1
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        v = UCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If v = "LIC" Then
            ws.Cells(r, 2).Value = "Licensed"
        ElseIf v = "GOV" Then
            ws.Cells(r, 2).Value = "Government"
        ElseIf Len(v) > 0 Then
            ws.Cells(SetupRow(ws, "refund type"), 4).Value = OLD_PREFIX_HEAD
            ws.Cells(SetupRow(ws, "refund type"), 4).Font.Bold = True
            ws.Cells(SetupRow(ws, "refund type"), 4).Interior.Color = GRAY
            ws.Cells(r, 4).Value = v
            ws.Cells(r, 2).Value = ""
            ws.Cells(r, 2).Interior.Color = YELLOW
        End If
        r = r + 1
    Loop
    r = SetupRow(ws, "a refund type the list")
    If r > 0 Then ws.Cells(r, 1).Value = TYPE_NOTE
End Sub

' A SETUP still holding the last version's letter gets the new one; a letter you changed stays
' as you wrote it. The notes beside them name the new fill-ins unless you wrote your own.
Private Sub UpgradeLetter(ws As Worksheet)
    Dim r As Long, v As String
    r = SetupRow(ws, "email text")
    If r > 0 Then
        If Replace(Trim$(CStr(ws.Cells(r, 2).Value)), vbCr, "") = OldLetter() Then
            ws.Cells(r, 2).Value = NewLetter()
            ws.Cells(r, 2).WrapText = True
            ws.Rows(r).AutoFit
        End If
        v = CStr(ws.Cells(r, 3).Value)
        If Len(v) = 0 Or v = OLD_LETTER_NOTE Then ws.Cells(r, 3).Value = LETTER_NOTE: ws.Cells(r, 3).WrapText = True
    End If
    r = SetupRow(ws, "email subject")
    If r > 0 Then
        v = CStr(ws.Cells(r, 3).Value)
        If Len(v) = 0 Or v = OLD_FILL_NOTE Then ws.Cells(r, 3).Value = FILL_NOTE
    End If
End Sub

' The letter the last version put on SETUP, word for word.
Private Function OldLetter() As String
    OldLetter = "Attached is our Georgia MF-12 waiver covering >TYPES< for >MONTH< >YEAR<, " & _
        "along with copies of the supporting invoices." & vbLf & _
        "Please sign, notarize, and return each attached MF-12 at your earliest convenience." & vbLf & vbLf & _
        "Thank you very much."
End Function

Private Function NewLetter() As String
    NewLetter = ">GREETING<" & vbLf & vbLf & _
        "Attached are our Georgia MF-12 waivers for >MONTH< >YEAR<, with copies of the supporting invoices>COUNTS<" & vbLf & vbLf & _
        "Please sign, notarize, and return each attached MF-12 at your earliest convenience." & vbLf & vbLf & _
        "Thank you very much."
End Function

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

Public Sub SetSetup(ws As Worksheet, ByVal labelStart As String, ByVal value As String)
    Dim r As Long
    r = SetupRow(ws, labelStart)
    If r = 0 Then Exit Sub
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = value
End Sub

' Returns a list of what's missing, blank when all good.
Public Function ReadSetup() As String
    Dim ws As Worksheet, miss As String, r As Long, t As String, m As Long, y As Long, hasOld As Boolean

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
    mMonth = 0
    t = PeriodText(ws.Cells(SetupRow(ws, "month to work"), 2).Value)

    If Len(mBuyer) = 0 Then miss = miss & "  Purchaser name" & vbLf
    If Len(mAddress) = 0 Then miss = miss & "  Purchaser street address" & vbLf
    If Len(mCity) = 0 Then miss = miss & "  Purchaser city" & vbLf
    If Len(mState) <> 2 Then miss = miss & "  Purchaser state, 2 letters" & vbLf
    If Not (mZip Like "#####" Or mZip Like "#####-####") Then miss = miss & "  Purchaser ZIP, like 12345 or 12345-6789" & vbLf
    If Len(mUs) = 0 Then miss = miss & "  Company short name for the email subject" & vbLf
    If InStr(mBase, "\") = 0 Then miss = miss & "  Packets folder, a full path like C:\GA Refunds or \\server\share\GA Refunds" & vbLf
    If Len(mTemplatePath) = 0 Then miss = miss & "  MF-12 template file" & vbLf
    If ParsePeriod(t, m, y) Then mMonth = y * 100 + m Else miss = miss & "  Month to work, a month like Aug-26" & vbLf
    If Len(mSubject) = 0 Then miss = miss & "  Email subject" & vbLf
    If Len(mBody) = 0 Then miss = miss & "  Email text" & vbLf

    ' the refund type list: the report's words, the name in folders and files, the email's words,
    ' and on an upgraded SETUP the prefix the last version put on its folders
    r = SetupRow(ws, "refund type")
    If r > 0 Then
        hasOld = (Squash(CStr(ws.Cells(r, 4).Value)) = LCase$(OLD_PREFIX_HEAD))
        r = r + 1
        Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
            t = SafeName(CStr(ws.Cells(r, 2).Value))
            mNType = mNType + 1
            ReDim Preserve mTypeText(1 To mNType): ReDim Preserve mTypeWord(1 To mNType)
            ReDim Preserve mTypePrefix(1 To mNType): ReDim Preserve mTypePhrase(1 To mNType)
            mTypeText(mNType) = Squash(CStr(ws.Cells(r, 1).Value))
            mTypeWord(mNType) = t
            If hasOld Then mTypePrefix(mNType) = UCase$(SafeName(CStr(ws.Cells(r, 4).Value)))
            If Len(mTypePrefix(mNType)) = 0 Then
                If LCase$(t) = "licensed" Then mTypePrefix(mNType) = "LIC"
                If LCase$(t) = "government" Then mTypePrefix(mNType) = "GOV"
            End If
            mTypePhrase(mNType) = Trim$(CStr(ws.Cells(r, 3).Value))
            If Len(mTypePhrase(mNType)) = 0 Then
                miss = miss & "  Refund type list, row " & r & ": needs the words for the email" & vbLf
            End If
            r = r + 1
        Loop
    End If
    If mNType = 0 Then miss = miss & "  The refund type list (under 'Refund type (as the report says it)')" & vbLf
    ReadSetup = miss
End Function

Public Function SetupRow(ws As Worksheet, ByVal labelStart As String) As Long
    Dim r As Long
    For r = 1 To 60
        If Left$(Squash(CStr(ws.Cells(r, 1).Value)), Len(labelStart)) = labelStart Then SetupRow = r: Exit Function
    Next r
End Function

Public Function SetupValue(ws As Worksheet, ByVal labelStart As String) As String
    Dim r As Long
    r = SetupRow(ws, labelStart)
    If r > 0 Then SetupValue = Trim$(CStr(ws.Cells(r, 2).Value))
End Function


' ======================================================================================
'  Reading the tabs
' ======================================================================================

' Empties what a read fills in. newClick empties SETUP's refund type list and the click's
' counts too; a click reads the tabs again partway (after Mark filed) but SETUP only once.
Public Sub ClearLists(Optional ByVal newClick As Boolean = False)
    Erase mRow, mSup, mMail, mPkt, mLook
    mNRow = 0: mNSup = 0: mNMail = 0: mNPkt = 0: mNLook = 0: mByHand = 0: mPeriodNote = ""
    If newClick Then
        Erase mTypeText, mTypeWord, mTypePrefix, mTypePhrase
        mNType = 0: mShortsFilled = 0: mPaidStuck = ""
    End If
End Sub

' Reads the contacts and the open rows, and sorts the rows into packets and emails.
Public Sub LoadAll()
    mStage = "reading Corp License Number"
    LoadContacts
    mStage = "reading Refunds Due"
    LoadRows
    mStage = "matching rows to suppliers"
    AssignPackets
    mStage = "writing the Corp names onto Refunds Due"
    WriteCorpNames
    PaidStuckLooks
End Sub

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

' Refunds Due, with its header row and last header column.
Private Function DueSheet(ByRef hr As Long, ByRef lastCol As Long) As Worksheet
    Set DueSheet = FindRefundsSheet(mBook)
    If DueSheet Is Nothing Then Exit Function
    hr = HeaderRow(DueSheet, Array(H_BOL, H_INV, H_LICENSE))
    lastCol = LastHeaderCol(DueSheet, hr)
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
                    If cShort > 0 Then .ShortName = SafeName(CellStr(ws.Cells(r, cShort).Value))
                    If cNet > 0 Then
                        n = LCase$(CellStr(ws.Cells(r, cNet).Value))
                        .IsNet = (n = "y" Or n = "yes" Or n = "x" Or n = "net" Or n = "true")
                    End If
                    If Len(.ShortName) > 0 Then taken = taken & "|" & UCase$(.ShortName) & "|"
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
                    taken = taken & "|" & UCase$(.ShortName) & "|"
                    mShortsFilled = mShortsFilled + 1
                End If
            End If
            If Len(.ShortName) = 0 Then .Problem = "No Short Name on " & ws.Name & " row " & .SheetRow
            If Len(.Corp) = 0 Then .Problem = "No Corp name on " & ws.Name & " row " & .SheetRow & " (it goes on the MF-12 as the seller)"
        End With
    Next i

    ' two licenses can't share a short name, or their folders would collide. Windows doesn't
    ' tell Echo from ECHO, so neither does this.
    For i = 1 To mNSup
        For j = i + 1 To mNSup
            If Len(mSup(i).ShortName) > 0 And UCase$(mSup(i).ShortName) = UCase$(mSup(j).ShortName) Then
                mSup(i).Problem = "Short Name " & mSup(i).ShortName & " is also on row " & mSup(j).SheetRow
                mSup(j).Problem = "Short Name " & mSup(j).ShortName & " is also on row " & mSup(i).SheetRow
            End If
        Next j
    Next i
End Sub

' Acme from Acme Oil Company: the Corp name's first word as typed, or its first two if
' another supplier already has the one word, then a number on the end if those are taken too.
' taken holds the Short Names already used, in capitals.
Private Function GuessShort(ByVal corp As String, ByVal taken As String) As String
    Dim w As Variant, s As String, i As Long, c As String, n As Long
    For i = 1 To Len(corp)
        c = Mid$(corp, i, 1)
        If c Like "[A-Za-z0-9]" Then s = s & c Else s = s & " "
    Next i
    s = Trim$(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    If Len(s) = 0 Then Exit Function
    w = Split(s, " ")
    i = 0
    If UBound(w) > 0 And UCase$(w(0)) = "THE" Then i = 1
    s = w(i)
    If InStr(taken, "|" & UCase$(s) & "|") > 0 And UBound(w) > i Then s = s & " " & w(i + 1)
    If InStr(taken, "|" & UCase$(s) & "|") > 0 Then
        n = 2
        Do While InStr(taken, "|" & UCase$(s) & " " & n & "|") > 0: n = n + 1: Loop
        s = s & " " & n
    End If
    GuessShort = s
End Function

' Every row on Refunds Due that's still open (no Filed Date, no Refund Receive Date) from the
' month to work on SETUP, but not one sent by hand. Other months' rows are left alone; only
' their BOLs are kept, so their INBOX files wait for their month.
Private Sub LoadRows()
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long, data As Variant, ok As Boolean
    Dim cType As Long, cPer As Long, cCorp As Long, cLic As Long, cBol As Long, cInv As Long, cInvDate As Long
    Dim cGal As Long, cRef As Long, cAllow As Long, cTtl As Long, cFuel As Long, cSent As Long, cSigned As Long
    Dim cFiled As Long, cRecv As Long, key As Long, odd As Object, k As Variant, d As Date, per As String
    Dim byDate As Boolean, bk As String

    Set mOtherBols = CreateObject("Scripting.Dictionary")
    Set ws = DueSheet(hr, lastCol)
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
            byDate = False
            If key = 0 Then
                d = ToDate(data(r, cInvDate), ok)
                If ok Then key = Year(d) * 100 + Month(d): byDate = True
            End If
            ' a row with no month to go by stays in, so its line on CHECK says what to fix
            If key > 0 And key <> mMonth Then
                bk = BolKey(CellStr(data(r, cBol)))
                If Len(bk) > 0 And Not mOtherBols.Exists(bk) Then mOtherBols.Add bk, key
            ElseIf IsHandRow(hr + r) Then
                mByHand = mByHand + 1
            Else
                If byDate Then
                    If odd.Exists(per) Then odd(per) = odd(per) & ", " & (hr + r) Else odd.Add per, CStr(hr + r)
                End If
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
    ' not a fix, only worth knowing: it goes in the small line under CHECK's bar
    For Each k In odd.Keys
        mPeriodNote = mPeriodNote & IIf(Len(mPeriodNote) > 0, "  ", "") & "Couldn't read the period " & _
            IIf(Len(k) > 0, "'" & k & "'", "(it's blank)") & " on row" & IIf(InStr(odd(k), ",") > 0, "s ", " ") & odd(k) & _
            ", so it went by the invoice date; fix Period there if that's wrong."
    Next k
End Sub

' A row's month as 202608: its Period, or its invoice date when the Period can't be read. 0 for neither.
Private Function RowMonth(ws As Worksheet, ByVal r As Long, ByVal cPer As Long, ByVal cInvDate As Long) As Long
    Dim d As Date, ok As Boolean
    If cPer > 0 Then RowMonth = PeriodKey(PeriodText(ws.Cells(r, cPer).Value))
    If RowMonth = 0 And cInvDate > 0 Then
        d = ToDate(ws.Cells(r, cInvDate).Value, ok)
        If ok Then RowMonth = Year(d) * 100 + Month(d)
    End If
End Function

' True for a row this click found sent by hand: MF-12 Sent filled in, no packet from this macro.
Private Function IsHandRow(ByVal sheetRow As Long) As Boolean
    If mHandRows Is Nothing Then Exit Function
    IsHandRow = mHandRows.Exists(CStr(sheetRow))
End Function

' The last row with anything in a header column, for reading a sheet down to its end.
Private Function LastDataRow(ws As Worksheet, ByVal hr As Long, ByVal lastCol As Long) As Long
    Dim c As Long, r As Long
    LastDataRow = hr
    For c = 1 To lastCol
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > LastDataRow Then LastDataRow = r
    Next c
End Function

' The months of the rows still open on Refunds Due, oldest first, for the month question.
Public Function OpenPeriods() As Collection
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long, data As Variant
    Dim cPer As Long, cBol As Long, cLic As Long, cFiled As Long, cRecv As Long, key As Long, i As Long, done As Boolean
    Dim out As New Collection

    Set OpenPeriods = out
    Set ws = DueSheet(hr, lastCol)
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
            ElseIf Len(mTypeWord(t)) = 0 Then
                .Problem = "Refund type '" & .TypeText & "' has no name for its folders on SETUP"
                t = 0
            Else
                t = FindWord(mTypeWord(t))
                .TypeNo = t
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
                If .HasSent Then
                    mMail(m).IsSent = True
                    If .SentOn > mMail(m).SentOn Then mMail(m).SentOn = .SentOn
                End If
            End If

            ' a packet is a period, a refund type and a license; the Short Name only labels it
            If m > 0 And t > 0 Then
                If Len(mSup(.Sup).ShortName) > 0 Then
                    key = PacketKey(.PerKey, mSup(.Sup).LicKey, mTypeWord(t))
                    p = FindPkt(key)
                    If p = 0 Then
                        mNPkt = mNPkt + 1
                        ReDim Preserve mPkt(1 To mNPkt)
                        p = mNPkt
                        mPkt(p).Key = key
                        mPkt(p).PName = mSup(.Sup).ShortName & " " & mTypeWord(t)
                        mPkt(p).Word = mTypeWord(t)
                        mPkt(p).TypeNo = t
                        mPkt(p).Prefix = mTypePrefix(t)
                        mPkt(p).PerKey = .PerKey
                        mPkt(p).Sup = .Sup
                        mPkt(p).Mail = m
                    End If
                    .Pkt = p
                    mPkt(p).NRows = mPkt(p).NRows + 1
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
            If InStr(mSup(s).Problem, " is also on row ") > 0 Then
                k = "Give each its own Short Name on " & CONTACTS_SHEET & ", then click the button again."
            Else
                k = "Fill it in on " & CONTACTS_SHEET & "."
            End If
            AddLook mSup(s).Problem, "License " & mSup(s).License & ", " & mSup(s).NRows & " open row(s)", k
        End If
    Next s
    For p = 1 To mNPkt
        mPkt(p).NForms = Int((mPkt(p).NRows + ROWS_PER_FORM - 1) / ROWS_PER_FORM)
        ' a supplier that needs a fix holds its packets too, so two suppliers' files never mix
        If Len(mPkt(p).Problem) = 0 And Len(mSup(mPkt(p).Sup).Problem) > 0 Then mPkt(p).Problem = mSup(mPkt(p).Sup).Problem
    Next p
End Sub

' What a packet is: "202608|12345GR|LICENSED".
Public Function PacketKey(ByVal perKey As Long, ByVal licKey As String, ByVal tw As String) As String
    PacketKey = perKey & "|" & licKey & "|" & UCase$(tw)
End Function

Private Function RowFix(ByVal problem As String) As String
    If InStr(problem, "isn't on " & CONTACTS_SHEET) > 0 Then
        RowFix = "Add the license to " & CONTACTS_SHEET & " with its contact, Short Name and NET."
    ElseIf InStr(problem, "SETUP list") > 0 Then
        RowFix = "Add it to the refund type list on SETUP (a name like Diversion, words like diversions)."
    ElseIf InStr(problem, "no name for its folders") > 0 Then
        RowFix = "Type its name for the folders and files in the refund type list on SETUP, like Diversion."
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
    Dim ws As Worksheet, hr As Long, lastCol As Long, c As Long, i As Long, nm As String
    Set ws = DueSheet(hr, lastCol)
    If ws Is Nothing Then Exit Sub
    c = FindCol(ws, hr, H_CORP, lastCol)
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
'  Rows that are paid, filed, sent and signed
' ======================================================================================

' The month to work's rows with a Refund Receive Date move to the bottom of Refunds Received,
' matched column by column on the headers, and come off Refunds Due. Safe to cut off halfway:
' a row that's already on Refunds Received (same license, BOL, invoice number and period) isn't
' copied again, and a row only comes off Refunds Due once its copy is there. A row Refunds
' Received won't take stays on Refunds Due with a line on CHECK; the rest still move. Hands
' back how many came off.
Public Function MoveReceived() As Long
    Dim rd As Worksheet, rr As Worksheet, hr As Long, rhr As Long, lastCol As Long, rLast As Long, cRecv As Long
    Dim lastRow As Long, r As Long, c As Long, dest As Long, map() As Long, moving() As Long, n As Long, i As Long, nm As String
    Dim have As Object, key As String, gone() As Boolean, cPer As Long, cInvDate As Long, mo As Long

    Set rd = DueSheet(hr, lastCol)
    cRecv = FindCol(rd, hr, H_RECEIVED, lastCol)
    If cRecv = 0 Then Exit Function
    cPer = FindCol(rd, hr, H_PERIOD, lastCol)
    cInvDate = FindCol(rd, hr, H_DATE, lastCol)
    lastRow = LastDataRow(rd, hr, lastCol)
    For r = hr + 1 To lastRow
        mo = RowMonth(rd, r, cPer, cInvDate)
        If Len(CellStr(rd.Cells(r, cRecv).Value)) > 0 And (mo = mMonth Or mo = 0) Then
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
    Set have = PaidRowsThere(rr, rhr, rLast, dest - 1)
    ReDim gone(1 To n)
    For i = 1 To n
        key = PaidRowKey(rd, moving(i), hr, lastCol)
        If have.Exists(key) Then
            gone(i) = True
        ElseIf CopyPaidRow(rd, rr, moving(i), dest, map, hr, lastCol, rLast) Then
            gone(i) = True
            dest = dest + 1
            have(key) = True
        End If
    Next i
    ' each of these is on Refunds Received now, so it can come off Refunds Due
    For i = n To 1 Step -1
        If gone(i) Then
            rd.Rows(moving(i)).Delete
            MoveReceived = MoveReceived + 1
        End If
    Next i
End Function

' One paid row onto Refunds Received, under its headers. False if that tab won't take it (it's
' protected, say); then whatever went in is cleared and CHECK says the row stays on Refunds Due.
' map gives each Refunds Due column's column on Refunds Received, 0 for none.
Private Function CopyPaidRow(rd As Worksheet, rr As Worksheet, ByVal src As Long, ByVal dest As Long, map As Variant, _
    ByVal hr As Long, ByVal lastCol As Long, ByVal rLast As Long) As Boolean
    Dim c As Long, v As Variant, why As String
    On Error GoTo Failed
    For c = 1 To lastCol
        If map(c) > 0 Then
            v = rd.Cells(src, c).Value
            ' text stays text, so a period like Aug-26 doesn't turn into a date
            If VarType(v) = vbString Then
                rr.Cells(dest, map(c)).NumberFormat = "@"
            Else
                rr.Cells(dest, map(c)).NumberFormat = rd.Cells(src, c).NumberFormat
            End If
            rr.Cells(dest, map(c)).Value = v
        End If
    Next c
    CopyPaidRow = True
    Exit Function
Failed:
    why = Err.Description
    Resume Undo
Undo:
    On Error Resume Next
    rr.Range(rr.Cells(dest, 1), rr.Cells(dest, rLast)).ClearContents
    On Error GoTo 0
    ' the tabs are read after the move, which starts CHECK's list over, so the line goes in then
    mPaidStuck = mPaidStuck & "BOL " & CellStr(rd.Cells(src, FindCol(rd, hr, H_BOL, lastCol)).Value) & " (license " & _
        CellStr(rd.Cells(src, FindCol(rd, hr, H_LICENSE, lastCol)).Value) & ": " & why & ")" & vbLf
End Function

' A line on CHECK for every paid row Refunds Received couldn't take this click.
Private Sub PaidStuckLooks()
    Dim x As Variant
    For Each x In Split(mPaidStuck, vbLf)
        If Len(x) > 0 Then AddLookItem "Refunds Received couldn't take the paid row for", _
            "Refunds Received couldn't take the paid rows for", CStr(x), "Each stays on Refunds Due for now.", _
            "Make room at the bottom of Refunds Received (unprotect the tab, or clear what's in the way), then click Build packets again."
    Next x
End Sub

' The rows already on Refunds Received, by license, BOL, invoice number and period.
Private Function PaidRowsThere(rr As Worksheet, ByVal rhr As Long, ByVal rLast As Long, ByVal lastRow As Long) As Object
    Dim have As Object, r As Long
    Set have = CreateObject("Scripting.Dictionary")
    For r = rhr + 1 To lastRow
        have(PaidRowKey(rr, r, rhr, rLast)) = True
    Next r
    Set PaidRowsThere = have
End Function

' "10004GR|94003|940030|202608": what makes a paid row the same row on either tab.
Private Function PaidRowKey(ws As Worksheet, ByVal r As Long, ByVal hr As Long, ByVal lastCol As Long) As String
    Dim c As Long, k As String, per As String
    c = FindCol(ws, hr, H_LICENSE, lastCol)
    If c > 0 Then k = LicenseKey(CellStr(ws.Cells(r, c).Value))
    c = FindCol(ws, hr, H_BOL, lastCol)
    If c > 0 Then k = k & "|" & BolKey(CellStr(ws.Cells(r, c).Value))
    c = FindCol(ws, hr, H_INV, lastCol)
    If c > 0 Then k = k & "|" & UCase$(CellStr(ws.Cells(r, c).Value))
    c = FindCol(ws, hr, H_PERIOD, lastCol)
    If c > 0 Then
        per = PeriodText(ws.Cells(r, c).Value)
        If PeriodKey(per) > 0 Then k = k & "|" & PeriodKey(per) Else k = k & "|" & Squash(per)
    End If
    PaidRowKey = k
End Function

' MF-12 Sent from the email in Sent Items, MF-12 Signed from the NOTARIZED copy's date, on
' each row where it's still blank. A date you typed is never changed.
Public Sub WriteStamps()
    Dim ws As Worksheet, hr As Long, lastCol As Long, cSent As Long, cSigned As Long, i As Long, m As Long, p As Long
    Set ws = DueSheet(hr, lastCol)
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

' Today's date as Filed Date on every row of one packet.
Public Sub StampFiled(ByVal p As Long)
    Dim ws As Worksheet, hr As Long, lastCol As Long, cFiled As Long, i As Long
    Set ws = DueSheet(hr, lastCol)
    cFiled = FindCol(ws, hr, H_FILED, lastCol)
    For i = 1 To mNRow
        If mRow(i).Pkt = p Then StampCell ws.Cells(mRow(i).SheetRow, cFiled), Date
    Next i
End Sub

Private Sub StampCell(c As Range, ByVal d As Date)
    c.NumberFormat = "m/d/yyyy"
    c.Value = DateSerial(Year(d), Month(d), Day(d))
End Sub


' ======================================================================================
'  What goes on each MF-12 and in each email
' ======================================================================================

' The packet's rows in the order they go on the form: by invoice date, then BOL.
Public Function PacketRows(ByVal p As Long) As Variant
    Dim idx() As Long, n As Long, i As Long, v As Variant
    For i = 1 To mNRow
        If mRow(i).Pkt = p Then n = n + 1: ReDim Preserve idx(1 To n): idx(n) = i
    Next i
    v = idx
    If n > 1 Then SortList v, SORT_ROWS
    PacketRows = v
End Function

' The invoice numbers on a BOL's rows, joined with + (7890+7891), for a filed invoice's name;
' only packet p's rows, or every open row when p is 0. shown is the BOL as a row has it.
Public Function BolInvoices(ByVal bolK As String, ByVal p As Long, ByRef shown As String) As String
    Dim r As Long, invs As String
    shown = ""
    For r = 1 To mNRow
        If mRow(r).BolK = bolK And (p = 0 Or mRow(r).Pkt = p) Then
            If InStr("+" & invs & "+", "+" & mRow(r).Inv & "+") = 0 Then invs = invs & IIf(Len(invs) > 0, "+", "") & mRow(r).Inv
            shown = mRow(r).Bol
        End If
    Next r
    BolInvoices = invs
End Function

Public Function FormValues(ByVal p As Long, ByVal formNo As Long) As Object
    Dim v As Object, ix As Variant, i As Long, slot As Long, k As String, s As Long
    Set v = CreateObject("Scripting.Dictionary")
    s = mPkt(p).Sup
    v("SELLER") = Array(mSup(s).Corp, mSup(s).Corp)       ' exactly as typed on Corp License Number
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

' A short fingerprint of everything that goes on one form, so a later run can tell if it changed.
' The seller counts in capitals, the way forms made before kept it, so a form only goes stale
' when what's on it changes, not its capitals.
Public Function FormFingerprint(ByVal p As Long, ByVal formNo As Long) As String
    Dim v As Object, k As Variant, h As Double, t As String, i As Long, n As Long
    Set v = FormValues(p, formNo)
    For Each k In v.Keys
        If Left$(k, 1) = "R" Then n = n + 1
        If k = "SELLER" Then t = k & "=" & UCase$(v(k)(0)) & ";" Else t = k & "=" & v(k)(0) & ";"
        For i = 1 To Len(t)
            h = h * 131 + AscW(Mid$(t, i, 1))
            h = h - Int(h / 4294967291#) * 4294967291#
        Next i
    Next k
    FormFingerprint = "v1-" & (n \ 5) & "-" & Format$(h, "0")
End Function

' A number the way PDF wants it: dot for the decimal, no thousands commas, no trailing zeros past what's asked.
Public Function PdfNum(ByVal x As Double, ByVal places As Long) As String
    Dim s As String
    s = Replace(Format$(x, "0." & String$(places, "0")), ",", ".")
    If places = 3 Then
        Do While Right$(s, 1) = "0": s = Left$(s, Len(s) - 1): Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If
    PdfNum = s
End Function

' The words with the email's fill-ins put in. asName stands in for the supplier's Short Name,
' to build a subject the way it was under another name.
Public Function FillWords(ByVal t As String, ByVal m As Long, Optional ByVal asName As String = "") As String
    If Len(asName) = 0 Then asName = mSup(mMail(m).Sup).ShortName
    t = Replace(t, ">MONTH<", MonthName(mMail(m).PerKey Mod 100))
    t = Replace(t, ">YEAR<", CStr(mMail(m).PerKey \ 100))
    t = Replace(t, ">SUPPLIER<", asName)
    t = Replace(t, ">US<", mUs)
    t = Replace(t, ">GREETING<", Greeting(mMail(m).Sup))
    t = Replace(t, ">COUNTS<", CountWords(m))
    t = Replace(t, ">CONTACT<", mSup(mMail(m).Sup).Contact)
    t = Replace(t, ">TYPES<", TypeWords(m))
    FillWords = t
End Function

' "Hi Jane," from the contact's first name, or "Hello," when there's no contact name.
Private Function Greeting(ByVal s As Long) As String
    Dim c As String
    c = Trim$(mSup(s).Contact)
    If Len(c) = 0 Then Greeting = "Hello,": Exit Function
    If InStr(c, " ") > 0 Then c = Left$(c, InStr(c, " ") - 1)
    Greeting = "Hi " & c & ","
End Function

' What the email attaches, by refund type in the SETUP list's order. With two types or more, a
' colon and a line each ("Licensed customer sales: 2 MF-12s, 28 invoices"); with one, the end
' of the sentence (", covering licensed customer sales: 1 MF-12, 12 invoices.").
Private Function CountWords(ByVal m As Long) As String
    Dim t As Long, p As Long, i As Long, n As Long, k As Long, ph As String
    Dim words() As String, forms() As Long, invs() As Long
    For t = 1 To mNType
        ph = mTypePhrase(t)
        If Len(ph) = 0 Then ph = LCase$(mTypeWord(t))
        For p = 1 To mNPkt
            If mPkt(p).Mail = m And mPkt(p).TypeNo = t Then
                k = 0
                For i = 1 To n
                    If words(i) = ph Then k = i
                Next i
                If k = 0 Then
                    n = n + 1: k = n
                    ReDim Preserve words(1 To n): ReDim Preserve forms(1 To n): ReDim Preserve invs(1 To n)
                    words(k) = ph
                End If
                forms(k) = forms(k) + mPkt(p).MailForms
                invs(k) = invs(k) + mPkt(p).MailInvoices
            End If
        Next p
    Next t
    If n = 0 Then
        CountWords = "."
    ElseIf n = 1 Then
        CountWords = ", covering " & words(1) & ": " & Counted(forms(1), invs(1)) & "."
    Else
        CountWords = ":"
        For i = 1 To n
            CountWords = CountWords & vbLf & UCase$(Left$(words(i), 1)) & Mid$(words(i), 2) & ": " & Counted(forms(i), invs(i))
        Next i
    End If
End Function

Private Function Counted(ByVal forms As Long, ByVal invs As Long) As String
    Counted = forms & " MF-12" & IIf(forms = 1, "", "s") & ", " & invs & " invoice" & IIf(invs = 1, "", "s")
End Function

' "licensed customer sales and government customer sales", in the SETUP list's order.
Public Function TypeWords(ByVal m As Long) As String
    Dim t As Long, p As Long, words() As String, n As Long, i As Long, has As Boolean
    For t = 1 To mNType
        has = False
        For p = 1 To mNPkt
            If mPkt(p).Mail = m And UCase$(mPkt(p).Word) = UCase$(mTypeWord(t)) Then has = True
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

' Who a supplier's email goes to: Email address, else Group Email Address, else Additional
' Email. The others go on CC.
Public Sub Addresses(ByVal s As Long, ByRef toAddr As String, ByRef ccAddr As String)
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


' ======================================================================================
'  The order things are listed in
' ======================================================================================

' Packets in the order to key them: oldest period first, then by refund type in the SETUP list's
' order (one period's type is one MFD-33 filing, so its MF-12s sit together), then by supplier
' short name inside the type.
Public Function PacketOrder() As Variant
    PacketOrder = Numbered(mNPkt, SORT_PACKETS)
End Function

' Emails oldest period first, then by supplier short name.
Public Function MailOrder() As Variant
    MailOrder = Numbered(mNMail, SORT_MAILS)
End Function

' 1 to n, sorted as that kind of list. Array() when n is 0.
Private Function Numbered(ByVal n As Long, ByVal kind As Long) As Variant
    Dim idx() As Long, i As Long, v As Variant
    If n = 0 Then Numbered = Array(): Exit Function
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    v = idx
    SortList v, kind
    Numbered = v
End Function

' The one sort: an insertion sort, smallest first, in place. kind says what the list holds and
' how two of them compare (see SORT_ROWS and the rest).
Public Sub SortList(a As Variant, ByVal kind As Long)
    Dim i As Long, j As Long, t As Variant
    For i = LBound(a) + 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If Not ComesFirst(kind, t, a(j)) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

Private Function ComesFirst(ByVal kind As Long, ByVal a As Variant, ByVal b As Variant) As Boolean
    Select Case kind
        Case SORT_ROWS
            If mRow(a).InvDate <> mRow(b).InvDate Then
                ComesFirst = mRow(a).InvDate < mRow(b).InvDate
            ElseIf mRow(a).BolK <> mRow(b).BolK Then
                ComesFirst = mRow(a).BolK < mRow(b).BolK
            Else
                ComesFirst = mRow(a).SheetRow < mRow(b).SheetRow
            End If
        Case SORT_PACKETS
            If mPkt(a).PerKey <> mPkt(b).PerKey Then
                ComesFirst = (mPkt(a).PerKey < mPkt(b).PerKey)
            ElseIf mPkt(a).TypeNo <> mPkt(b).TypeNo Then
                ComesFirst = (mPkt(a).TypeNo < mPkt(b).TypeNo)
            Else
                ComesFirst = (mSup(mPkt(a).Sup).ShortName < mSup(mPkt(b).Sup).ShortName)
            End If
        Case SORT_MAILS
            If mMail(a).PerKey <> mMail(b).PerKey Then
                ComesFirst = (mMail(a).PerKey < mMail(b).PerKey)
            Else
                ComesFirst = (mSup(mMail(a).Sup).ShortName < mSup(mMail(b).Sup).ShortName)
            End If
        Case Else
            ComesFirst = (UCase$(a) < UCase$(b))
    End Select
End Function


' ======================================================================================
'  Odds and ends
' ======================================================================================

Public Sub AddLook(ByVal what As String, ByVal detail As String, ByVal howTo As String)
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

' A problem that can come up for many things at once, like INBOX files no row has the BOL of,
' is one line naming each of them. one is how the line starts for one thing, many for more:
' "INBOX: no open row has the BOLs of 50001 VI.pdf, 59001 VI.pdf and 77777 VI.pdf". Things
' with the same start but a different fix get a line each.
Public Sub AddLookItem(ByVal one As String, ByVal many As String, ByVal thing As String, ByVal detail As String, _
    ByVal howTo As String)
    Dim i As Long, k As Long, parts As Variant, txt As String
    For i = 1 To mNLook
        If mLook(i).Starts = one And mLook(i).HowTo = howTo Then Exit For
    Next i
    If i > mNLook Then
        mNLook = mNLook + 1
        ReDim Preserve mLook(1 To mNLook)
        i = mNLook
        mLook(i).Starts = one
        mLook(i).Many = many
        mLook(i).Detail = detail
        mLook(i).HowTo = howTo
    ElseIf InStr(vbLf & mLook(i).Items & vbLf, vbLf & thing & vbLf) > 0 Then
        Exit Sub
    End If
    mLook(i).Items = mLook(i).Items & IIf(Len(mLook(i).Items) > 0, vbLf, "") & thing
    parts = Split(mLook(i).Items, vbLf)
    For k = 0 To UBound(parts)
        If k = 0 Then
            txt = parts(k)
        ElseIf k = UBound(parts) Then
            txt = txt & " and " & parts(k)
        Else
            txt = txt & ", " & parts(k)
        End If
    Next k
    mLook(i).What = IIf(UBound(parts) = 0, one, many) & " " & txt
End Sub

Public Function SheetNamed(ByVal nm As String) As Worksheet
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

Public Function Squash(ByVal s As String) As String
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

Public Function PeriodText(ByVal v As Variant) As String
    If VarType(v) = vbDate Then PeriodText = Format$(v, "mmm-yy") Else PeriodText = CellStr(v)
End Function

' 202608 from Aug-26, 08/2026, 2026-08 and the like. 0 if it isn't a month.
Public Function PeriodKey(ByVal t As String) As Long
    Dim m As Long, y As Long
    If ParsePeriod(t, m, y) Then PeriodKey = y * 100 + m
End Function

' Aug-26 from 202608.
Public Function PeriodLabel(ByVal key As Long) As String
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

' The first line of the refund type list with this name in folders and files. Two report
' wordings with one name share a packet.
Private Function FindWord(ByVal tw As String) As Long
    Dim i As Long
    For i = 1 To mNType
        If UCase$(mTypeWord(i)) = UCase$(tw) Then FindWord = i: Exit Function
    Next i
End Function

Public Function LicenseKey(ByVal s As String) As String
    LicenseKey = UCase$(Replace(Replace(Trim$(s), " ", ""), "-", ""))
End Function

' BOLs compared without spaces or case, and without leading zeros when they're all digits.
Public Function BolKey(ByVal s As String) As String
    s = UCase$(Replace(Trim$(s), " ", ""))
    If IsAllDigits(s) Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If
    BolKey = s
End Function

Public Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "#" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Public Function IsAllLetters(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[A-Za-z]" Then Exit Function
    Next i
    IsAllLetters = True
End Function

' Drops the characters Windows won't take in a file name.
Public Function SafeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "")
    Next c
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SafeName = Trim$(s)
End Function
