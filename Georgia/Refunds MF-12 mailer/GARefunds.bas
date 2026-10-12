' GA refunds: builds each supplier's MF-12 packet and an Outlook draft to send it with.
'
' Lives in the GA tracking workbook. Each month:
'   1. Paste the refund report into Refunds Due, like always.
'   2. Pull the invoices and save them into the month's INBOX folder as "<BOL> VI" and
'      "<BOL> CI". Any order is fine.
'   3. Alt F8, run Refunds1_BuildPackets. It makes the month's folders, renames and files
'      everything in INBOX into its packet folder, fills each packet's MF-12 from the
'      report rows, and lays it all out on the CHECK sheet. Run it as often as you like
'      while the invoices come in.
'   4. When CHECK says a supplier is READY, Alt F8, run Refunds2_MakeDrafts. One Outlook
'      draft per supplier, its MF-12s and VIs attached, addressed to its contact.
'      It never sends. Look each one over in Outlook and send it yourself.
'   5. When the notarized forms are back, the ENTRY sheet is the typing sheet for the Georgia
'      portal: one block per MF-12, the columns in the portal's order, a total under each
'      block to check the portal's running total against.
'
' Suppliers are matched on the license number. Anything it can't place goes on the LOOK
' list on the CHECK sheet, and that supplier gets no draft until it's fixed.
' Every run also writes the Corp name from Corp License Number onto each matched row of
' Refunds Due, so the report's own spelling of a supplier never shows anywhere you work from.
' It never deletes your files and never writes over a file that's already there.
' Each MF-12 it makes carries a fingerprint of the rows on it, so if the report changes
' after a form is made, CHECK says so instead of letting the old form go out.
'
' The first run adds a SETUP sheet (fill the yellow cells) and two columns on Corp License
' Number: Short Name (goes in folder and file names, like MANSFIELD) and NET (Y for the
' suppliers who want the allowance taken off the amounts on their MF-12).
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
Private Const H_CONTACT As String = "contact name|contact"
Private Const H_EMAIL As String = "email address|email"
Private Const H_GROUP As String = "group email address|group email"
Private Const H_SHORT As String = "short name"
Private Const H_NET As String = "net"

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
    SheetRow As Long
    Problem As String
    NRows As Long
    Held As String           ' why no draft, blank when ready
    Status As String
    Files As String
    FileCount As Long
    MB As Double
End Type

Private Type RPkt
    PName As String          ' LIC-MANSFIELD
    Prefix As String
    Sup As Long
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
End Type

Private Type RLook
    What As String
    Detail As String
    HowTo As String
End Type

Private mRow() As RRow, mNRow As Long
Private mSup() As RSup, mNSup As Long
Private mPkt() As RPkt, mNPkt As Long
Private mLook() As RLook, mNLook As Long
Private mTypeText() As String, mTypePrefix() As String, mTypePhrase() As String, mNType As Long

Private mBook As Workbook
Private mPeriod As String
Private mMonth As Long, mYear As Long
Private mMonthDir As String
Private mInboxLeft As Long

' from SETUP
Private mBuyer As String, mAddress As String, mCity As String, mState As String, mZip As String
Private mUs As String, mPattern As String, mTemplatePath As String
Private mSubject As String, mBody As String

Private mTemplate As String      ' the template's bytes
Private mStage As String         ' what it was doing, for the error message
Private mQuiet As Boolean        ' test runs: no popups
Private mQuietPeriod As String
Private mTestDraftDir As String  ' test runs: drafts go to text files here instead of Outlook
Private mNotes As String

' ======================================================================================
'  The two macros
' ======================================================================================

Public Sub Refunds1_BuildPackets()
    RunStep 1
End Sub

Public Sub Refunds2_MakeDrafts()
    RunStep 2
End Sub

' Runs a step with no popups and hands back what it would have said, for trying it on a copy.
Public Function RefundsTest(ByVal stepNo As Long, ByVal period As String, ByVal draftDir As String) As String
    mQuiet = True
    mQuietPeriod = period
    mTestDraftDir = draftDir
    mNotes = ""
    RunStep stepNo
    mQuiet = False
    mTestDraftDir = ""
    RefundsTest = mNotes
End Function

Private Sub RunStep(ByVal stepNo As Long)
    Dim msg As String, i As Long, ready As Long, held As Long, drafted As Long

    On Error GoTo Failed
    mStage = "getting started"
    ' Excel keeps these between runs, so start every run from nothing
    Erase mRow, mSup, mPkt, mLook, mTypeText, mTypePrefix, mTypePhrase
    mNRow = 0: mNSup = 0: mNPkt = 0: mNLook = 0: mNType = 0: mTemplate = "": mInboxLeft = 0
    Set mBook = Book()
    If mBook Is Nothing Then Say "Can't find the Refunds Due sheet. Run this from the GA tracking workbook.", vbExclamation: Exit Sub

    mStage = "writing the HOW TO tab"
    WriteHowTo

    mStage = "checking the setup"
    msg = EnsureSetup()
    If Len(msg) > 0 Then Say msg & vbLf & "The HOW TO tab at the front has the two buttons and the steps.", vbInformation: Exit Sub
    msg = ReadSetup()
    If Len(msg) > 0 Then Say "Fill these in on the SETUP sheet first:" & vbLf & msg, vbExclamation: Exit Sub
    If Not LoadTemplate() Then
        Say "The MF-12 template on SETUP isn't there, or isn't the right file:" & vbLf & mTemplatePath & vbLf & _
            "It's mf12_template.pdf, which comes with this macro.", vbExclamation
        Exit Sub
    End If

    mStage = "reading Corp License Number"
    LoadContacts

    mStage = "picking the period"
    mPeriod = PickPeriod(stepNo)
    If Len(mPeriod) = 0 Then Exit Sub

    mStage = "reading Refunds Due"
    If Not LoadRows() Then Say "No rows on Refunds Due for period " & mPeriod & ".", vbExclamation: Exit Sub
    SetMonth

    mStage = "matching rows to suppliers"
    AssignPackets

    mMonthDir = MonthFolder()
    Application.ScreenUpdating = False
    mStage = "writing the Corp names onto Refunds Due"
    WriteCorpNames
    If stepNo = 1 Then
        mStage = "making folders under " & mMonthDir
        MakeFolders
        mStage = "filing the INBOX"
        SortInbox
    ElseIf Not FolderExists(mMonthDir) Then
        Say "No folder for " & mPeriod & " yet (" & mMonthDir & "). Run Refunds1_BuildPackets first.", vbExclamation
        Exit Sub
    Else
        CountInbox
    End If

    mStage = "checking the packet folders"
    CheckFiles
    If stepNo = 1 Then
        mStage = "filling the MF-12s"
        WriteForms
    End If
    mStage = "deciding who is ready"
    DecideReady
    If stepNo = 2 Then
        mStage = "making the Outlook drafts"
        MakeDrafts
    End If

    mStage = "writing the ENTRY sheet"
    WriteEntry
    mStage = "writing the CHECK sheet"
    WriteCheck stepNo
    Application.ScreenUpdating = True

    For i = 1 To mNSup
        If mSup(i).NRows > 0 Then
            If Len(mSup(i).Held) = 0 Then ready = ready + 1 Else held = held + 1
            If Left$(mSup(i).Status, 7) = "DRAFTED" Then drafted = drafted + 1
        End If
    Next i
    If stepNo = 1 Then
        msg = "Packets built for " & mPeriod & "." & vbLf & ready & " supplier(s) ready for drafts, " & held & " waiting."
        If mNLook > 0 Then msg = msg & vbLf & mNLook & " thing(s) on the CHECK sheet need a look."
        If ready > 0 Then msg = msg & vbLf & "When you're happy with CHECK, run Refunds2_MakeDrafts."
    Else
        msg = drafted & " draft(s) made for " & mPeriod & ". " & held & " supplier(s) held back." & vbLf & _
              "They're in Outlook Drafts. Nothing was sent."
    End If
    Say msg, vbInformation
    Exit Sub

Failed:
    Close
    Application.ScreenUpdating = True
    Say "Stopped with an error while " & mStage & ":" & vbLf & Err.Description, vbCritical
End Sub

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

' Adds the SETUP sheet and the two new columns the first time. Returns what to tell you.
Private Function EnsureSetup() As String
    Dim ws As Worksheet, cs As Worksheet, hr As Long, lastCol As Long, msg As String

    Set cs = FindContactsSheet(mBook)
    If cs Is Nothing Then
        EnsureSetup = "Can't find the Corp License Number sheet (the one with License Number and Email address columns)."
        Exit Function
    End If
    hr = HeaderRow(cs, Array(H_LICENSE, H_EMAIL))
    lastCol = cs.Cells(hr, cs.Columns.Count).End(xlToLeft).Column
    If FindCol(cs, hr, H_SHORT, lastCol) = 0 Then
        lastCol = lastCol + 1
        cs.Cells(hr, lastCol).Value = "Short Name"
        LookLike cs.Cells(hr, lastCol - 1), cs.Cells(hr, lastCol)
        msg = msg & "Added a Short Name column to " & cs.Name & ": the name that goes in folder and file names, like MANSFIELD." & vbLf
    End If
    If FindCol(cs, hr, H_NET, lastCol) = 0 Then
        lastCol = lastCol + 1
        cs.Cells(hr, lastCol).Value = "NET"
        LookLike cs.Cells(hr, lastCol - 1), cs.Cells(hr, lastCol)
        msg = msg & "Added a NET column to " & cs.Name & ": put Y for the suppliers who want the allowance taken off." & vbLf
    End If

    Set ws = SheetNamed(SETUP_SHEET)
    If ws Is Nothing Then
        MakeSetupSheet
        msg = msg & "Added a SETUP sheet. Fill in its yellow cells." & vbLf
    End If
    If Len(msg) > 0 Then EnsureSetup = msg & vbLf & "Then run Refunds1_BuildPackets again."
End Function

Private Sub MakeSetupSheet()
    Dim ws As Worksheet, r As Long
    Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
    ws.Name = SETUP_SHEET
    ws.Columns(1).ColumnWidth = 44
    ws.Columns(2).ColumnWidth = 60
    ws.Columns(3).ColumnWidth = 40
    ws.Range("B8").NumberFormat = "@"     ' a ZIP like 02134 keeps its zero
    ws.Range("A1").Value = "GA refunds setup"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Fill the yellow cells once. The macro reads them by the words in column A, so keep those."
    SetupLine ws, 4, "Purchaser name (as it goes on the MF-12)", ""
    SetupLine ws, 5, "Purchaser street address", ""
    SetupLine ws, 6, "Purchaser city", ""
    SetupLine ws, 7, "Purchaser state (2 letters)", ""
    SetupLine ws, 8, "Purchaser ZIP", ""
    SetupLine ws, 9, "Company short name for the email subject", ""
    SetupLine ws, 11, "Month folder", ""
    ws.Range("C11").Value = "Full path. {YYYY}, {MM} and {MONTH} get filled in. Example: \\server\share\Georgia\Refunds\{YYYY}-{MM} {MONTH}"
    SetupLine ws, 12, "MF-12 template file", ""
    ws.Range("C12").Value = "Full path to mf12_template.pdf, which comes with this macro."
    SetupLine ws, 14, "Email subject", ">MONTH< >YEAR< GA MF-12 Waiver - >US< / >SUPPLIER<"
    SetupLine ws, 15, "Email text", "Attached is our Georgia MF-12 waiver covering >TYPES< for >MONTH< >YEAR<, " & _
        "along with copies of the supporting invoices." & vbLf & _
        "Please sign, notarize, and return each attached MF-12 at your earliest convenience." & vbLf & vbLf & _
        "Thank you very much."
    ws.Range("B15").WrapText = True
    ws.Rows(15).AutoFit
    ws.Range("C15").Value = "Alt+Enter starts a new line inside the cell."
    ws.Range("C14").Value = "These get filled in per supplier: >MONTH< >YEAR< >SUPPLIER< >US< >TYPES< >CONTACT<"
    r = 17
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
    ws.Range("A1:C30").VerticalAlignment = xlTop
    ws.Rows(15).AutoFit
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

' Returns a list of what's missing, blank when all good.
Private Function ReadSetup() As String
    Dim ws As Worksheet, miss As String, r As Long, t As String

    Set ws = SheetNamed(SETUP_SHEET)
    mBuyer = SetupValue(ws, "purchaser name")
    mAddress = SetupValue(ws, "purchaser street")
    mCity = SetupValue(ws, "purchaser city")
    mState = UCase$(SetupValue(ws, "purchaser state"))
    mZip = SetupValue(ws, "purchaser zip")
    If IsAllDigits(mZip) And Len(mZip) < 5 Then mZip = Right$("0000" & mZip, 5)
    mUs = SetupValue(ws, "company short name")
    mPattern = SetupValue(ws, "month folder")
    mTemplatePath = SetupValue(ws, "mf-12 template")
    mSubject = SetupValue(ws, "email subject")
    mBody = SetupValue(ws, "email text")
    If Right$(mPattern, 1) = "\" Then mPattern = Left$(mPattern, Len(mPattern) - 1)

    If Len(mBuyer) = 0 Then miss = miss & "  Purchaser name" & vbLf
    If Len(mAddress) = 0 Then miss = miss & "  Purchaser street address" & vbLf
    If Len(mCity) = 0 Then miss = miss & "  Purchaser city" & vbLf
    If Len(mState) <> 2 Then miss = miss & "  Purchaser state, 2 letters" & vbLf
    If Not (mZip Like "#####" Or mZip Like "#####-####") Then miss = miss & "  Purchaser ZIP, like 12345 or 12345-6789" & vbLf
    If Len(mUs) = 0 Then miss = miss & "  Company short name for the email subject" & vbLf
    If InStr(mPattern, "\") = 0 Or InStr(mPattern, "{MM}") + InStr(mPattern, "{MONTH}") = 0 Or InStr(mPattern, "{YYYY}") = 0 Then
        miss = miss & "  Month folder, a full path with {YYYY} and {MM} or {MONTH} in it" & vbLf
    End If
    If Len(mTemplatePath) = 0 Then miss = miss & "  MF-12 template file" & vbLf
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
        If ws.Name <> SETUP_SHEET And ws.Name <> CHECK_SHEET And ws.Name <> ENTRY_SHEET And ws.Name <> HOWTO_SHEET Then
            If HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE, H_PERIOD)) > 0 Then Set FindRefundsSheet = ws: Exit Function
        End If
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
    Dim lic As String, i As Long, j As Long, n As String

    Set ws = FindContactsSheet(mBook)
    hr = HeaderRow(ws, Array(H_LICENSE, H_EMAIL))
    lastCol = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    cLic = FindCol(ws, hr, H_LICENSE, lastCol)
    cCorp = FindCol(ws, hr, H_CORP, lastCol)
    cContact = FindCol(ws, hr, H_CONTACT, lastCol)
    cEmail = FindCol(ws, hr, H_EMAIL, lastCol)
    cGroup = FindCol(ws, hr, H_GROUP, lastCol)
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
                    If cShort > 0 Then .ShortName = UCase$(SafeName(CellStr(ws.Cells(r, cShort).Value)))
                    If cNet > 0 Then
                        n = LCase$(CellStr(ws.Cells(r, cNet).Value))
                        .IsNet = (n = "y" Or n = "yes" Or n = "x" Or n = "net" Or n = "true")
                    End If
                    If Len(.ShortName) = 0 Then .Problem = "No Short Name on " & ws.Name & " row " & r
                    If Len(.Corp) = 0 Then .Problem = "No Corp name on " & ws.Name & " row " & r & " (it goes on the MF-12 as the seller)"
                End With
            End If
        End If
    Next r

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

Private Function PickPeriod(ByVal stepNo As Long) As String
    Dim ws As Worksheet, hr As Long, c As Long, r As Long, lastRow As Long, guess As String, chk As Worksheet

    If mQuiet Then PickPeriod = mQuietPeriod: Exit Function
    Set chk = SheetNamed(CHECK_SHEET)
    If stepNo = 2 And Not chk Is Nothing Then guess = CellStr(chk.Range("B1").Value)
    If Len(guess) = 0 Then
        Set ws = FindRefundsSheet(mBook)
        hr = HeaderRow(ws, Array(H_BOL, H_INV, H_LICENSE))
        c = FindCol(ws, hr, H_PERIOD, LastHeaderCol(ws, hr))
        If c > 0 Then
            lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            For r = lastRow To hr + 1 Step -1
                guess = PeriodText(ws.Cells(r, c).Value)
                If Len(guess) > 0 Then Exit For
            Next r
        End If
    End If
    PickPeriod = Trim$(InputBox(IIf(stepNo = 1, "Build packets", "Make drafts") & " for which period?" & vbLf & _
        "(the way the Period column shows it)", "GA refunds", guess))
End Function

Private Function LoadRows() As Boolean
    Dim ws As Worksheet, hr As Long, lastCol As Long, lastRow As Long, r As Long, data As Variant, ok As Boolean
    Dim cType As Long, cPer As Long, cCorp As Long, cLic As Long, cBol As Long, cInv As Long, cInvDate As Long
    Dim cGal As Long, cRef As Long, cAllow As Long, cTtl As Long, cFuel As Long, want As String

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
    If cType = 0 Or cPer = 0 Or cCorp = 0 Or cLic = 0 Or cBol = 0 Or cInv = 0 Or cInvDate = 0 Or cGal = 0 Or _
       cRef = 0 Or cAllow = 0 Or cTtl = 0 Or cFuel = 0 Then
        Err.Raise vbObjectError + 10, , "Refunds Due is missing a column. It needs: Refund Type (or Schedule), Period, Corp, " & _
            "License Number, BOL, INV, Inv Date, Gallons, Refund, Allowance, TTL Refund, Fuel."
    End If
    lastRow = ws.Cells(ws.Rows.Count, cBol).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, cPer).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, cPer).End(xlUp).Row
    If lastRow <= hr Then Exit Function
    data = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lastRow, lastCol)).Value
    want = Squash(mPeriod)

    For r = 1 To UBound(data, 1)
        If Squash(PeriodText(data(r, cPer))) = want Then
            mNRow = mNRow + 1
            ReDim Preserve mRow(1 To mNRow)
            With mRow(mNRow)
                .SheetRow = hr + r
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
            End With
        End If
    Next r
    LoadRows = (mNRow > 0)
End Function

' ======================================================================================
'  Rows to suppliers to packets
' ======================================================================================

Private Sub AssignPackets()
    Dim i As Long, t As Long, s As Long, p As Long, key As String, probs As Object, k As Variant

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

            ' a packet needs a type and a supplier with a short name
            If .Sup > 0 And t > 0 Then
                If Len(mSup(.Sup).ShortName) > 0 Then
                    key = .Prefix & "-" & mSup(.Sup).ShortName
                    p = FindPkt(key)
                    If p = 0 Then
                        mNPkt = mNPkt + 1
                        ReDim Preserve mPkt(1 To mNPkt)
                        p = mNPkt
                        mPkt(p).PName = key
                        mPkt(p).Prefix = .Prefix
                        mPkt(p).Sup = .Sup
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
                If .Sup > 0 Then
                    If Len(mSup(.Sup).Held) = 0 Then mSup(.Sup).Held = "a row needs a look (row " & .SheetRow & ")"
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
            AddLook mSup(s).Problem, "License " & mSup(s).License & ", " & mSup(s).NRows & " row(s) this period", "Fill it in on " & CONTACTS_SHEET & "."
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

' The Corp name as typed on Corp License Number goes onto every matched row of Refunds Due
' for this period, so the report's own spelling of a supplier never shows anywhere you work
' from. A row whose license isn't on Corp License Number keeps what was pasted.
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
'  Folders and the INBOX
' ======================================================================================

Private Sub SetMonth()
    Dim i As Long, counts(1 To 12) As Long, best As Long
    If ParsePeriod(mPeriod, mMonth, mYear) Then Exit Sub
    ' period text it can't read: go by the invoice dates
    For i = 1 To mNRow
        If mRow(i).InvDate > 0 Then counts(Month(mRow(i).InvDate)) = counts(Month(mRow(i).InvDate)) + 1
    Next i
    For i = 1 To 12
        If counts(i) > 0 Then
            If best = 0 Then
                best = i
            ElseIf counts(i) > counts(best) Then
                best = i
            End If
        End If
    Next i
    If best = 0 Then Err.Raise vbObjectError + 11, , "Can't tell the month from period '" & mPeriod & "'."
    mMonth = best
    For i = 1 To mNRow
        If mRow(i).InvDate > 0 Then
            If Month(mRow(i).InvDate) = best Then mYear = Year(mRow(i).InvDate): Exit For
        End If
    Next i
    AddLook "Couldn't read the period '" & mPeriod & "'", "Went by the invoice dates: " & MonthName(mMonth) & " " & mYear, _
        "Nothing to do if that's right."
End Sub

Private Function MonthFolder() As String
    Dim p As String
    p = Replace(mPattern, "{YYYY}", CStr(mYear))
    p = Replace(p, "{MM}", Format$(mMonth, "00"))
    p = Replace(p, "{MONTH}", UCase$(MonthName(mMonth)))
    MonthFolder = p
End Function

Private Sub MakeFolders()
    Dim p As Long
    MkDirs mMonthDir
    MkDirs mMonthDir & "\INBOX"
    MkDirs mMonthDir & "\MFD-33"
    For p = 1 To mNPkt
        mPkt(p).Folder = mMonthDir & "\" & mPkt(p).PName
        MkDirs mPkt(p).Folder
        MkDirs mPkt(p).Folder & "\UPLOAD"
    Next p
End Sub

' Files each "<BOL> VI" / "<BOL> CI" into its packet folder under its full name.
Private Sub SortInbox()
    Dim files As Variant, i As Long, f As String, stem As String, ext As String, bolPart As String, tag As String, sq As String
    Dim r As Long, pk As Long, twoPkts As Boolean, waiting As Long, invs As String, bolShown As String, target As String

    files = ListFiles(mMonthDir & "\INBOX")
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
                    AddLook "INBOX: no " & mPeriod & " row has BOL " & bolPart, f & " stays in INBOX.", _
                        "Check the BOL. It may belong to another month."
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
                        Name mMonthDir & "\INBOX\" & f As target
                        mInboxLeft = mInboxLeft - 1
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub CountInbox()
    Dim files As Variant, i As Long
    files = ListFiles(mMonthDir & "\INBOX")
    If IsEmpty(files) Then Exit Sub
    For i = 1 To UBound(files)
        If Not IsJunkFile(CStr(files(i))) Then mInboxLeft = mInboxLeft + 1
    Next i
    If mInboxLeft > 0 Then AddLook "INBOX still has " & mInboxLeft & " file(s)", "INBOX in the month folder", _
        "Run Refunds1_BuildPackets to file them."
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

    For p = 1 To mNPkt
        mPkt(p).Folder = mMonthDir & "\" & mPkt(p).PName
        files = ListFiles(mPkt(p).Folder)
        mPkt(p).VIMissing = ""
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
            AddLook mPkt(p).PName & ": no VI yet for BOL " & mPkt(p).VIMissing, "The " & mPkt(p).PName & " folder", _
                "Pull the vendor invoice, save it in INBOX as '<BOL> VI', run Refunds1_BuildPackets."
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
                "The number of forms changed, or it's an old copy. Move it out of the folder and run Refunds1_BuildPackets."
        ElseIf Len(bad) > 0 Then
            mPkt(p).FormState = "MF-12 OUT OF DATE"
            AddLook mPkt(p).PName & ": the MF-12 doesn't match the rows anymore", bad, _
                "The report changed since it was made (or it was edited by hand). Move it out of the folder and run Refunds1_BuildPackets to make a fresh one."
        ElseIf onDisk = mPkt(p).NForms Then
            mPkt(p).FormState = "KEPT"
        ElseIf onDisk = 0 Then
            mPkt(p).FormState = "MISSING"
        Else
            mPkt(p).FormState = "SOME FORMS MISSING"
            AddLook mPkt(p).PName & ": only " & onDisk & " of " & mPkt(p).NForms & " MF-12 forms are there", "The " & mPkt(p).PName & " folder", _
                "Move the MF-12s out of the folder and run Refunds1_BuildPackets to make them all fresh."
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
    FormFileName = mPkt(p).PName & " " & UCase$(MonthName(mMonth)) & " " & mYear & " MF-12" & _
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
'  Who's ready, and the drafts
' ======================================================================================

Private Sub DecideReady()
    Dim s As Long, p As Long, why As String
    For s = 1 To mNSup
        If mSup(s).NRows > 0 Then
            why = mSup(s).Held
            If Len(why) = 0 And Len(mSup(s).Problem) > 0 Then why = mSup(s).Problem
            If Len(why) = 0 And Len(mSup(s).Email) = 0 And Len(mSup(s).GroupMail) = 0 Then
                why = "no email address on " & CONTACTS_SHEET
                AddLook mSup(s).ShortName & ": no email address", CONTACTS_SHEET & " row " & mSup(s).SheetRow, "Fill in Email address or Group Email Address."
            End If
            For p = 1 To mNPkt
                If mPkt(p).Sup = s And Len(why) = 0 Then
                    If Len(mPkt(p).Problem) > 0 Then
                        why = mPkt(p).PName & ": " & mPkt(p).Problem
                    ElseIf Len(mPkt(p).VIMissing) > 0 Then
                        why = mPkt(p).PName & ": waiting on VI for BOL " & mPkt(p).VIMissing
                    ElseIf mPkt(p).FormState <> "MADE" And mPkt(p).FormState <> "KEPT" Then
                        why = mPkt(p).PName & ": MF-12 " & LCase$(mPkt(p).FormState)
                    End If
                End If
            Next p
            mSup(s).Held = why
            If Len(why) = 0 Then
                mSup(s).Status = "READY"
                ListAttachments s
                If mSup(s).MB > BIG_EMAIL_MB Then mSup(s).Status = "READY, but " & Format$(mSup(s).MB, "0.0") & " MB may be too big for one email"
            Else
                mSup(s).Status = "HELD: " & why
            End If
        End If
    Next s
End Sub

' Everything in the supplier's packet folders except the CI copies, notarized copies and the UPLOAD folder.
' MF-12s first.
Private Sub ListAttachments(ByVal s As Long)
    Dim p As Long, files As Variant, i As Long, pass As Long, f As String, stem As String, ext As String, isForm As Boolean
    mSup(s).Files = "": mSup(s).FileCount = 0: mSup(s).MB = 0
    For pass = 1 To 2
        For p = 1 To mNPkt
            If mPkt(p).Sup = s Then
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
                                mSup(s).Files = mSup(s).Files & mPkt(p).Folder & "\" & f & vbLf
                                mSup(s).FileCount = mSup(s).FileCount + 1
                                mSup(s).MB = mSup(s).MB + FileLen(mPkt(p).Folder & "\" & f) / 1048576#
                            End If
                        End If
                    Next i
                End If
            End If
        Next p
    Next pass
End Sub

Private Sub MakeDrafts()
    Dim s As Long, ol As Object, ns As Object, found As Object, subj As String, body As String

    If Len(mTestDraftDir) = 0 Then
        Set ol = OutlookApp()
        If ol Is Nothing Then Err.Raise vbObjectError + 14, , "Couldn't open Outlook."
        Set ns = ol.GetNamespace("MAPI")
    End If
    For s = 1 To mNSup
        If mSup(s).NRows > 0 And Len(mSup(s).Held) = 0 Then
            mStage = "making the draft for " & mSup(s).ShortName
            subj = FillWords(mSubject, s)
            body = FillWords(mBody, s)
            If Len(mTestDraftDir) > 0 Then
                If FileExists(mTestDraftDir & "\" & mSup(s).ShortName & " draft.txt") Then
                    mSup(s).Status = "ALREADY IN DRAFTS"
                Else
                    TestDraft s, subj, body
                    mSup(s).Status = "DRAFTED " & Format$(Now, "m/d h:nn AM/PM")
                End If
            Else
                Set found = FindMail(ns.GetDefaultFolder(5), subj)        ' Sent Items
                If Not found Is Nothing Then
                    mSup(s).Status = "ALREADY SENT " & Format$(found.SentOn, "m/d/yyyy")
                Else
                    Set found = FindMail(ns.GetDefaultFolder(16), subj)   ' Drafts
                    If Not found Is Nothing Then
                        mSup(s).Status = "ALREADY IN DRAFTS"
                    Else
                        OutlookDraft ol, s, subj, body
                        mSup(s).Status = "DRAFTED " & Format$(Now, "m/d h:nn AM/PM")
                    End If
                End If
            End If
        End If
    Next s
End Sub

Private Sub OutlookDraft(ol As Object, ByVal s As Long, ByVal subj As String, ByVal body As String)
    Dim m As Object, insp As Object, f As Variant, html As String, pos As Long
    Set m = ol.CreateItem(0)
    If Len(mSup(s).Email) > 0 Then
        m.To = mSup(s).Email
        If Len(mSup(s).GroupMail) > 0 Then m.CC = mSup(s).GroupMail
    Else
        m.To = mSup(s).GroupMail
    End If
    m.Subject = subj
    For Each f In Split(mSup(s).Files, vbLf)
        If Len(f) > 0 Then m.Attachments.Add CStr(f)
    Next f
    ' opening the item's window (without showing it) puts your default signature in the body
    On Error Resume Next
    Set insp = m.GetInspector
    html = m.HTMLBody
    On Error GoTo 0
    pos = InStr(1, LCase$(html), "<body")
    If pos > 0 Then pos = InStr(pos, html, ">")
    If pos > 0 Then
        m.HTMLBody = Left$(html, pos) & BodyHtml(body) & Mid$(html, pos + 1)
    Else
        m.Body = body & vbCrLf & vbCrLf & m.Body
    End If
    m.Save
    m.Close 0
End Sub

Private Sub TestDraft(ByVal s As Long, ByVal subj As String, ByVal body As String)
    Dim f As Integer
    f = FreeFile
    Open mTestDraftDir & "\" & mSup(s).ShortName & " draft.txt" For Output As #f
    If Len(mSup(s).Email) > 0 Then
        Print #f, "To: " & mSup(s).Email
        Print #f, "CC: " & mSup(s).GroupMail
    Else
        Print #f, "To: " & mSup(s).GroupMail
        Print #f, "CC: "
    End If
    Print #f, "Subject: " & subj
    Print #f, "Body: " & body
    Print #f, "Attachments:"
    Print #f, mSup(s).Files;
    Close #f
End Sub

Private Function FillWords(ByVal t As String, ByVal s As Long) As String
    t = Replace(t, ">MONTH<", MonthName(mMonth))
    t = Replace(t, ">YEAR<", CStr(mYear))
    t = Replace(t, ">SUPPLIER<", mSup(s).ShortName)
    t = Replace(t, ">US<", mUs)
    t = Replace(t, ">CONTACT<", mSup(s).Contact)
    t = Replace(t, ">TYPES<", TypeWords(s))
    FillWords = t
End Function

' "licensed customer sales and government customer sales", in the SETUP list's order.
Private Function TypeWords(ByVal s As Long) As String
    Dim t As Long, p As Long, words() As String, n As Long, i As Long, has As Boolean
    For t = 1 To mNType
        has = False
        For p = 1 To mNPkt
            If mPkt(p).Sup = s And mPkt(p).Prefix = mTypePrefix(t) Then has = True
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
    Dim ws As Worksheet, r As Long, i As Long, s As Long, p As Long, t As Long, first As Long, ready As Long, held As Long
    Dim n As Long, g As Double, rf As Double, al As Double, tt As Double

    Set ws = SheetNamed(CHECK_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = CHECK_SHEET
    End If
    ws.Hyperlinks.Delete
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "Period"
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = mPeriod
    ws.Range("C1").NumberFormat = "@"
    ws.Range("C1").Value = MonthName(mMonth) & " " & mYear
    ws.Range("E1").Value = IIf(stepNo = 1, "Packets built ", "Drafts made ") & Format$(Now, "m/d/yyyy h:nn AM/PM")
    ws.Range("A1:C1").Font.Bold = True
    For s = 1 To mNSup
        If mSup(s).NRows > 0 Then
            If Len(mSup(s).Held) = 0 Then ready = ready + 1 Else held = held + 1
        End If
    Next s
    With ws.Range("A2:K2")
        .Merge
        .Font.Bold = True
        .Font.Size = 12
        If mNLook = 0 Then
            .Value = "All clear. " & ready & " supplier(s) ready" & IIf(stepNo = 1, " for drafts.", ".")
            .Interior.Color = GREEN
        Else
            .Value = "LOOK: " & mNLook & " thing(s) below need a look. " & ready & " supplier(s) ready, " & held & " held back."
            .Interior.Color = RED
        End If
    End With
    ws.Range("A3").Value = "Month folder"
    ws.Hyperlinks.Add Anchor:=ws.Range("B3"), Address:=mMonthDir, TextToDisplay:=mMonthDir
    r = 5

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

    r = TableHead(ws, r, Array("Supplier", "Seller on the MF-12", "License", "Types", "Rows", "Amount on MF-12s", "NET", "To", "CC", "Attachments", "Status"))
    For s = 1 To mNSup
        If mSup(s).NRows > 0 Then
            ws.Cells(r, 1).Value = mSup(s).ShortName
            ws.Cells(r, 2).Value = mSup(s).Corp
            ws.Cells(r, 3).NumberFormat = "@"
            ws.Cells(r, 3).Value = mSup(s).License
            ws.Cells(r, 4).Value = TypeWords(s)
            ws.Cells(r, 5).Value = mSup(s).NRows
            g = 0
            For p = 1 To mNPkt
                If mPkt(p).Sup = s Then g = g + mPkt(p).Amount
            Next p
            ws.Cells(r, 6).Value = Round(g, 2)
            ws.Cells(r, 6).NumberFormat = "#,##0.00"
            ws.Cells(r, 7).Value = IIf(mSup(s).IsNet, "NET", "gross")
            If Len(mSup(s).Email) > 0 Then
                ws.Cells(r, 8).Value = mSup(s).Email
                ws.Cells(r, 9).Value = mSup(s).GroupMail
            Else
                ws.Cells(r, 8).Value = mSup(s).GroupMail
            End If
            If mSup(s).FileCount > 0 Then
                ws.Cells(r, 10).Value = AttachSummary(s)
            End If
            ws.Cells(r, 11).Value = mSup(s).Status
            If Len(mSup(s).Held) = 0 Then ws.Cells(r, 11).Interior.Color = GREEN Else ws.Cells(r, 11).Interior.Color = RED
            r = r + 1
        End If
    Next s
    ' rows whose supplier isn't on Corp License Number
    For i = 1 To mNRow
        If mRow(i).Sup = 0 Then
            ws.Cells(r, 1).Value = "(no match)"
            ws.Cells(r, 2).Value = mRow(i).Corp
            ws.Cells(r, 3).NumberFormat = "@"
            ws.Cells(r, 3).Value = mRow(i).License
            ws.Cells(r, 11).Value = "HELD: row " & mRow(i).SheetRow & ", " & mRow(i).Problem
            ws.Cells(r, 11).Interior.Color = RED
            r = r + 1
        End If
    Next i
    r = r + 1

    first = r
    r = TableHead(ws, r, Array("Packet", "Rows", "Gallons", "Refund", "Allowance", "TTL Refund", "On the MF-12", "MF-12 forms", "VIs", "Folder"))
    For p = 1 To mNPkt
        ws.Cells(r, 1).Value = mPkt(p).PName
        ws.Cells(r, 2).Value = mPkt(p).NRows
        ws.Cells(r, 3).Value = Round(mPkt(p).Gallons, 3)
        ws.Cells(r, 4).Value = Round(mPkt(p).Refund, 2)
        ws.Cells(r, 5).Value = Round(mPkt(p).Allow, 2)
        ws.Cells(r, 6).Value = Round(mPkt(p).Ttl, 2)
        ws.Cells(r, 7).Value = Round(mPkt(p).Amount, 2)
        ws.Cells(r, 8).Value = mPkt(p).NForms & " " & IIf(Len(mPkt(p).Problem) > 0, "HELD", mPkt(p).FormState)
        If Len(mPkt(p).VIMissing) = 0 Then ws.Cells(r, 9).Value = "all there" Else ws.Cells(r, 9).Value = "missing " & mPkt(p).VIMissing
        If FolderExists(mPkt(p).Folder) Then ws.Hyperlinks.Add Anchor:=ws.Cells(r, 10), Address:=mPkt(p).Folder, TextToDisplay:=mPkt(p).PName
        r = r + 1
    Next p
    r = r + 1

    ' totals by refund type, gross, the way the MFD-33 wants them
    r = TableHead(ws, r, Array("Totals by refund type (gross, for the MFD-33)", "Rows", "Gallons", "Refund", "Allowance", "TTL Refund"))
    For t = 1 To mNType
        n = 0: g = 0: rf = 0: al = 0: tt = 0
        For i = 1 To mNRow
            If Squash(mRow(i).TypeText) = mTypeText(t) Then
                n = n + 1: g = g + mRow(i).Gallons: rf = rf + mRow(i).Refund: al = al + mRow(i).Allow: tt = tt + mRow(i).Ttl
            End If
        Next i
        If n > 0 Then
            ws.Cells(r, 1).Value = mTypePrefix(t) & ": " & mTypePhrase(t)
            ws.Cells(r, 2).Value = n: ws.Cells(r, 3).Value = Round(g, 3): ws.Cells(r, 4).Value = Round(rf, 2)
            ws.Cells(r, 5).Value = Round(al, 2): ws.Cells(r, 6).Value = Round(tt, 2)
            r = r + 1
        End If
    Next t

    ws.Range(ws.Cells(first, 3), ws.Cells(r, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 4), ws.Cells(r, 7)).NumberFormat = "#,##0.00"
    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 34
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 30
    ws.Columns("E:G").ColumnWidth = 14
    ws.Columns("H:I").ColumnWidth = 26
    ws.Columns(10).ColumnWidth = 34
    ws.Columns(11).ColumnWidth = 50
    ws.Range("A5:K" & r).VerticalAlignment = xlTop
    ws.Range("A5:K" & r).WrapText = True
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
Private Function AttachSummary(ByVal s As Long) As String
    Dim f As Variant, nm As String, stem As String, ext As String, forms As Long, vis As Long, other As Long
    For Each f In Split(mSup(s).Files, vbLf)
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
    AttachSummary = mSup(s).FileCount & " files, " & Format$(mSup(s).MB, "0.0") & " MB: " & forms & " MF-12, " & vis & " VI" & _
        IIf(other > 0, ", " & other & " other", "")
End Function

' ======================================================================================
'  The HOW TO tab: two buttons and the steps, at the front of the workbook
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

    ws.Range("B1").Value = "GA refunds: MF-12 packets, Outlook drafts, portal entry"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 18
    HowToButton ws, "btnBuild", "1.  Build packets", "Refunds1_BuildPackets", 0
    HowToButton ws, "btnDrafts", "2.  Make drafts", "Refunds2_MakeDrafts", 1

    r = 4
    r = HowToLine(ws, r, "Each month", True)
    r = HowToLine(ws, r, "1. Paste the refund report into Refunds Due like always.", False)
    r = HowToLine(ws, r, "2. Click Build packets. It makes the month's folders and an INBOX.", False)
    r = HowToLine(ws, r, "3. Save the invoices into INBOX as ""<BOL> VI"" and ""<BOL> CI"" (""12345 VI"", ""12345 CI"").", False)
    r = HowToLine(ws, r, "4. Click Build packets again. It renames and files everything and fills the MF-12s. " & _
        "CHECK shows what is ready and what needs a look.", False)
    r = HowToLine(ws, r, "5. When a supplier says READY on CHECK, click Make drafts. The drafts land in Outlook Drafts. " & _
        "Nothing is sent; look each one over and send it yourself.", False)
    r = HowToLine(ws, r, "6. When a notarized MF-12 comes back, save it in its packet folder beside the one you sent, " & _
        "with NOTARIZED at the end of the name. The macro leaves those alone. At filing time, type NOTARIZED " & _
        "in the month folder's search box to list them all.", False)
    r = HowToLine(ws, r, "7. Portal entry: the ENTRY tab is the typing sheet. One group per refund type (licensed, " & _
        "government, diversions), which is one MFD-33 filing; inside it one block per MF-12, the columns in the " & _
        "portal's order, rows top to bottom. Exempt Portion is not on the sheet; leave it blank in the portal. " & _
        "When a block is in, the portal's running total should match the total under it, and the MFD-33 total " & _
        "closes the group. It redraws on every click, so click Build packets first if the report changed.", False)
    r = r + 1
    r = HowToLine(ws, r, "The tabs", True)
    r = HowToLine(ws, r, "SETUP holds the share folder, the MF-12 template, the refund types and the email words. " & _
        "Yellow cells are yours to fill.", False)
    r = HowToLine(ws, r, "CHECK is the state of the month: each supplier, READY or what it is waiting on, and anything " & _
        "the macro could not place.", False)
    r = HowToLine(ws, r, "ENTRY is the typing sheet for the Georgia portal.", False)
    r = HowToLine(ws, r, "Corp License Number is the contacts list. The Corp name there prints on every MF-12 exactly " & _
        "as typed, and every click copies it onto the matching rows of Refunds Due, so trim any extra text " & _
        "off it once, like "" - ST (COMPANY)"". Short Name and NET are the two columns the macro added.", False)
    r = r + 1
    r = HowToLine(ws, r, "Rules it keeps", True)
    r = HowToLine(ws, r, "It never sends an email, never deletes a file, and never writes over one. A packet with " & _
        "anything on CHECK gets no draft until it is fixed.", False)
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
    b.Left = ws.Range("B2").Left + 4 + slot * 300
    b.Top = ws.Range("B2").Top + 4
    b.Width = 280
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

' The portal takes each invoice typed by hand, one field at a time, no paste. So: one block
' per MF-12 packet, the columns in the portal's own order, the rows in the order they sit on
' the form, and a total under each block to check the portal's running total against.
' Values only. Redrawn every run.
Private Sub WriteEntry()
    Dim ws As Worksheet, r As Long, p As Long, i As Long, k As Long, ix As Variant, order As Variant, s As Long, t As Long
    Dim g As Double, a As Double, fg As Double, fa As Double, formNo As Long, slot As Long, head As String
    Dim curT As Long, bannerRow As Long, tg As Double, ta As Double, tRows As Long, tForms As Long, tHeld As Long

    Set ws = SheetNamed(ENTRY_SHEET)
    If ws Is Nothing Then
        Set ws = mBook.Worksheets.Add(After:=mBook.Worksheets(mBook.Worksheets.Count))
        ws.Name = ENTRY_SHEET
    End If
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = "Portal entry for " & MonthName(mMonth) & " " & mYear & " (" & mPeriod & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Value = "One group per refund type, which is one MFD-33 filing; inside it, one block per MF-12. " & _
        "Type each row into the Georgia portal left to right, top to bottom. " & _
        "Exempt Portion in the portal stays blank, so it is not on this sheet. " & _
        "Tax Paid to Seller is the gross refund on every row, even for a NET supplier; NET is only on its MF-12. " & _
        "When a block is in, the portal's running total should match the total under it; the MFD-33 total closes each group. " & _
        "Drawn " & Format$(Now, "m/d/yyyy h:nn AM/PM") & "."
    With ws.Range("A2:G2")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(2).RowHeight = 45
    r = 4

    If mNPkt = 0 Then
        ws.Cells(r, 1).Value = "No packets this period."
    End If
    order = PacketOrder()
    curT = 0
    For i = 1 To mNPkt
        p = order(i)
        s = mPkt(p).Sup
        t = TypeIndex(mPkt(p).Prefix)
        ' A new refund type: close the last group with its MFD-33 total, open this one with a banner.
        ' The banner's counts are filled in once the group is done (see EntryGroupClose).
        If t <> curT Then
            If curT > 0 Then r = EntryGroupClose(ws, r, curT, bannerRow, tg, ta, tRows, tForms, tHeld)
            curT = t
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
    Next i
    If curT > 0 Then r = EntryGroupClose(ws, r, curT, bannerRow, tg, ta, tRows, tForms, tHeld)

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

' Closes a refund-type group on ENTRY: the MFD-33 total line, two blank rows, and the banner
' above the group filled in with what it holds. Hands back the next row.
Private Function EntryGroupClose(ws As Worksheet, ByVal r As Long, ByVal t As Long, ByVal bannerRow As Long, _
    ByVal tg As Double, ByVal ta As Double, ByVal tRows As Long, ByVal tForms As Long, ByVal tHeld As Long) As Long
    Dim txt As String
    ws.Cells(r, 1).Value = "MFD-33 total, " & mTypePhrase(t) & ", " & tRows & " row" & IIf(tRows = 1, "", "s") & _
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
    txt = UCase$(mTypePhrase(t)) & ": one MFD-33 filing, " & tForms & " MF-12" & IIf(tForms = 1, "", "s") & _
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

' Packets in the order to key them: by refund type in the SETUP list's order (one type is one
' MFD-33 filing, so all its MF-12s sit together), then by supplier short name inside the type.
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
    ta = TypeIndex(mPkt(a).Prefix): tb = TypeIndex(mPkt(b).Prefix)
    If ta <> tb Then
        PktBefore = (ta < tb)
    Else
        na = mSup(mPkt(a).Sup).ShortName: nb = mSup(mPkt(b).Sup).ShortName
        PktBefore = (na < nb)
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

Private Function FindPkt(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To mNPkt
        If mPkt(i).PName = nm Then FindPkt = i: Exit Function
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
