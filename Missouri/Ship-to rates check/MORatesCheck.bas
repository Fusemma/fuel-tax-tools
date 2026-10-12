Option Compare Binary
Option Explicit

' MO Rates Check
'
' Each quarter the Missouri DOR posts a boundary file (every address range in
' the state and the jurisdiction code it sits in) and a total rate file (each
' code's rates). This database keeps the last two quarters and shows where the
' use tax rate changed between them.
'
' Setup, once: in a new blank Access database press Alt+F11, then Insert >
' Module. Press Ctrl+A, paste this whole page over what's there, press Ctrl+S
' and name it MORatesCode. Then click anywhere inside "Sub SetUp" just below
' and press F5. After that, everything runs from the MO Rates Check form.
'
' The state's files are big, so each quarter goes in a file of its own in a
' folder on this PC (Settings, DataFolder), not in this database. A new
' quarter is a new file and the oldest one is deleted, so this database
' stays small and never needs compacting.

Public gQuiet As Boolean       ' True while a test runs: no message boxes
Public gLastMessage As String  ' the last message, shown or not

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal fmt As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, _
    ByVal src As LongPtr, ByVal bytes As LongPtr)

Private Const APP_NAME As String = "MO Rates Check"
Private Const FORM_MAIN As String = "MO Rates Check"
Private Const RPT_CHANGES As String = "Use tax changes"
Private Const FORM_SHIPTOS As String = "Ship-tos to fix"
Private Const FORM_MAP As String = "Look up on the map"
Private Const MAP_URL As String = "https://missouri.atr.avalara.com/"
Private Const DATA_DEFAULT As String = "%LOCALAPPDATA%\MO Rates Check"
Private Const BOUNDARY_SPEC As String = "Kind:T1;Lo:L;Hi:L;LoTxt:T12;HiTxt:T12;Side:T1;Street:T80;" & _
    "NameKey:T80;PreKey:T4;SufKey:T8;PostKey:T4;FullKey:T100;City:T40;CityKey:T40;Zip:T5;" & _
    "ZipLo:T9;ZipHi:T9;County:T3;JCode:T13"
Private Const SHIPTOS_SPEC As String = "AcctNo:T40;ShipToNo:T40;CustomerName:T120;ShipToName:T120;" & _
    "Address1:T150;Address2:T150;City:T60;State:T20;County:T60;Tax:D;TaxExemption:T60;ProductTypes:T60;" & _
    "CustomerStatus:T20;ShipToStatus:T20;FirstSeen:W;LastSeen:W;PinnedCode:T13;PinnedOn:W;" & _
    "Result:T10;Why:T255;CodeNow:T13;HowNow:T10;WhereNow:T10;LineNow:T10;UseNow:D;CodeLast:T13;UseLast:D;" & _
    "Exempt:T3;CheckedOn:W"
Private Const CHECKS_SPEC As String = "Quarter:T8;CheckedOn:W;AcctNo:T40;ShipToNo:T40;Code:T13;How:T10;" & _
    "UseRate:D;Tax:D;Result:T10;Why:T255"
Private Const RATECHANGES_SPEC As String = "JCode:T13;County:T3;CountyName:T40;OldUse:D;NewUse:D;" & _
    "OldRate:T20;NewRate:T20;Cities:T255;CitiesAll:M;AddrRows:L;ZipRows:L"
Private Const MOVES_SPEC As String = "County:T3;CountyName:T40;City:T40;Zip:T5;Street:T80;Side:T1;" & _
    "LoTxt:T12;HiTxt:T12;LoNum:L;Houses:T80;OldCodes:T120;NewCodes:T120;OldUse:D;NewUse:D;" & _
    "OldRate:T20;NewRate:T20"
Private Const MOVEGROUPS_SPEC As String = "County:T3;CountyName:T40;City:T40;Zip:T5;Street:T80;LoNum:L;" & _
    "Houses:M;OldCodes:T120;NewCodes:T120;OldUse:D;NewUse:D;OldRate:T20;NewRate:T20"

Private mSyn As Collection     ' word -> standard spelling (AVENUE -> AVE)
Private mSuffix As Collection  ' street suffixes (ST, AVE, RD ...)
Private mUnit As Collection    ' words that start a unit (STE, APT, LOT ...)
Private mDir As Collection     ' N S E W NE NW SE SW
Private mKeys As Collection    ' street pieces already worked out, during a load
Private mCodes() As String     ' Effective's list, kept between calls
Private mCityKeys As Collection
Private mCounty As Collection  ' county code -> name
Private mCountyKeys As Collection ' county name, standardized -> code

' One street's rows from one quarter, while comparing.
Private Type StreetRows
    n As Long
    num() As Boolean
    lo() As Long
    hi() As Long
    loTxt() As String
    hiTxt() As String
    side() As String
    code() As String
    street() As String
    city() As String
End Type

' The rows that could be one ship-to's street, while matching.
Private Type MatchRows
    n As Long
    num() As Boolean
    lo() As Long
    hi() As Long
    loTxt() As String
    hiTxt() As String
    side() As String
    pre() As String
    nm() As String
    suf() As String
    post() As String
    full() As String
    code() As String
    street() As String
    city() As String
End Type

' A Boundary table on one of its indexes, its fields at hand (looking a field
' up by name on every row is slow).
Private Type Cursor
    rs As DAO.Recordset
    key As DAO.Field
    zip As DAO.Field
    full As DAO.Field
    nm As DAO.Field
    pre As DAO.Field
    suf As DAO.Field
    post As DAO.Field
    lo As DAO.Field
    hi As DAO.Field
    loTxt As DAO.Field
    hiTxt As DAO.Field
    side As DAO.Field
    code As DAO.Field
    street As DAO.Field
    city As DAO.Field
End Type

'=============================================================== setup

Public Sub SetUp()
    Dim db As DAO.Database, old As Boolean
    On Error GoTo Fail
    Set db = CurrentDb
    Application.SetOption "Track Name AutoCorrect Info", False
    Application.SetOption "Perform Name AutoCorrect", False
    Application.SetOption "Auto Compact", False
    old = DropOldTables(db)
    MakeTables db
    MakeCounties db
    MakeShipTosQuery db
    PointChanges
    MakeChangesReport
    MakeShipTosForm
    MakeMapForm
    MakeMainForm
    SetStartForm db
    If old Then
        Say "Set up is done. The state's data has moved out of this file, so shrink it once: File, Info, " & _
            "Compact & Repair Database. Then load the two quarters again.", vbInformation
    Else
        Say "Set up is done. Next: click Load new quarter on the MO Rates Check form.", vbInformation
    End If
    If Not gQuiet Then DoCmd.OpenForm FORM_MAIN
    Exit Sub
Fail:
    Say "Set up stopped: " & Err.Description & " (error " & Err.Number & ")", vbExclamation
End Sub

' An earlier version kept the state's data in this file; it lives in the data
' folder now. True if there was any to drop.
Private Function DropOldTables(db As DAO.Database) As Boolean
    Dim t As Variant
    For Each t In Array("Boundary", "BoundaryLast", "BoundaryNew", "Rates", "RateChanges", "Moves", "MoveGroups")
        If TableExists(db, CStr(t)) Then
            DoCmd.Close acReport, RPT_CHANGES, acSaveNo
            DropTable db, CStr(t)
            DropOldTables = True
        End If
    Next
    If DropOldTables Then
        db.Execute "DELETE FROM Settings WHERE SettingName IN ('NowQuarter','LastQuarter','NowFile','LastFile'," & _
            "'RateChangeCount','RateChangeUnused','MoveLineCount','MoveStreetCount','MoveZipCount','ComparedOn')", _
            dbFailOnError
    End If
End Function

Private Sub MakeTables(db As DAO.Database)
    If Not TableExists(db, "Settings") Then
        NewTable db, "Settings", "SettingName:T50;SettingValue:T255"
        db.Execute "CREATE UNIQUE INDEX PrimaryKey ON Settings (SettingName) WITH PRIMARY", dbFailOnError
    End If
    If GetOpt("DataFolder") = "" Then SetOpt "DataFolder", DATA_DEFAULT
    If Not TableExists(db, "Counties") Then
        NewTable db, "Counties", "County:T3;CountyName:T40"
        db.Execute "CREATE UNIQUE INDEX PrimaryKey ON Counties (County) WITH PRIMARY", dbFailOnError
    End If
    If Not TableExists(db, "ShipTos") Then
        NewTable db, "ShipTos", SHIPTOS_SPEC
        db.Execute "CREATE UNIQUE INDEX PrimaryKey ON ShipTos (AcctNo, ShipToNo) WITH PRIMARY", dbFailOnError
    End If
    If Not TableExists(db, "Checks") Then
        NewTable db, "Checks", CHECKS_SPEC
        db.Execute "CREATE INDEX idxQuarter ON Checks (Quarter)", dbFailOnError
    End If
    db.TableDefs.Refresh
End Sub

' The county codes as the DOR writes them in jurisdiction codes. Names are the
' Census Bureau's. The DOR uses 193 for Ste. Genevieve County (the Census
' moved it to 186); both are here.
Private Sub MakeCounties(db As DAO.Database)
    Dim s As String, p As Variant, nm As String
    s = "001 Adair|003 Andrew|005 Atchison|007 Audrain|009 Barry|011 Barton|013 Bates|015 Benton|"
    s = s & "017 Bollinger|019 Boone|021 Buchanan|023 Butler|025 Caldwell|027 Callaway|029 Camden|"
    s = s & "031 Cape Girardeau|033 Carroll|035 Carter|037 Cass|039 Cedar|041 Chariton|043 Christian|"
    s = s & "045 Clark|047 Clay|049 Clinton|051 Cole|053 Cooper|055 Crawford|057 Dade|059 Dallas|"
    s = s & "061 Daviess|063 DeKalb|065 Dent|067 Douglas|069 Dunklin|071 Franklin|073 Gasconade|"
    s = s & "075 Gentry|077 Greene|079 Grundy|081 Harrison|083 Henry|085 Hickory|087 Holt|089 Howard|"
    s = s & "091 Howell|093 Iron|095 Jackson|097 Jasper|099 Jefferson|101 Johnson|103 Knox|105 Laclede|"
    s = s & "107 Lafayette|109 Lawrence|111 Lewis|113 Lincoln|115 Linn|117 Livingston|119 McDonald|"
    s = s & "121 Macon|123 Madison|125 Maries|127 Marion|129 Mercer|131 Miller|133 Mississippi|"
    s = s & "135 Moniteau|137 Monroe|139 Montgomery|141 Morgan|143 New Madrid|145 Newton|147 Nodaway|"
    s = s & "149 Oregon|151 Osage|153 Ozark|155 Pemiscot|157 Perry|159 Pettis|161 Phelps|163 Pike|"
    s = s & "165 Platte|167 Polk|169 Pulaski|171 Putnam|173 Ralls|175 Randolph|177 Ray|179 Reynolds|"
    s = s & "181 Ripley|183 St. Charles|185 St. Clair|186 Ste. Genevieve|187 St. Francois|189 St. Louis|"
    s = s & "193 Ste. Genevieve|195 Saline|197 Schuyler|199 Scotland|201 Scott|203 Shannon|205 Shelby|"
    s = s & "207 Stoddard|209 Stone|211 Sullivan|213 Taney|215 Texas|217 Vernon|219 Warren|"
    s = s & "221 Washington|223 Wayne|225 Webster|227 Worth|229 Wright"
    db.Execute "DELETE FROM Counties", dbFailOnError
    For Each p In Split(s, "|")
        nm = Mid$(p, 5) & " County"
        db.Execute "INSERT INTO Counties (County, CountyName) VALUES ('" & Left$(p, 3) & "','" & nm & "')", dbFailOnError
    Next
    db.Execute "INSERT INTO Counties (County, CountyName) VALUES ('510','St. Louis City')", dbFailOnError
End Sub

' The Use tax changes report reads this quarter's file in the data folder.
' The query is rewritten only when that file's name changes.
Private Sub PointChanges()
    Dim db As DAO.Database, sql As String, src As String, qd As DAO.QueryDef, have As Boolean
    Set db = CurrentDb
    src = "[;DATABASE=" & QuarterFile(NowQuarter()) & "]."
    For Each qd In db.QueryDefs
        If qd.Name = "qChanges" Then
            have = True
            If InStr(1, qd.SQL, src, vbTextCompare) > 0 Then Exit Sub
        End If
    Next
    sql = "SELECT CountyName, 1 AS Part, '' AS City, '' AS Zip, 'All of ' & JCode AS Place, " & _
        "IIf(Nz(Cities,'')='','(in the file by zip+4 only)',Cities) AS Area, '' AS Houses, 0 AS LoNum, " & _
        "OldUse, NewUse, OldRate, NewRate, '' AS Codes FROM " & src & "RateChanges " & _
        "WHERE Nz(AddrRows,0)+Nz(ZipRows,0)>0 " & _
        "UNION ALL SELECT CountyName, 2, City, Zip, Street, City & ' ' & Zip, Houses, LoNum, OldUse, NewUse, " & _
        "OldRate, NewRate, OldCodes & ' to ' & NewCodes FROM " & src & "MoveGroups"
    DoCmd.Close acReport, RPT_CHANGES, acSaveNo
    If have Then db.QueryDefs.Delete "qChanges"
    db.CreateQueryDef "qChanges", sql
End Sub

Private Sub SetStartForm(db As DAO.Database)
    On Error Resume Next
    db.Properties("StartUpForm") = FORM_MAIN
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("StartUpForm", dbText, FORM_MAIN)
    End If
End Sub

'---------------------------------------------------------------- the form

' The main form: a few lines on what's loaded, the ship-to counts in big
' numbers, what to click next, the buttons, each with a line on what it
' does, and the whole how-to at the right, so the notes never come along.
Private Sub MakeMainForm()
    Dim frm As Form, nm As String, c As Control, i As Long, kinds As Variant, colors As Variant
    DeleteObject acForm, FORM_MAIN
    Set frm = CreateForm()
    nm = frm.Name
    frm.Caption = APP_NAME
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 3            ' only show when the window is too small
    frm.AutoCenter = True
    frm.Width = 20100
    frm.Section(acDetail).Height = 10700
    frm.Section(acDetail).BackColor = RGB(255, 255, 255)

    Set c = CreateControl(nm, acLabel, acDetail, , , 360, 200, 12960, 800)
    c.Caption = APP_NAME
    Look c, 30, True

    StatusBox nm, "=LoadedLine()", 1060
    StatusBox nm, "=ChangesLine()", 1480
    StatusBox nm, "=ListLine()", 1900

    ' the ship-to counts
    kinds = Array("Fix", "Map", "OK", "Skipped")
    colors = Array(RGB(192, 0, 0), RGB(0, 70, 160), RGB(0, 120, 0), RGB(110, 110, 110))
    For i = 0 To 3
        Set c = CreateControl(nm, acTextBox, acDetail, , , 360 + i * 3300, 2420, 3000, 760)
        c.ControlSource = "=BigNumber(""" & kinds(i) & """)"
        Look c, 32, True
        c.ForeColor = colors(i)
        Fixed c
        Set c = CreateControl(nm, acTextBox, acDetail, , , 360 + i * 3300, 3180, 3000, 380)
        c.ControlSource = "=BigLabel(""" & kinds(i) & """)"
        Look c, 13, False
        c.ForeColor = RGB(90, 90, 90)
        Fixed c
    Next

    Set c = CreateControl(nm, acTextBox, acDetail, , , 360, 3760, 13080, 560)
    c.ControlSource = "=NextLine()"
    c.Name = "Next step"
    Look c, 16, True
    c.BackStyle = 1
    c.BackColor = RGB(255, 242, 204)
    Fixed c

    Heading nm, "Each quarter", 4600
    AddButton nm, "Load new quarter", "=LoadQuarter()", 360, 5000, _
        "Pick the boundary .csv. Its rate .xlsx in the same folder loads with it."
    AddButton nm, "Use tax changes", "=ShowChanges()", 4720, 5000, _
        "Shows the places and streets where the use tax rate changed."

    Heading nm, "Ship-tos", 7000
    AddButton nm, "Load ship-to list", "=LoadShipToList()", 360, 7400, _
        "Pick the ship-to list (.xlsx or .csv). Every ship-to gets checked by itself."
    AddButton nm, "Ship-tos to fix", "=ShowShipTos()", 4720, 7400, _
        "The ship-tos whose rate in the system isn't the state's. One scrolling list."
    AddButton nm, "Look up on the map", "=ShowMapList()", 9080, 7400, _
        "The ship-tos it couldn't place. Type the code from the state's map."

    HowToPanel nm, 13560, 200, 6440, 10400

    DoCmd.Close acForm, nm, acSaveYes
    DoCmd.Rename FORM_MAIN, acForm, nm
End Sub

' The how-to beside the buttons: where to get the state's files, which two,
' and every step after. The text is HowToText.
Private Sub HowToPanel(ByVal frmName As String, ByVal x As Long, ByVal y As Long, _
        ByVal w As Long, ByVal h As Long)
    Dim c As Control
    Set c = CreateControl(frmName, acRectangle, acDetail, , , x, y, w, h)
    c.SpecialEffect = 0
    c.BackStyle = 1
    c.BackColor = RGB(242, 242, 242)
    c.BorderStyle = 0
    Set c = CreateControl(frmName, acLabel, acDetail, , , x + 140, y + 120, w - 280, 440)
    c.Caption = "How to use it"
    Look c, 16, True
    Set c = CreateControl(frmName, acTextBox, acDetail, , , x + 140, y + 620, w - 280, h - 740)
    c.ControlSource = "=HowToText()"
    c.TextFormat = 1              ' rich text, for the bold
    Look c, 11, False
    c.ForeColor = RGB(40, 40, 40)
    Fixed c
End Sub

Private Sub StatusBox(ByVal frmName As String, ByVal source As String, ByVal y As Long)
    Dim c As Control
    Set c = CreateControl(frmName, acTextBox, acDetail, , , 360, y, 13080, 400)
    c.ControlSource = source
    Look c, 15, False
    Fixed c
End Sub

Private Sub Heading(ByVal frmName As String, ByVal caption As String, ByVal y As Long)
    Dim c As Control
    Set c = CreateControl(frmName, acLabel, acDetail, , , 360, y, 6000, 380)
    c.Caption = caption
    Look c, 14, True
    c.ForeColor = RGB(90, 90, 90)
End Sub

' A text box that only shows something: no cursor, no tab stop.
Private Sub Fixed(c As Control)
    c.Locked = True
    c.TabStop = False
    c.Enabled = False
End Sub

Private Sub AddButton(ByVal frmName As String, ByVal caption As String, ByVal onClick As String, _
        ByVal x As Long, ByVal y As Long, ByVal tip As String)
    Dim c As Control
    Set c = CreateControl(frmName, acCommandButton, acDetail, , , x, y, 4100, 900)
    c.Caption = caption
    c.OnClick = onClick
    c.FontSize = 15
    c.FontBold = True
    Set c = CreateControl(frmName, acLabel, acDetail, , , x, y + 960, 4100, 760)
    c.Caption = tip
    Look c, 11, False
    c.ForeColor = RGB(70, 70, 70)
End Sub

Private Sub Look(c As Control, ByVal size As Integer, ByVal bold As Boolean)
    c.FontName = "Arial"
    c.FontSize = size
    c.FontBold = bold
    c.ForeColor = RGB(0, 0, 0)
    On Error Resume Next
    c.BorderStyle = 0
    c.BackStyle = 0
End Sub

'-------------------------------------------------------------- the report

Private Sub MakeChangesReport()
    Dim rpt As Report, nm As String, c As Control, fc As Object
    DeleteObject acReport, RPT_CHANGES
    Set rpt = CreateReport()
    nm = rpt.Name
    rpt.RecordSource = "qChanges"
    rpt.Caption = RPT_CHANGES
    On Error Resume Next
    rpt.Printer.Orientation = 2
    rpt.Printer.LeftMargin = 720
    rpt.Printer.RightMargin = 720
    rpt.Printer.TopMargin = 720
    rpt.Printer.BottomMargin = 720
    On Error GoTo 0
    rpt.Width = 14400

    CreateGroupLevel nm, "=1", True, False          ' the title block
    CreateGroupLevel nm, "CountyName", True, False
    CreateGroupLevel nm, "Part", True, False
    CreateGroupLevel nm, "City", False, False
    CreateGroupLevel nm, "Zip", False, False
    CreateGroupLevel nm, "Place", False, False
    CreateGroupLevel nm, "LoNum", False, False
    rpt.GroupLevel(1).KeepTogether = 2
    rpt.GroupLevel(2).KeepTogether = 2

    ' title block
    rpt.Section(5).Height = 2000
    Set c = Box(nm, 5, "=ChangesTitle()", 0, 0, 14400, 720, 28, True)
    Set c = Box(nm, 5, "=ChangesSummary()", 0, 800, 14400, 1100, 16, False)
    c.CanGrow = True
    rpt.Section(5).CanGrow = True

    ' county band
    rpt.Section(7).Height = 760
    rpt.Section(7).BackColor = RGB(214, 228, 240)
    Set c = Box(nm, 7, "=UCase([CountyName])", 120, 100, 7200, 560, 22, True)
    Set c = Box(nm, 7, "=CountyLine(Sum(IIf([Part]=1,1,0)),Sum(IIf([Part]=2,1,0)))", 7400, 180, 6880, 460, 14, False)
    c.TextAlign = 3

    ' part heading and column heads
    rpt.Section(9).Height = 1000
    Set c = Box(nm, 9, "=IIf([Part]=1,""Places with a new use rate (every address in them)""," & _
        """Streets where houses moved to a different use rate"")", 0, 120, 14400, 440, 16, True)
    c.FontItalic = True
    HeadLabel nm, "Where", 0, 4180, 1
    HeadLabel nm, "City on the mail", 4220, 2300, 1
    HeadLabel nm, "House numbers", 6560, 2900, 1
    HeadLabel nm, "Old rate", 9480, 1160, 3
    HeadLabel nm, "New rate", 10660, 1240, 3
    HeadLabel nm, "Change", 11920, 1080, 3
    HeadLabel nm, "Codes", 13080, 1320, 1

    ' one line per place or street
    rpt.Section(0).Height = 400
    rpt.Section(0).CanGrow = True
    On Error Resume Next
    rpt.Section(0).AlternateBackColor = RGB(242, 242, 242)
    On Error GoTo 0
    Set c = Box(nm, 0, "Place", 0, 20, 4180, 370, 16, False)
    c.CanGrow = True
    Set c = Box(nm, 0, "Area", 4220, 40, 2300, 340, 11, False)
    c.CanGrow = True
    Set c = Box(nm, 0, "Houses", 6560, 20, 2900, 370, 16, False)
    c.CanGrow = True
    Set c = Box(nm, 0, "OldRate", 9480, 20, 1160, 370, 16, False)
    c.TextAlign = 3
    Set c = Box(nm, 0, "NewRate", 10660, 20, 1240, 370, 16, True)
    c.TextAlign = 3
    Set c = Box(nm, 0, "=ChangeText([OldUse],[NewUse])", 11920, 40, 1080, 340, 13, True)
    c.TextAlign = 3
    Set fc = c.FormatConditions.Add(acExpression, , "[NewUse]>[OldUse]")
    fc.ForeColor = RGB(192, 0, 0)
    Set fc = c.FormatConditions.Add(acExpression, , "[NewUse]<[OldUse]")
    fc.ForeColor = RGB(0, 120, 0)
    Set c = Box(nm, 0, "Codes", 13080, 20, 1320, 370, 8, False)
    c.ForeColor = RGB(80, 80, 80)
    c.CanGrow = True

    ' page footer
    rpt.Section(3).Height = 0
    rpt.Section(4).Height = 400
    Set c = Box(nm, 4, "=ChangesTitle() & ""     Page "" & [Page] & "" of "" & [Pages]", 0, 60, 14400, 300, 10, False)
    c.ForeColor = RGB(90, 90, 90)

    DoCmd.Close acReport, nm, acSaveYes
    DoCmd.Rename RPT_CHANGES, acReport, nm
End Sub

' A report text box showing a field or an expression.
Private Function Box(ByVal rptName As String, ByVal sect As Integer, ByVal source As String, ByVal x As Long, _
        ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal size As Integer, ByVal bold As Boolean) As Control
    Dim c As Control
    Set c = CreateReportControl(rptName, acTextBox, sect, , , x, y, w, h)
    c.ControlSource = source
    Look c, size, bold
    Set Box = c
End Function

Private Sub HeadLabel(ByVal rptName As String, ByVal caption As String, ByVal x As Long, ByVal w As Long, _
        ByVal align As Integer)
    Dim c As Control
    Set c = CreateReportControl(rptName, acLabel, 9, , , x, 600, w, 340)
    c.Caption = caption
    Look c, 12, True
    c.ForeColor = RGB(60, 60, 60)
    c.TextAlign = align
End Sub

'------------------------------------------------- step two's report and form

Private Sub MakeShipTosQuery(db As DAO.Database)
    Dim sql As String
    sql = "SELECT IIf(Result='Fix',1,2) AS Part, IIf(Exempt='Yes',1,0) AS IsExempt, County, City, ShipToName, " & _
        "AcctNo, ShipToNo, IIf(LineNow='Address 1',Address1,IIf(LineNow='Address 2' Or LineNow='both',Address2," & _
        "Trim(Nz(Address2,'') & '  ' & Nz(Address1,'')))) AS Addr, Tax, UseNow, CodeNow, Why, HowNow " & _
        "FROM ShipTos WHERE Result='Fix' OR Result='Map'"
    On Error Resume Next
    db.QueryDefs.Delete "qShipTos"
    On Error GoTo 0
    db.CreateQueryDef "qShipTos", sql
End Sub

' Ship-tos to fix: one scrolling list, the ones to fix first, the ones to look
' up on the map last. One click on a 100879/4 copies it.
Private Sub MakeShipTosForm()
    Dim frm As Form, nm As String, c As Control
    DeleteObject acReport, FORM_SHIPTOS
    DeleteObject acForm, FORM_SHIPTOS
    Set frm = CreateForm()
    nm = frm.Name
    frm.Caption = FORM_SHIPTOS
    frm.RecordSource = "SELECT * FROM qShipTos ORDER BY Part, IsExempt, County, City, ShipToName"
    ListLook frm
    frm.AllowEdits = False

    Set c = CreateControl(nm, acTextBox, acHeader, , , 240, 120, 12200, 600)
    c.ControlSource = "=ShipTosTitle()"
    Look c, 22, True
    Fixed c
    Set c = CreateControl(nm, acTextBox, acHeader, , , 240, 760, 12200, 700)
    c.ControlSource = "=ShipTosHelp()"
    Look c, 13, False
    Fixed c
    CopiedLabel nm
    MapHead nm, "Ship-to (click to copy)", 240, 3300
    MapHead nm, "Address, city, county", 3660, 3700
    MapHead nm, "Rate now", 7440, 1300, 3
    MapHead nm, "State's rate", 8800, 1360, 3
    MapHead nm, "What to do", 10360, 3700
    MapHead nm, "Code", 14200, 1800

    ' only the 100879/4 box takes a click; the rest just shows
    Set c = MapBox(nm, "ShipToName", 240, 40, 3300, 360, 14, True)
    Fixed c
    CopyBox nm
    Set c = MapBox(nm, "=[Addr] & Chr(13) & Chr(10) & [City] & IIf(Nz([County],'')<>'',', ' & [County],'')", _
        3660, 40, 3700, 640, 13, False)
    Fixed c
    Set c = MapBox(nm, "=RateText([Tax])", 7440, 40, 1300, 380, 15, False)
    c.TextAlign = 3
    Fixed c
    Set c = MapBox(nm, "=IIf([Part]=1,RateText([UseNow]),"""")", 8800, 40, 1360, 380, 15, True)
    c.TextAlign = 3
    Fixed c
    Set c = MapBox(nm, "=WhatToDo([Part],[Why],[IsExempt])", 10360, 40, 3700, 640, 12, False)
    Fixed c
    Set c = MapBox(nm, "CodeNow", 14200, 60, 1800, 300, 10, False)
    c.ForeColor = RGB(80, 80, 80)
    Fixed c

    DoCmd.Close acForm, nm, acSaveYes
    DoCmd.Rename FORM_SHIPTOS, acForm, nm
End Sub

' The look the two lists share: one row per ship-to, a blue band on top.
Private Sub ListLook(frm As Form)
    frm.DefaultView = 1
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = True
    frm.Width = 16300
    DoCmd.RunCommand acCmdFormHdrFtr
    frm.Section(acHeader).Height = 2200
    frm.Section(acHeader).BackColor = RGB(214, 228, 240)
    frm.Section(acDetail).Height = 760
    frm.Section(acDetail).BackColor = RGB(255, 255, 255)
    frm.Section(acFooter).Height = 0
End Sub

' The 100879/4 box (account 100879, ship-to 4): one click copies it.
Private Sub CopyBox(ByVal frmName As String)
    Dim c As Control
    Set c = CreateControl(frmName, acTextBox, acDetail, , , 240, 420, 2200, 300)
    c.ControlSource = "=[AcctNo] & ""/"" & [ShipToNo]"
    c.Name = "Acct and ship-to"
    Look c, 12, True
    c.ForeColor = RGB(0, 70, 160)
    c.FontUnderline = True
    c.Locked = True
    c.TabStop = False
    c.OnClick = "=CopyShipTo([AcctNo],[ShipToNo])"
End Sub

' Says what the last click copied.
Private Sub CopiedLabel(ByVal frmName As String)
    Dim c As Control
    Set c = CreateControl(frmName, acLabel, acHeader, , , 12700, 1300, 3400, 360)
    c.Name = "Copied"
    c.Caption = " "
    Look c, 13, True
    c.ForeColor = RGB(0, 120, 0)
    c.TextAlign = 3
End Sub

' The ship-tos that couldn't be placed for sure, one per line, with a box
' for the code from the state's map.
Private Sub MakeMapForm()
    Dim frm As Form, nm As String, c As Control
    DeleteObject acForm, FORM_MAP
    Set frm = CreateForm()
    nm = frm.Name
    frm.Caption = FORM_MAP
    frm.RecordSource = "SELECT * FROM ShipTos WHERE Result='Map' ORDER BY County, City, ShipToName"
    ListLook frm

    Set c = CreateControl(nm, acLabel, acHeader, , , 240, 120, 9000, 600)
    c.Caption = "Look these up on the state's map"
    Look c, 22, True
    Set c = CreateControl(nm, acLabel, acHeader, , , 240, 760, 10200, 700)
    c.Caption = "For each one, find the address on the state's map and type the code it shows into the " & _
        "white box at the right, like 37000-051-000. Then click Check again. Click a 100879/4 to copy it."
    Look c, 13, False
    Set c = CreateControl(nm, acLabel, acHeader, , , 240, 1460, 4000, 320)
    c.Caption = "Open the state's map"
    c.HyperlinkAddress = MAP_URL
    Look c, 13, True
    c.ForeColor = RGB(0, 70, 160)
    c.FontUnderline = True
    Set c = CreateControl(nm, acCommandButton, acHeader, , , 12700, 200, 3300, 900)
    c.Caption = "Check again"
    c.OnClick = "=CheckAgain()"
    c.FontSize = 16
    c.FontBold = True
    CopiedLabel nm
    MapHead nm, "Ship-to (click to copy)", 240, 3000
    MapHead nm, "Address 2 / Address 1", 3300, 3700
    MapHead nm, "City, county", 7060, 2240
    MapHead nm, "Why it needs a look", 9360, 4300
    MapHead nm, "Code from the map", 13760, 2300

    Set c = MapBox(nm, "ShipToName", 240, 40, 3000, 360, 14, True)
    CopyBox nm
    Set c = MapBox(nm, "=Trim(Nz([Address2],'') & Chr(13) & Chr(10) & Nz([Address1],''))", 3300, 40, 3700, 640, 13, False)
    Set c = MapBox(nm, "=[City] & Chr(13) & Chr(10) & [County]", 7060, 40, 2240, 640, 13, False)
    Set c = MapBox(nm, "Why", 9360, 40, 4300, 640, 11, False)
    Set c = CreateControl(nm, acTextBox, acDetail, , , 13760, 120, 2200, 480)
    c.ControlSource = "PinnedCode"
    c.Name = "Code from the map"
    c.FontName = "Arial"
    c.FontSize = 16
    c.FontBold = True
    c.BackColor = RGB(255, 255, 255)
    c.BorderColor = RGB(0, 70, 160)
    c.ValidationRule = "Is Null Or Like ""#####-###-###"""
    c.ValidationText = "Type the code the way the map shows it, like 37000-051-000."

    DoCmd.Close acForm, nm, acSaveYes
    DoCmd.Rename FORM_MAP, acForm, nm
End Sub

Private Sub MapHead(ByVal frmName As String, ByVal caption As String, ByVal x As Long, ByVal w As Long, _
        Optional ByVal align As Integer = 1)
    Dim c As Control
    Set c = CreateControl(frmName, acLabel, acHeader, , , x, 1860, w, 300)
    c.Caption = caption
    Look c, 11, True
    c.ForeColor = RGB(60, 60, 60)
    c.TextAlign = align
End Sub

Private Function MapBox(ByVal frmName As String, ByVal source As String, ByVal x As Long, ByVal y As Long, _
        ByVal w As Long, ByVal h As Long, ByVal size As Integer, ByVal bold As Boolean) As Control
    Dim c As Control
    Set c = CreateControl(frmName, acTextBox, acDetail, , , x, y, w, h)
    c.ControlSource = source
    Look c, size, bold
    c.Locked = True
    c.TabStop = False
    Set MapBox = c
End Function

'------------------------------------------------- step two's buttons

Public Function LoadShipToList()
    Dim p As String, msg As String
    On Error GoTo Fail
    p = PickFile("Pick the ship-to list (.xlsx or .csv)", "Ship-to list", "*.xlsx; *.xls; *.csv", GetOpt("ListFolder"))
    If p = "" Then Exit Function
    SetOpt "ListFolder", Left$(p, InStrRev(p, "\"))
    DoCmd.Hourglass True
    If LoadShipToFile(p) Then
        msg = gLastMessage
        If NowQuarter() <> "" Then
            CheckShipTos
            msg = msg & vbCrLf & vbCrLf & ShipToLine()
        Else
            msg = msg & vbCrLf & vbCrLf & "Load a quarter's files too; the ship-tos get checked then."
        End If
        TidyUp
        RefreshLists
        Say msg, vbInformation
        If Val(GetOpt("ShipToFix")) + Val(GetOpt("ShipToMap")) > 0 Then ShowShipTos
    End If
    TidyUp
    Exit Function
Fail:
    TidyUp
    Say "Loading the list stopped: " & Err.Description & " (error " & Err.Number & ")", vbExclamation
End Function

' The map list's button, after typing codes.
Public Function CheckAgain()
    On Error GoTo Fail
    If CurrentProject.AllForms(FORM_MAP).IsLoaded Then
        If Forms(FORM_MAP).Dirty Then Forms(FORM_MAP).Dirty = False
    End If
    If NowQuarter() = "" Then Say "Load a quarter's files first.", vbInformation: Exit Function
    If DCount("*", "ShipTos") = 0 Then Say "Load the ship-to list first.", vbInformation: Exit Function
    DoCmd.Hourglass True
    CheckShipTos
    TidyUp
    RefreshLists
    Say ShipToLine(), vbInformation
    Exit Function
Fail:
    TidyUp
    Say "Checking stopped: " & Err.Description & " (error " & Err.Number & ")", vbExclamation
End Function

Public Function ShowShipTos()
    On Error GoTo Fail
    If DCount("*", "ShipTos") = 0 Then Say "Load the ship-to list first.", vbInformation: Exit Function
    If DCount("*", "ShipTos", "Result='Fix' OR Result='Map'") = 0 Then
        Say "Nothing to fix right now.", vbInformation
        Exit Function
    End If
    If CurrentProject.AllForms(FORM_SHIPTOS).IsLoaded Then Forms(FORM_SHIPTOS).Requery
    DoCmd.OpenForm FORM_SHIPTOS
    Exit Function
Fail:
    Say "Couldn't open the list: " & Err.Description, vbExclamation
End Function

Public Function ShowMapList()
    On Error GoTo Fail
    If DCount("*", "ShipTos", "Result='Map'") = 0 Then
        Say "Nothing to look up on the map right now.", vbInformation
        Exit Function
    End If
    DoCmd.OpenForm FORM_MAP
    Exit Function
Fail:
    Say "Couldn't open the list: " & Err.Description, vbExclamation
End Function

' One click on a 100879/4 box puts that text on the clipboard.
Public Function CopyShipTo(ByVal acct As Variant, ByVal shipTo As Variant) As Boolean
    Dim s As String
    s = Nz(acct, "") & "/" & Nz(shipTo, "")
    CopyShipTo = ToClipboard(s)
    On Error Resume Next
    With Screen.ActiveControl
        .SelStart = 0
        .SelLength = Len(s)
    End With
    If CopyShipTo Then
        Screen.ActiveForm.Controls("Copied").Caption = "Copied " & s
    Else
        Screen.ActiveForm.Controls("Copied").Caption = "Didn't copy. Click it again."
    End If
End Function

Public Function ShipTosTitle() As String
    ShipTosTitle = "Ship-tos to fix, " & QLabel(GetOpt("ShipToCheckedFor")) & " (rates from " & _
        DateText(GetOpt("ShipToCheckedFor")) & ")"
End Function

Public Function ShipTosHelp() As String
    Dim nFix As Long, nMap As Long, s As String
    nFix = Val(GetOpt("ShipToFix"))
    nMap = Val(GetOpt("ShipToMap"))
    s = Plural(nFix, "to fix in the system", "to fix in the system")
    If nMap > 0 Then s = s & ", then " & Plural(nMap, "to look up on the map", "to look up on the map") & " at the bottom"
    ShipTosHelp = s & ". Click a 100879/4 to copy it, then paste it into the system."
End Function

' What the list says to do for one ship-to.
Public Function WhatToDo(ByVal part As Variant, ByVal why As Variant, ByVal exempt As Variant) As String
    Dim s As String
    s = Nz(why, "")
    If Nz(part, 1) = 2 Then s = "Look it up on the map: " & s
    If Nz(exempt, 0) = 1 Then s = s & " (exempt, checked anyway)"
    WhatToDo = s
End Function

Public Function ShipToLine() As String
    Dim s As String
    On Error Resume Next
    If GetOpt("ShipToCheckedFor") = "" Then
        ShipToLine = Plural(DCount("*", "ShipTos"), "ship-to", "ship-tos") & " in the list, not checked yet. " & _
            "They get checked once a quarter's files are loaded."
        Exit Function
    End If
    s = "Ship-tos, checked for " & QLabel(GetOpt("ShipToCheckedFor")) & ": " & _
        Plural(Val(GetOpt("ShipToCount")), "ship-to", "ship-tos") & " in the list; " & _
        Plural(Val(GetOpt("ShipToFix")), "to fix", "to fix") & ", " & _
        Plural(Val(GetOpt("ShipToMap")), "to look up on the map", "to look up on the map") & ", " & _
        Plural(Val(GetOpt("ShipToOK")), "fine", "fine") & ", " & _
        Plural(Val(GetOpt("ShipToSkipped")), "skipped as inactive", "skipped as inactive") & "."
    If Val(GetOpt("ShipToByCounty")) > 0 Then s = s & " " & Plural(Val(GetOpt("ShipToByCounty")), "was", "were") & _
        " placed through the county, because the state's file has the street under a different mail city."
    ShipToLine = s
End Function

'================================================== the main form's lines

Public Function LoadedLine() As String
    Dim nowQ As String, lastQ As String
    On Error Resume Next
    FindQuarters nowQ, lastQ
    If nowQ = "" Then
        LoadedLine = "Nothing loaded on this PC yet."
    ElseIf lastQ = "" Then
        LoadedLine = "Loaded: " & QLabel(nowQ) & " only. Rates from " & DateText(nowQ) & "."
    Else
        LoadedLine = "Loaded: " & QLabel(lastQ) & " and " & QLabel(nowQ) & ". New rates from " & DateText(nowQ) & "."
    End If
End Function

Public Function ChangesLine() As String
    Dim nowQ As String, lastQ As String, places As Long, unused As Long, streets As Long, zips As Long
    On Error Resume Next
    If ReadChanges(nowQ, lastQ, places, unused, streets, zips) Then
        ChangesLine = "Use tax changes: " & Plural(places, "place", "places") & " with a new rate, " & _
            Plural(streets, "street", "streets") & " where houses moved to a new rate."
    End If
End Function

Public Function ListLine() As String
    Dim n As Long
    On Error Resume Next
    n = DCount("*", "ShipTos")
    If n = 0 Then
        ListLine = "Ship-to list: none loaded yet."
    ElseIf GetOpt("ShipToCheckedFor") = "" Then
        ListLine = "Ship-to list: " & Plural(n, "ship-to", "ship-tos") & ", not checked yet."
    Else
        ListLine = "Ship-to list: " & Plural(n, "ship-to", "ship-tos") & ", checked against " & _
            QLabel(GetOpt("ShipToCheckedFor")) & "."
    End If
End Function

Public Function BigNumber(ByVal kind As String) As String
    On Error Resume Next
    If GetOpt("ShipToCheckedFor") = "" Or DCount("*", "ShipTos") = 0 Then Exit Function
    BigNumber = Format$(Val(GetOpt("ShipTo" & kind)), "#,##0")
End Function

Public Function BigLabel(ByVal kind As String) As String
    If BigNumber(kind) = "" Then Exit Function
    Select Case kind
        Case "Fix": BigLabel = "to fix"
        Case "Map": BigLabel = "to look up on the map"
        Case "OK": BigLabel = "fine"
        Case "Skipped": BigLabel = "inactive, skipped"
    End Select
End Function

' The one thing to click now.
Public Function NextLine() As String
    Dim nowQ As String, lastQ As String, nFix As Long, nMap As Long, s As String
    On Error Resume Next
    FindQuarters nowQ, lastQ
    nFix = Val(GetOpt("ShipToFix"))
    nMap = Val(GetOpt("ShipToMap"))
    If nowQ = "" Then
        s = "load this quarter and the one before. Steps 1 to 3 at the right."
    ElseIf lastQ = "" Then
        s = "click Load new quarter again for the other quarter. It takes two."
    ElseIf DCount("*", "ShipTos") = 0 Then
        s = "click Load ship-to list and pick the list."
    ElseIf GetOpt("ShipToCheckedFor") <> nowQ Then
        s = "click Load ship-to list again, so it's checked against " & QLabel(nowQ) & "."
    ElseIf nFix > 0 And nMap > 0 Then
        s = "click Ship-tos to fix and fix those rates. Then Look up on the map."
    ElseIf nFix > 0 Then
        s = "click Ship-tos to fix and fix those rates in the system."
    ElseIf nMap > 0 Then
        s = "click Look up on the map, type each code, then click Check again."
    Else
        s = "nothing to fix. When next quarter is out, do steps 1 to 3 at the right."
    End If
    NextLine = "Next: " & s
End Function

' The How to use it panel on the main form, as rich text: bold for the
' headings and the two files to download.
Public Function HowToText() As String
    Dim s As String, pad As String, arrow As String
    pad = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    arrow = "&nbsp;&gt;&nbsp;"        ' Access drops a plain space next to the first &gt;
    s = "<b>Each quarter</b><br>"
    s = s & "1. On dor.mo.gov go to Taxation" & arrow & "Business" & arrow & "Sales/Use Tax. Find the box " & _
        "Remote Seller and Marketplace Facilitator. Each quarter there, like Q4 2026, has 3 links.<br>"
    s = s & "2. Under the new quarter, download 2 of the 3 into one folder. Keep their names.<br>"
    s = s & pad & "<b>DOR Boundary Database</b> (a .csv)<br>"
    s = s & pad & "<b>DOR Jurisdiction and Total Rate Database</b> (an .xlsx)<br>"
    s = s & pad & "Skip DOR Rate Database. It isn't used.<br>"
    s = s & "3. Click Load new quarter and pick the .csv. The .xlsx in the same folder loads " & _
        "with it. About a minute.<br>"
    s = s & "4. Use tax changes opens by itself.<br>"
    s = s & "First time only: do 1 to 3 for the quarter before too. It takes two to compare.<br><br>"
    s = s & "<b>Ship-tos</b><br>"
    s = s & "5. When a ship-to list comes, click Load ship-to list and pick it. Every ship-to " & _
        "gets checked, and again each new quarter. A later list only needs the new ship-tos.<br>"
    s = s & "6. Click Ship-tos to fix. Click a number like 100879/4 to copy it, paste it into " & _
        "the system and fix the rate.<br>"
    s = s & "7. Click Look up on the map. Look each address up on the state's map, type the " & _
        "code, click Check again. The code is kept.<br><br>"
    s = s & "<b>Good to know</b><br>"
    s = s & "The yellow Next line always says what to click.<br>"
    s = s & "No Save button. Everything is kept right away.<br>"
    s = s & "The state's data sits on this PC, not in this file: the folder is DataFolder in " & _
        "the Settings table. Keep it off the server. On another PC, load the two quarters again.<br>"
    s = s & "Only the use rate counts. Nothing downloads by itself.<br>"
    s = s & "Inactive ship-tos are skipped. Exempt ones are checked and listed last.<br>"
    s = s & "A changed address in a later list clears that ship-to's typed code."
    HowToText = s
End Function

'================================================== what the buttons call

Public Function LoadQuarter()
    Dim bPath As String
    On Error GoTo Fail
    bPath = PickFile("Pick the DOR Boundary Database file (dor-boundary-database-...csv)", "Boundary file", "*.csv", GetOpt("Folder"))
    If bPath = "" Then Exit Function
    SetOpt "Folder", Left$(bPath, InStrRev(bPath, "\"))
    LoadQuarterFiles bPath, ""
    RefreshLists
    If LastQuarter() <> "" And InStr(gLastMessage, "Done.") = 1 Then ShowChanges
    Exit Function
Fail:
    TidyUp
    Say "Loading stopped: " & Err.Description & " (error " & Err.Number & ")", vbExclamation
End Function

Public Function ShowChanges()
    On Error GoTo Fail
    If LastQuarter() = "" Then
        Say "Two quarters have to be loaded before there's anything to compare.", vbInformation
        Exit Function
    End If
    PointChanges
    DoCmd.OpenReport RPT_CHANGES, acViewPreview
    Exit Function
Fail:
    Say "Couldn't open the report: " & Err.Description, vbExclamation
End Function

' After a load or a check: the main form's lines and both lists.
Private Sub RefreshLists()
    On Error Resume Next
    If CurrentProject.AllForms(FORM_MAIN).IsLoaded Then Forms(FORM_MAIN).Recalc
    If CurrentProject.AllForms(FORM_SHIPTOS).IsLoaded Then Forms(FORM_SHIPTOS).Requery
    If CurrentProject.AllForms(FORM_MAP).IsLoaded Then Forms(FORM_MAP).Requery
End Sub

'=========================================================== loading files

' Loads one quarter: the boundary file, and the total rate file for the same
' quarter (found in the same folder by its DOR name, or asked for). Both go
' into a new file in the data folder. The two newest quarters are kept; an
' older quarter's file is deleted. Returns True when it loaded.
Public Function LoadQuarterFiles(ByVal bPath As String, ByVal rPath As String) As Boolean
    Dim db As DAO.Database, other As DAO.Database, msg As String, asof As Long, nowQ As String, lastQ As String
    Dim q As String, n As Long, fileQ As String, newPath As String, keepNow As String, keepLast As String
    Dim f As String, gone As String, item As Variant, places As Long, unused As Long, streets As Long, zips As Long

    On Error GoTo Fail
    If Dir$(bPath) = "" Then Say "Can't find " & bPath, vbExclamation: Exit Function
    DoCmd.Hourglass True

    asof = ReadFileDate(bPath, msg)
    If asof = 0 Then TidyUp: Say msg, vbExclamation: Exit Function
    q = CStr(asof)

    FindQuarters nowQ, lastQ
    If nowQ <> "" And lastQ <> "" And q < lastQ Then
        TidyUp
        Say "That boundary file is for " & QLabel(q) & ", older than both quarters already loaded (" & _
            QLabel(lastQ) & " and " & QLabel(nowQ) & "). Nothing changed.", vbExclamation
        Exit Function
    End If
    ' the two quarters kept after this load
    If q >= nowQ Then
        keepNow = q
        If q = nowQ Then keepLast = lastQ Else keepLast = nowQ
    Else
        keepNow = nowQ
        keepLast = q
    End If

    ' the rate file for the same quarter
    If rPath = "" Then
        rPath = Left$(bPath, InStrRev(bPath, "\")) & "dor-jurisdiction-total-rate-" & LCase$(QFileTag(q)) & ".xlsx"
        If Dir$(rPath) = "" Then
            DoCmd.Hourglass False
            Say "The rate file isn't next to the .csv, so pick it now: the DOR Jurisdiction and " & _
                "Total Rate Database file for " & QLabel(q) & _
                " (dor-jurisdiction-total-rate-" & LCase$(QFileTag(q)) & ".xlsx).", vbInformation
            rPath = PickFile("Pick the total rate file for " & QLabel(q), "Excel file", "*.xlsx", Left$(bPath, InStrRev(bPath, "\")))
            If rPath = "" Then TidyUp: Exit Function
            DoCmd.Hourglass True
        End If
    End If
    fileQ = QuarterInName(rPath)
    If fileQ <> "" And fileQ <> q Then
        TidyUp
        Say "The rate file looks like " & QLabel(fileQ) & ", but the boundary file is " & QLabel(q) & _
            ". Pick the rate file for the same quarter. Nothing changed.", vbExclamation
        Exit Function
    End If

    ' read both files into a new file first, so a bad file changes nothing
    MakeFolder DataFolder()
    newPath = DataFolder() & "\Loading.accdb"
    KillFile newPath
    Set db = DBEngine.CreateDatabase(newPath, dbLangGeneral, dbVersion120)
    db.Close
    Set db = DBEngine.OpenDatabase(newPath, True)  ' only this load uses it: quicker
    n = ImportBoundary(bPath, db, asof, msg)
    If n > 0 Then
        If Not ImportRates(rPath, db, msg) Then n = 0
    End If
    If n = 0 Then
        db.Close
        Set db = Nothing
        KillFile newPath
        TidyUp
        Say msg, vbExclamation
        Exit Function
    End If
    NewTable db, "Info", "InfoName:T50;InfoValue:T255"
    InfoSet db, "File", FileName(bPath)
    InfoSet db, "Rows", CStr(n)
    InfoSet db, "LoadedOn", Format$(Now, "yyyy-mm-dd hh:nn")

    ' compare the two quarters; the answer is kept in the newer one's file
    CloseReports
    If keepLast <> "" Then
        If keepNow = q Then
            Set other = DBEngine.OpenDatabase(QuarterFile(keepLast), False, True)
            CompareQuarters db, other, keepLast
        Else
            Set other = DBEngine.OpenDatabase(QuarterFile(keepNow))
            CompareQuarters other, db, q
        End If
        other.Close
        Set other = Nothing
    End If
    db.Close
    Set db = Nothing

    ' the new file takes its place, and a third quarter's file goes
    KillFile QuarterFile(q)
    Name newPath As QuarterFile(q)
    newPath = ""
    f = Dir$(DataFolder() & "\Quarter ????-??-??.accdb")
    Do While Len(f) > 0
        If f <> FileName(QuarterFile(keepNow)) And f <> FileName(QuarterFile(keepLast)) Then gone = gone & "|" & f
        f = Dir$()
    Loop
    For Each item In Split(Mid$(gone, 2), "|")
        If Len(item) > 0 Then KillFile DataFolder() & "\" & item
    Next
    PointChanges

    If keepLast = "" Then
        TidyUp
        Say "Loaded " & QLabel(q) & ": " & Format$(n, "#,##0") & " rows. Next, load the other quarter's file " & _
            "too. It takes two quarters to compare.", vbInformation
        LoadQuarterFiles = True
        Exit Function
    End If
    If DCount("*", "ShipTos") > 0 Then CheckShipTos
    TidyUp
    ReadChanges nowQ, lastQ, places, unused, streets, zips
    Say "Done. Use tax, " & QLabel(lastQ) & " to " & QLabel(nowQ) & ": " & _
        Plural(places, "place has", "places have") & " a new use rate; on " & _
        Plural(streets, "street", "streets") & " some houses moved to a different use rate.", vbInformation
    LoadQuarterFiles = True
    Exit Function
Fail:
    msg = "Loading stopped: " & Err.Description & " (error " & Err.Number & ")"
    On Error Resume Next
    If Not other Is Nothing Then other.Close
    If Not db Is Nothing Then db.Close
    KillFile newPath
    TidyUp
    Say msg, vbExclamation
End Function

' Closes the Use tax changes report, which holds this quarter's file open.
Private Sub CloseReports()
    On Error Resume Next
    DoCmd.Close acReport, RPT_CHANGES, acSaveNo
End Sub

' First pass over the boundary file: check it is one, and find its date (the
' newest Beginning Effective Date; the file answers from that day on).
Private Function ReadFileDate(ByVal path As String, ByRef msg As String) As Long
    Dim f As Integer, s As String, p1 As Long, p2 As Long, d As Long, mx As Long, n As Long, total As Double
    f = FreeFile
    Open path For Input Access Read Shared As #f
    total = LOF(f)
    If EOF(f) Then
        Close #f
        msg = "That file is empty."
        Exit Function
    End If
    Line Input #f, s
    s = NoMark(s)
    If LCase$(Left$(s, 29)) <> "type,beginning effective date" Then
        Close #f
        msg = "That doesn't look like the DOR Boundary Database file. Its first line should start with " & _
            """type,Beginning Effective Date"". Pick the dor-boundary-database file (.csv)."
        Exit Function
    End If
    SysCmd acSysCmdInitMeter, "Checking the boundary file", 100
    Do While Not EOF(f)
        Line Input #f, s
        p1 = InStr(1, s, ",")
        If p1 > 0 Then
            p2 = InStr(p1 + 1, s, ",")
            If p2 > p1 + 1 Then
                d = CLng(Val(Mid$(s, p1 + 1, p2 - p1 - 1)))
                If d > mx Then mx = d
            End If
        End If
        n = n + 1
        If n Mod 20000 = 0 Then
            SysCmd acSysCmdUpdateMeter, Int(100# * Seek(f) / (total + 1))
            DoEvents
        End If
    Loop
    Close #f
    SysCmd acSysCmdRemoveMeter
    If mx < 20000101 Or mx > 29991231 Then
        msg = "Couldn't find the dates in that file."
        Exit Function
    End If
    ReadFileDate = mx
End Function

' Second pass: the rows in force on the file's date go into the new file's
' Boundary table, with the street worked out once (the standard spelling used
' for matching). Returns the number of rows kept, 0 if it failed (msg says why).
Private Function ImportBoundary(ByVal path As String, db As DAO.Database, ByVal asof As Long, _
        ByRef msg As String) As Long
    Dim rs As DAO.Recordset, f As Integer, s As String, a() As String, tbl As String
    Dim i As Long, code As String, kind As String, lo As String, hi As String
    Dim b As Long, e As Long, n As Long, kept As Long, bad As Long, total As Double, t As Long
    Dim pieces() As String, sd As String, k As String, lastStreet As String, lastCity As String
    Dim street As String, city As String, cityKey As String
    Dim fKind As DAO.Field, fLo As DAO.Field, fHi As DAO.Field, fLoTxt As DAO.Field, fHiTxt As DAO.Field
    Dim fSide As DAO.Field, fStreet As DAO.Field, fName As DAO.Field, fPre As DAO.Field, fSuf As DAO.Field
    Dim fPost As DAO.Field, fFull As DAO.Field, fCity As DAO.Field, fCityKey As DAO.Field, fZip As DAO.Field
    Dim fZipLo As DAO.Field, fZipHi As DAO.Field, fCode As DAO.Field, fCounty As DAO.Field

    tbl = "Boundary"
    NewTable db, tbl, BOUNDARY_SPEC
    InitWords
    Set mKeys = New Collection
    Set mCityKeys = New Collection

    Set rs = db.OpenRecordset(tbl, dbOpenTable)
    Set fKind = rs.Fields("Kind"): Set fLo = rs.Fields("Lo"): Set fHi = rs.Fields("Hi")
    Set fLoTxt = rs.Fields("LoTxt"): Set fHiTxt = rs.Fields("HiTxt"): Set fSide = rs.Fields("Side")
    Set fStreet = rs.Fields("Street"): Set fName = rs.Fields("NameKey"): Set fPre = rs.Fields("PreKey")
    Set fSuf = rs.Fields("SufKey"): Set fPost = rs.Fields("PostKey"): Set fFull = rs.Fields("FullKey")
    Set fCity = rs.Fields("City"): Set fCityKey = rs.Fields("CityKey"): Set fZip = rs.Fields("Zip")
    Set fZipLo = rs.Fields("ZipLo"): Set fZipHi = rs.Fields("ZipHi"): Set fCode = rs.Fields("JCode")
    Set fCounty = rs.Fields("County")

    f = FreeFile
    Open path For Input Access Read Shared As #f
    total = LOF(f)
    Line Input #f, s
    lastStreet = Chr$(0)
    lastCity = Chr$(0)
    SysCmd acSysCmdInitMeter, "Reading the boundary file for " & QLabel(CStr(asof)), 100
    Do While Not EOF(f)
        Line Input #f, s
        n = n + 1
        ' the code is found by its look, never by the header (the header is
        ' shifted): the first field after the 21st that looks like one
        code = ""
        If InStr(1, s, """") > 0 Then
            a = SplitQuoted(s)
            For i = 21 To UBound(a)
                If LooksLikeCode(a(i)) Then code = a(i): Exit For
            Next
        Else
            ' the 21 fields used, then the rest in one piece (quicker than all 90)
            a = Split(s, ",", 22)
            If UBound(a) = 21 Then code = CodeIn(a(21))
        End If
        If code = "" Or UBound(a) < 20 Then
            bad = bad + 1
        Else
            b = CLng(Val(a(1)))
            e = CLng(Val(a(2)))
            kind = a(0)
            If b <= asof And asof <= e And (kind = "A" Or kind = "4" Or kind = "Z") Then
                rs.AddNew
                fKind.Value = kind
                fCode.Value = code
                fCounty.Value = Mid$(code, 7, 3)
                If kind = "A" Then
                    lo = Trim$(a(3))
                    hi = Trim$(a(4))
                    If IsDigits(lo) And IsDigits(hi) Then
                        b = CLng(lo)
                        e = CLng(hi)
                        If b > e Then t = b: b = e: e = t
                        fLo.Value = b
                        fHi.Value = e
                    Else
                        fLoTxt.Value = lo
                        fHiTxt.Value = hi
                    End If
                    sd = Trim$(a(5))
                    If sd <> "O" And sd <> "E" Then sd = "B"
                    fSide.Value = sd
                    ' most rows are on the same street and city as the row before
                    k = a(6) & "|" & a(7) & "|" & a(8) & "|" & a(9)
                    If k <> lastStreet Then
                        lastStreet = k
                        street = Left$(JoinWords(Trim$(a(6)), Trim$(a(7)), Trim$(a(8)), Trim$(a(9))), 80)
                        pieces = Split(StreetPieces(a(6), a(7), a(8), a(9)), "|")
                        pieces(0) = Left$(pieces(0), 80)
                        pieces(1) = Left$(pieces(1), 4)
                        pieces(2) = Left$(pieces(2), 8)
                        pieces(3) = Left$(pieces(3), 4)
                        pieces(4) = Left$(pieces(4), 100)
                    End If
                    If a(14) <> lastCity Then
                        lastCity = a(14)
                        city = Left$(Trim$(a(14)), 40)
                        cityKey = Left$(CachedCityKey(a(14)), 40)
                    End If
                    fStreet.Value = street
                    fName.Value = pieces(0)
                    fPre.Value = pieces(1)
                    fSuf.Value = pieces(2)
                    fPost.Value = pieces(3)
                    fFull.Value = pieces(4)
                    fCity.Value = city
                    fCityKey.Value = cityKey
                    fZip.Value = Left$(Trim$(a(15)), 5)
                Else
                    fZipLo.Value = Left$(Trim$(a(17)) & Trim$(a(18)), 9)
                    fZipHi.Value = Left$(Trim$(a(19)) & Trim$(a(20)), 9)
                End If
                rs.Update
                kept = kept + 1
            End If
        End If
        If n Mod 5000 = 0 Then
            SysCmd acSysCmdUpdateMeter, Int(100# * Seek(f) / (total + 1))
            DoEvents
        End If
    Loop
    Close #f
    rs.Close
    Set mKeys = Nothing
    Set mCityKeys = Nothing

    If kept = 0 Then
        msg = "No rows in that file were in force on " & DateText(CStr(asof)) & ", so nothing was loaded."
        Exit Function
    End If
    If bad > n \ 100 Then
        msg = Format$(bad, "#,##0") & " of " & Format$(n, "#,##0") & " rows in that file couldn't be read, " & _
            "so nothing was loaded. Has the DOR changed the file's layout?"
        Exit Function
    End If
    SysCmd acSysCmdSetStatus, "Indexing the boundary rows"
    db.Execute "CREATE INDEX idxStreet ON [" & tbl & "] (Zip, FullKey)", dbFailOnError
    db.Execute "CREATE INDEX idxCity ON [" & tbl & "] (CityKey, NameKey)", dbFailOnError
    db.Execute "CREATE INDEX idxCode ON [" & tbl & "] (JCode)", dbFailOnError
    db.Execute "CREATE INDEX idxCounty ON [" & tbl & "] (County, NameKey)", dbFailOnError
    SysCmd acSysCmdClearStatus
    ImportBoundary = kept
End Function

' The total rate file: one row per jurisdiction code with its Sales and Use
' rates. Rates come as 0.06975 in some quarters and "6.9750%" in others, and
' Q4 2026 mixed the two inside one column, so the file is read through Excel,
' cell by cell as written (Access's own import blanks the odd ones out).
Private Function ImportRates(ByVal path As String, db As DAO.Database, ByRef msg As String) As Boolean
    Dim out As DAO.Recordset, v As Variant, r As Long, c As Long
    Dim iCode As Long, iSales As Long, iUse As Long, nm As String, code As String, n As Long

    v = ReadSheet(path, msg)
    If IsEmpty(v) Then Exit Function
    iCode = 0: iSales = 0: iUse = 0
    For c = 1 To UBound(v, 2)
        nm = LCase$(Replace(CellText(v(1, c)), " ", ""))
        If nm = "jurisdictioncode" Then iCode = c
        If nm = "sales" Then iSales = c
        If nm = "use" Then iUse = c
    Next
    If iCode = 0 Or iSales = 0 Or iUse = 0 Then
        msg = "That doesn't look like the DOR Jurisdiction and Total Rate Database file (not DOR Rate " & _
            "Database): it needs columns named Jurisdiction Code, Sales and Use. Nothing changed."
        Exit Function
    End If

    NewTable db, "Rates", "JCode:T13;SalesRate:D;UseRate:D"
    db.Execute "CREATE UNIQUE INDEX PrimaryKey ON Rates (JCode) WITH PRIMARY", dbFailOnError
    Set out = db.OpenRecordset("Rates", dbOpenTable)
    out.Index = "PrimaryKey"
    For r = 2 To UBound(v, 1)
        code = Trim$(CellText(v(r, iCode)))
        If LooksLikeCode(code) Then
            ' a code listed twice: the later row wins
            out.Seek "=", code
            If out.NoMatch Then
                out.AddNew
                out!JCode = code
                n = n + 1
            Else
                out.Edit
            End If
            out!SalesRate = RateValue(v(r, iSales))
            out!UseRate = RateValue(v(r, iUse))
            out.Update
        End If
    Next
    out.Close
    If n = 0 Then
        msg = "No jurisdiction codes found in the rate file. Nothing changed."
        Exit Function
    End If
    ImportRates = True
End Function

' The first sheet of an Excel file as a grid of values, row 1 first. Empty
' (and msg says why) if Excel couldn't open it. Excel is let go piece by
' piece, or a hidden copy of it can stay open until Access closes.
Private Function ReadSheet(ByVal path As String, ByRef msg As String) As Variant
    Dim xl As Object, wb As Object, ws As Object, v As Variant, lastRow As Long, lastCol As Long
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(path, 0, True)
    Set ws = wb.Worksheets(1)
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With
    If lastRow = 1 And lastCol = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = ws.Cells(1, 1).Value2
    Else
        v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    End If
    Set ws = Nothing
    wb.Close False
    Set wb = Nothing
    xl.Quit
    Set xl = Nothing
    ReadSheet = v
    Exit Function
Fail:
    msg = "Couldn't open " & FileName(path) & " in Excel: " & Err.Description & ". Nothing changed."
    On Error Resume Next
    Set ws = Nothing
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    If Not xl Is Nothing Then xl.Quit
    Set xl = Nothing
    ReadSheet = Empty
End Function

' A cell's value as text ("" for empty or an error).
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = CStr(v)
End Function

' 0.06975, "6.9750%" or "0.06975" -> 0.06975. Empty -> Null.
Private Function RateValue(ByVal v As Variant) As Variant
    Dim s As String
    RateValue = Null
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(v)
        If s = "" Then Exit Function
        If Right$(s, 1) = "%" Then
            RateValue = Round(Val(Left$(s, Len(s) - 1)) / 100, 6)
        Else
            RateValue = Round(Val(s), 6)
        End If
    Else
        RateValue = Round(CDbl(v), 6)
    End If
End Function

'============================================== step one: what changed

' This quarter's file (dbN) against last quarter's (dbL). The answer goes in
' this quarter's file: RateChanges, Moves, MoveGroups and the counts in Info.
Private Sub CompareQuarters(dbN As DAO.Database, dbL As DAO.Database, ByVal lastQ As String)
    Dim useL As Collection, useN As Collection
    Set useL = RateMap(dbL)
    Set useN = RateMap(dbN)
    LoadCountyNames
    SysCmd acSysCmdSetStatus, "Finding the places whose use rate changed"
    InfoSet dbN, "ComparedWith", ""
    CompareRates dbN, useL
    CompareStreets dbN, dbL, useL, useN
    InfoSet dbN, "ComparedOn", Format$(Now, "yyyy-mm-dd hh:nn")
    InfoSet dbN, "ComparedWith", lastQ
    SysCmd acSysCmdClearStatus
End Sub

' Every code's use rate in one quarter's file (Null where the file has none).
Private Function RateMap(db As DAO.Database) As Collection
    Dim rs As DAO.Recordset, c As New Collection
    Set rs = db.OpenRecordset("SELECT JCode, UseRate FROM Rates", dbOpenSnapshot)
    Do While Not rs.EOF
        If IsNull(rs!UseRate) Then c.Add Null, CStr(rs!JCode) Else c.Add CDbl(rs!UseRate), CStr(rs!JCode)
        rs.MoveNext
    Loop
    rs.Close
    Set RateMap = c
End Function

Private Sub LoadCountyNames()
    Dim rs As DAO.Recordset
    Set mCounty = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT County, CountyName FROM Counties", dbOpenSnapshot)
    Do While Not rs.EOF
        mCounty.Add CStr(rs!CountyName), CStr(rs!County)
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function CountyName(ByVal county As String) As String
    On Error Resume Next
    CountyName = mCounty(county)
    If Err.Number <> 0 Then CountyName = "County " & county
End Function

' Codes in both quarters' rate files whose use rate changed, with the mail
' cities their addresses use. db is this quarter's file.
Private Sub CompareRates(db As DAO.Database, useL As Collection)
    Dim rs As DAO.Recordset, out As DAO.Recordset, code As String, prev As String, vL As Variant, vN As Variant
    Dim keys() As String, names() As String, cnts() As Long, k As Long, nk As Long, moved As Boolean, sql As String

    DropTable db, "RateChanges"
    NewTable db, "RateChanges", RATECHANGES_SPEC
    Set rs = db.OpenRecordset("SELECT JCode, UseRate FROM Rates", dbOpenSnapshot)
    Set out = db.OpenRecordset("RateChanges", dbOpenTable)
    Do While Not rs.EOF
        code = rs!JCode
        If InSet(useL, code) Then
            vL = useL(code)
            vN = rs!UseRate
            If IsNull(vL) Or IsNull(vN) Then
                moved = (IsNull(vL) <> IsNull(vN))
            Else
                moved = (Abs(vN - vL) > 0.0000001)
            End If
            If moved Then
                out.AddNew
                out!JCode = code
                out!County = Mid$(code, 7, 3)
                out!CountyName = CountyName(Mid$(code, 7, 3))
                out!OldUse = vL
                out!NewUse = vN
                out!OldRate = RateText(vL)
                out!NewRate = RateText(vN)
                out.Update
            End If
        End If
        rs.MoveNext
    Loop
    out.Close
    rs.Close

    ' mail cities: grouped the way the state's spellings meet (SAINT LOUIS = ST. LOUIS)
    Set rs = db.OpenRecordset("SELECT B.JCode, B.CityKey, B.City, Count(*) AS N FROM Boundary AS B " & _
        "INNER JOIN RateChanges AS R ON B.JCode = R.JCode WHERE B.Kind='A' " & _
        "GROUP BY B.JCode, B.CityKey, B.City ORDER BY B.JCode", dbOpenSnapshot)
    ReDim keys(0 To 15): ReDim names(0 To 15): ReDim cnts(0 To 15)
    prev = ""
    Do
        If rs.EOF Then code = Chr$(1) Else code = rs!JCode
        If code <> prev And prev <> "" Then
            SaveCities db, prev, keys, names, cnts, nk
            nk = 0
        End If
        If rs.EOF Then Exit Do
        prev = code
        ' same city key: keep the alphabetically first spelling, add the counts
        For k = 0 To nk - 1
            If keys(k) = Nz(rs!CityKey, "") Then Exit For
        Next
        If k = nk Then
            If nk > UBound(keys) Then
                ReDim Preserve keys(0 To 2 * nk): ReDim Preserve names(0 To 2 * nk): ReDim Preserve cnts(0 To 2 * nk)
            End If
            keys(nk) = Nz(rs!CityKey, ""): names(nk) = Nz(rs!City, ""): cnts(nk) = 0
            nk = nk + 1
        ElseIf StrComp(Nz(rs!City, ""), names(k), vbBinaryCompare) < 0 Then
            names(k) = Nz(rs!City, "")
        End If
        cnts(k) = cnts(k) + rs!N
        rs.MoveNext
    Loop
    rs.Close

    ' a code only zip+4 or zip rows use still has addresses; one no row uses has none
    Set rs = db.OpenRecordset("SELECT B.JCode, Count(*) AS N FROM Boundary AS B INNER JOIN RateChanges AS R " & _
        "ON B.JCode = R.JCode WHERE B.Kind<>'A' GROUP BY B.JCode", dbOpenSnapshot)
    Do While Not rs.EOF
        db.Execute "UPDATE RateChanges SET ZipRows=" & rs!N & " WHERE JCode='" & rs!JCode & "'", dbFailOnError
        rs.MoveNext
    Loop
    rs.Close
    sql = "SELECT Count(*) FROM RateChanges WHERE IIf(AddrRows Is Null,0,AddrRows)+IIf(ZipRows Is Null,0,ZipRows)"
    InfoSet db, "RateChangeCount", CStr(CountSql(db, sql & ">0"))
    InfoSet db, "RateChangeUnused", CStr(CountSql(db, sql & "=0"))
End Sub

Private Sub SaveCities(db As DAO.Database, ByVal code As String, keys() As String, names() As String, _
        cnts() As Long, ByVal nk As Long)
    Dim i As Long, j As Long, total As Long, all As String, shown As String, ts As String, tl As Long
    ' most addresses first, then by city key
    For i = 1 To nk - 1
        For j = i To 1 Step -1
            If cnts(j) > cnts(j - 1) Or (cnts(j) = cnts(j - 1) And StrComp(keys(j), keys(j - 1), vbBinaryCompare) < 0) Then
                ts = keys(j): keys(j) = keys(j - 1): keys(j - 1) = ts
                ts = names(j): names(j) = names(j - 1): names(j - 1) = ts
                tl = cnts(j): cnts(j) = cnts(j - 1): cnts(j - 1) = tl
            Else
                Exit For
            End If
        Next
    Next
    For i = 0 To nk - 1
        total = total + cnts(i)
        If i > 0 Then all = all & "|"
        all = all & names(i)
        If i < 3 Then
            If i > 0 Then shown = shown & ", "
            shown = shown & names(i)
        End If
    Next
    If nk > 3 Then shown = shown & " and " & (nk - 3) & " more"
    db.Execute "UPDATE RateChanges SET Cities='" & Replace(Left$(shown, 255), "'", "''") & "', CitiesAll='" & _
        Replace(all, "'", "''") & "', AddrRows=" & total & " WHERE JCode='" & code & "'", dbFailOnError
End Sub

' Every street in both quarters, compared house number by house number. The
' state re-splits ranges between quarters, so rows can't be matched one to one.
' db is this quarter's file and gets the answer; dbL is last quarter's.
Private Sub CompareStreets(db As DAO.Database, dbL As DAO.Database, useL As Collection, useN As Collection)
    Dim cN As Cursor, cL As Cursor, rsOut As DAO.Recordset
    Dim L As StreetRows, N As StreetRows, zip As String, key As String, done As Long, total As Long

    DropTable db, "Moves"
    NewTable db, "Moves", MOVES_SPEC
    InitRows L
    InitRows N
    ReDim mCodes(0 To 15)
    total = CountSql(db, "SELECT Count(*) FROM Boundary")
    OpenCursor cN, db, "idxStreet", "Zip"
    OpenCursor cL, dbL, "idxStreet", "Zip"
    Set rsOut = db.OpenRecordset("Moves", dbOpenTable)
    SysCmd acSysCmdInitMeter, "Comparing streets", 100
    If Not cN.rs.EOF Then cN.rs.MoveFirst
    Do While Not cN.rs.EOF
        zip = cN.zip.Value & ""
        key = cN.full.Value & ""
        If Len(key) = 0 Then
            cN.rs.MoveNext
            done = done + 1
        Else
            ReadStreet cN, zip, key, N, True
            done = done + N.n
            cL.rs.Seek "=", zip, key
            If Not cL.rs.NoMatch Then
                ReadStreet cL, zip, key, L, False
                CompareStreet zip, L, N, useL, useN, rsOut
            End If
        End If
        If done Mod 2000 < N.n + 1 Then
            SysCmd acSysCmdUpdateMeter, Int(100# * done / (total + 1))
            DoEvents
        End If
    Loop
    rsOut.Close
    cL.rs.Close
    cN.rs.Close
    SysCmd acSysCmdRemoveMeter
    GroupMoves db
    InfoSet db, "MoveLineCount", CStr(CountSql(db, "SELECT Count(*) FROM Moves"))
    InfoSet db, "MoveStreetCount", CStr(CountSql(db, "SELECT Count(*) FROM (SELECT DISTINCT Zip, Street FROM Moves)"))
    InfoSet db, "MoveZipCount", CStr(CountSql(db, "SELECT Count(*) FROM (SELECT DISTINCT Zip FROM Moves)"))
End Sub

' One report line per street and code change, its house numbers together
' (940; 950; 991 to 995, odd side).
Private Sub GroupMoves(db As DAO.Database)
    Dim rs As DAO.Recordset, out As DAO.Recordset, k As String, prevK As String, open1 As Boolean
    DropTable db, "MoveGroups"
    NewTable db, "MoveGroups", MOVEGROUPS_SPEC
    Set rs = db.OpenRecordset("SELECT * FROM Moves ORDER BY Zip, Street, OldCodes, NewCodes, LoNum, LoTxt, Side", _
        dbOpenSnapshot)
    Set out = db.OpenRecordset("MoveGroups", dbOpenTable)
    Do While Not rs.EOF
        k = rs!zip & "|" & rs!street & "|" & rs!oldCodes & "|" & rs!newCodes
        If k <> prevK Then
            If open1 Then out.Update
            out.AddNew
            open1 = True
            out!County = rs!County
            out!CountyName = rs!CountyName
            out!city = rs!city
            out!zip = rs!zip
            out!street = rs!street
            out!LoNum = rs!LoNum
            out!houses = rs!houses
            out!oldCodes = rs!oldCodes
            out!newCodes = rs!newCodes
            out!OldUse = rs!OldUse
            out!NewUse = rs!NewUse
            out!OldRate = rs!OldRate
            out!NewRate = rs!NewRate
            prevK = k
        Else
            out!houses = out!houses & "; " & rs!houses
        End If
        rs.MoveNext
    Loop
    If open1 Then out.Update
    out.Close
    rs.Close
End Sub

Private Sub InitRows(sr As StreetRows)
    sr.n = 0
    ReDim sr.num(0 To 63): ReDim sr.lo(0 To 63): ReDim sr.hi(0 To 63)
    ReDim sr.loTxt(0 To 63): ReDim sr.hiTxt(0 To 63): ReDim sr.side(0 To 63)
    ReDim sr.code(0 To 63): ReDim sr.street(0 To 63): ReDim sr.city(0 To 63)
End Sub

Private Sub GrowRows(sr As StreetRows)
    Dim m As Long
    m = 2 * (UBound(sr.lo) + 1) - 1
    ReDim Preserve sr.num(0 To m): ReDim Preserve sr.lo(0 To m): ReDim Preserve sr.hi(0 To m)
    ReDim Preserve sr.loTxt(0 To m): ReDim Preserve sr.hiTxt(0 To m): ReDim Preserve sr.side(0 To m)
    ReDim Preserve sr.code(0 To m): ReDim Preserve sr.street(0 To m): ReDim Preserve sr.city(0 To m)
End Sub

' Reads the rows of one street (same zip and street) and moves past them. The
' street's and city's names are only needed from this quarter.
Private Sub ReadStreet(cur As Cursor, ByVal zip As String, ByVal key As String, sr As StreetRows, _
        ByVal names As Boolean)
    sr.n = 0
    Do While Not cur.rs.EOF
        If cur.zip.Value & "" <> zip Or cur.full.Value & "" <> key Then Exit Do
        If sr.n > UBound(sr.lo) Then GrowRows sr
        If IsNull(cur.lo.Value) Then
            sr.num(sr.n) = False
            sr.loTxt(sr.n) = cur.loTxt.Value & ""
            sr.hiTxt(sr.n) = cur.hiTxt.Value & ""
        Else
            sr.num(sr.n) = True
            sr.lo(sr.n) = cur.lo.Value
            sr.hi(sr.n) = cur.hi.Value
        End If
        If IsNull(cur.side.Value) Then sr.side(sr.n) = "B" Else sr.side(sr.n) = cur.side.Value
        sr.code(sr.n) = cur.code.Value & ""
        If names Then
            sr.street(sr.n) = cur.street.Value & ""
            sr.city(sr.n) = cur.city.Value & ""
        End If
        sr.n = sr.n + 1
        cur.rs.MoveNext
    Loop
End Sub

Private Sub CompareStreet(ByVal zip As String, L As StreetRows, N As StreetRows, useL As Collection, _
        useN As Collection, rsOut As DAO.Recordset)
    Dim streetName As String, cityName As String
    If Not SameRows(L, N) Then
        If NumberSig(L) <> NumberSig(N) Then CompareNumbers zip, L, N, useL, useN, rsOut, streetName, cityName
    End If
    CompareLettered zip, L, N, useL, useN, rsOut, streetName, cityName
End Sub

' True when both quarters list the street's numbered rows alike and in the
' same order: nothing on it moved. Most streets; quicker than NumberSig.
Private Function SameRows(L As StreetRows, N As StreetRows) As Boolean
    Dim i As Long, j As Long
    i = NextNumbered(L, 0)
    j = NextNumbered(N, 0)
    Do While i < L.n And j < N.n
        If L.lo(i) <> N.lo(j) Or L.hi(i) <> N.hi(j) Or L.side(i) <> N.side(j) Or L.code(i) <> N.code(j) Then Exit Function
        i = NextNumbered(L, i + 1)
        j = NextNumbered(N, j + 1)
    Loop
    SameRows = (i >= L.n And j >= N.n)
End Function

' The first numbered row from i on (sr.n if there's none).
Private Function NextNumbered(sr As StreetRows, ByVal i As Long) As Long
    Do While i < sr.n
        If sr.num(i) Then Exit Do
        i = i + 1
    Loop
    NextNumbered = i
End Function

' The numbered rows of a street as one sorted string, to spot a street that
' didn't change at all.
Private Function NumberSig(sr As StreetRows) As String
    Dim s() As String, i As Long, m As Long
    If sr.n = 0 Then Exit Function
    ReDim s(0 To sr.n - 1)
    For i = 0 To sr.n - 1
        If sr.num(i) Then
            s(m) = sr.lo(i) & "|" & sr.hi(i) & "|" & sr.side(i) & "|" & sr.code(i)
            m = m + 1
        End If
    Next
    SortStrings s, m
    NumberSig = JoinN(s, m, ";")
End Function

Private Sub CompareNumbers(ByVal zip As String, L As StreetRows, N As StreetRows, useL As Collection, _
        useN As Collection, rsOut As DAO.Recordset, streetName As String, cityName As String)
    Dim pts() As Long, np As Long, i As Long, k As Long, a As Long, b As Long
    Dim parI As Integer, par As String, f As Long, lst As Long, eL As String, eN As String, moved As Boolean
    Dim active As Boolean, rFirst As Long, rLast As Long, rEL As String, rEN As String
    Dim oFirst() As Long, oLast() As Long, oEL() As String, oEN() As String, no As Long
    Dim eFirst() As Long, eLast() As Long, eEL() As String, eEN() As String, eUsed() As Boolean, ne As Long

    ' every place a row starts or stops, from both quarters
    ReDim pts(0 To 2 * (L.n + N.n) + 1)
    For i = 0 To L.n - 1
        If L.num(i) Then pts(np) = L.lo(i): pts(np + 1) = L.hi(i) + 1: np = np + 2
    Next
    For i = 0 To N.n - 1
        If N.num(i) Then pts(np) = N.lo(i): pts(np + 1) = N.hi(i) + 1: np = np + 2
    Next
    If np < 2 Then Exit Sub
    SortLongs pts, np
    k = 0
    For i = 1 To np - 1
        If pts(i) <> pts(k) Then k = k + 1: pts(k) = pts(i)
    Next
    np = k + 1

    ReDim oFirst(0 To np): ReDim oLast(0 To np): ReDim oEL(0 To np): ReDim oEN(0 To np)
    ReDim eFirst(0 To np): ReDim eLast(0 To np): ReDim eEL(0 To np): ReDim eEN(0 To np)
    For parI = 1 To 2
        If parI = 1 Then par = "O" Else par = "E"
        active = False
        For k = 0 To np - 2
            a = pts(k)
            b = pts(k + 1) - 1
            If SideRange(a, b, par, f, lst) Then
                eL = Effective(L, a, b, par)
                eN = Effective(N, a, b, par)
                If eL <> "" Or eN <> "" Then
                    moved = False
                    If eL <> "" And eN <> "" And eL <> eN Then
                        moved = (UseSet(eL, useL) <> UseSet(eN, useN))
                    End If
                    If moved And active And rEL = eL And rEN = eN Then
                        rLast = lst
                    Else
                        If active Then
                            If par = "O" Then
                                oFirst(no) = rFirst: oLast(no) = rLast: oEL(no) = rEL: oEN(no) = rEN: no = no + 1
                            Else
                                eFirst(ne) = rFirst: eLast(ne) = rLast: eEL(ne) = rEL: eEN(ne) = rEN: ne = ne + 1
                            End If
                        End If
                        active = moved
                        If moved Then rFirst = f: rLast = lst: rEL = eL: rEN = eN
                    End If
                End If
            End If
        Next
        If active Then
            If par = "O" Then
                oFirst(no) = rFirst: oLast(no) = rLast: oEL(no) = rEL: oEN(no) = rEN: no = no + 1
            Else
                eFirst(ne) = rFirst: eLast(ne) = rLast: eEL(ne) = rEL: eEN(ne) = rEN: ne = ne + 1
            End If
        End If
    Next
    If no = 0 And ne = 0 Then Exit Sub

    If streetName = "" Then
        streetName = MostCommon(N.street, N.n)
        cityName = MostCommon(N.city, N.n)
    End If
    ' an odd run and an even run over the same houses with the same codes: both sides
    ReDim eUsed(0 To np)
    For i = 0 To no - 1
        For k = 0 To ne - 1
            If Not eUsed(k) Then
                If eEL(k) = oEL(i) And eEN(k) = oEN(i) And Abs(eFirst(k) - oFirst(i)) <= 1 _
                        And Abs(eLast(k) - oLast(i)) <= 1 Then Exit For
            End If
        Next
        If k < ne Then
            eUsed(k) = True
            AddMove rsOut, zip, streetName, cityName, "B", CStr(IIf(oFirst(i) < eFirst(k), oFirst(i), eFirst(k))), _
                CStr(IIf(oLast(i) > eLast(k), oLast(i), eLast(k))), oEL(i), oEN(i), useL, useN
        Else
            AddMove rsOut, zip, streetName, cityName, "O", CStr(oFirst(i)), CStr(oLast(i)), oEL(i), oEN(i), useL, useN
        End If
    Next
    For k = 0 To ne - 1
        If Not eUsed(k) Then
            AddMove rsOut, zip, streetName, cityName, "E", CStr(eFirst(k)), CStr(eLast(k)), eEL(k), eEN(k), useL, useN
        End If
    Next
End Sub

' First and last house number of one side (odd or even) between a and b.
Private Function SideRange(ByVal a As Long, ByVal b As Long, ByVal par As String, f As Long, l As Long) As Boolean
    If par = "O" Then
        If a Mod 2 = 1 Then f = a Else f = a + 1
        If b Mod 2 = 1 Then l = b Else l = b - 1
    Else
        If a Mod 2 = 0 Then f = a Else f = a + 1
        If b Mod 2 = 0 Then l = b Else l = b - 1
    End If
    SideRange = (f <= l)
End Function

' The codes in force for house numbers a to b on one side: the narrowest rows
' covering them win (the state lists a wide range, then the houses inside it
' that sit somewhere else). "" if no row covers them.
Private Function Effective(sr As StreetRows, ByVal a As Long, ByVal b As Long, ByVal par As String) As String
    Dim i As Long, w As Long, best As Long, nc As Long
    best = -1
    For i = 0 To sr.n - 1
        If sr.num(i) Then
            If sr.lo(i) <= a And sr.hi(i) >= b Then
                If sr.side(i) = "B" Or sr.side(i) = par Then
                    w = sr.hi(i) - sr.lo(i)
                    If best = -1 Or w < best Then
                        best = w
                        nc = 0
                        AddUnique mCodes, nc, sr.code(i)
                    ElseIf w = best Then
                        AddUnique mCodes, nc, sr.code(i)
                    End If
                End If
            End If
        End If
    Next
    If nc = 0 Then Exit Function
    If nc = 1 Then Effective = mCodes(0): Exit Function
    SortStrings mCodes, nc
    Effective = JoinN(mCodes, nc, ", ")
End Function

' Lettered house numbers (1616B) can't be put on a number line; the same range,
' written the same way in both quarters, is compared as it stands.
Private Sub CompareLettered(ByVal zip As String, L As StreetRows, N As StreetRows, useL As Collection, _
        useN As Collection, rsOut As DAO.Recordset, streetName As String, cityName As String)
    Dim kL() As String, cL() As String, nL As Long, kN() As String, cN() As String, nN As Long
    Dim i As Long, j As Long, p() As String
    For i = 0 To L.n - 1
        If Not L.num(i) Then Exit For
    Next
    If i = L.n Then Exit Sub   ' no lettered house numbers
    GroupLettered L, kL, cL, nL
    If nL = 0 Then Exit Sub
    GroupLettered N, kN, cN, nN
    For i = 0 To nN - 1
        For j = 0 To nL - 1
            If kL(j) = kN(i) Then Exit For
        Next
        If j < nL Then
            If cL(j) <> cN(i) Then
                If UseSet(cL(j), useL) <> UseSet(cN(i), useN) Then
                    If streetName = "" Then
                        streetName = MostCommon(N.street, N.n)
                        cityName = MostCommon(N.city, N.n)
                    End If
                    p = Split(kN(i), "|")
                    AddMove rsOut, zip, streetName, cityName, p(2), p(0), p(1), cL(j), cN(i), useL, useN
                End If
            End If
        End If
    Next
End Sub

Private Sub GroupLettered(sr As StreetRows, keys() As String, codes() As String, nk As Long)
    Dim i As Long, k As Long, key As String, one() As String, cnt As Long
    nk = 0
    ReDim keys(0 To sr.n): ReDim codes(0 To sr.n)
    For i = 0 To sr.n - 1
        If Not sr.num(i) Then
            key = sr.loTxt(i) & "|" & sr.hiTxt(i) & "|" & sr.side(i)
            For k = 0 To nk - 1
                If keys(k) = key Then Exit For
            Next
            If k = nk Then keys(nk) = key: codes(nk) = "": nk = nk + 1
            ' add the code to this key's sorted list
            If codes(k) = "" Then
                codes(k) = sr.code(i)
            ElseIf InStr(1, ", " & codes(k) & ", ", ", " & sr.code(i) & ", ") = 0 Then
                one = Split(codes(k) & ", " & sr.code(i), ", ")
                cnt = UBound(one) + 1
                SortStrings one, cnt
                codes(k) = JoinN(one, cnt, ", ")
            End If
        End If
    Next
End Sub

' The use rates of a list of codes, as one string to compare.
Private Function UseSet(ByVal codes As String, useC As Collection) As String
    Dim c As Variant, r() As String, nr As Long, v As Variant
    ReDim r(0 To 31)
    For Each c In Split(codes, ", ")
        v = UseOf(useC, CStr(c))
        If IsNull(v) Then AddUnique r, nr, "none" Else AddUnique r, nr, Format$(v, "0.000000")
    Next
    SortStrings r, nr
    UseSet = JoinN(r, nr, "|")
End Function

Private Function UseOf(useC As Collection, ByVal code As String) As Variant
    On Error Resume Next
    UseOf = Null
    UseOf = useC(code)
End Function

' One use rate for a list of codes, or Null if they differ or one has none.
Private Function OneUse(ByVal codes As String, useC As Collection, txt As String) As Variant
    Dim c As Variant, v As Variant, first As Boolean, val1 As Variant, mixed As Boolean
    first = True
    val1 = Null
    For Each c In Split(codes, ", ")
        v = UseOf(useC, CStr(c))
        If first Then
            val1 = v
            first = False
        ElseIf IsNull(v) Or IsNull(val1) Then
            If Not (IsNull(v) And IsNull(val1)) Then mixed = True
        ElseIf Abs(v - val1) > 0.0000001 Then
            mixed = True
        End If
    Next
    If mixed Then
        OneUse = Null
        txt = "mixed"
    ElseIf IsNull(val1) Then
        OneUse = Null
        txt = "no rate"
    Else
        OneUse = val1
        txt = RateText(val1)
    End If
End Function

Private Sub AddMove(rsOut As DAO.Recordset, ByVal zip As String, ByVal streetName As String, _
        ByVal cityName As String, ByVal side As String, ByVal lo As String, ByVal hi As String, _
        ByVal oldCodes As String, ByVal newCodes As String, useL As Collection, useN As Collection)
    Dim txt As String, v As Variant, houses As String
    rsOut.AddNew
    rsOut!County = Mid$(newCodes, 7, 3)
    rsOut!CountyName = CountyName(Mid$(newCodes, 7, 3))
    rsOut!city = Left$(cityName, 40)
    rsOut!zip = zip
    rsOut!street = Left$(streetName, 80)
    rsOut!side = side
    rsOut!loTxt = lo
    rsOut!hiTxt = hi
    rsOut!LoNum = CLng(Val(lo))
    If lo = hi Then houses = lo Else houses = lo & " to " & hi
    If lo <> hi Or Not IsDigits(lo) Then
        Select Case side
            Case "O": houses = houses & ", odd side"
            Case "E": houses = houses & ", even side"
            Case Else: houses = houses & ", both sides"
        End Select
    End If
    rsOut!houses = Left$(houses, 80)
    rsOut!oldCodes = Left$(oldCodes, 120)
    rsOut!newCodes = Left$(newCodes, 120)
    v = OneUse(oldCodes, useL, txt)
    rsOut!OldUse = v
    rsOut!OldRate = txt
    v = OneUse(newCodes, useN, txt)
    rsOut!NewUse = v
    rsOut!NewRate = txt
    rsOut.Update
End Sub

' The most common value, ties going to the alphabetically first.
Private Function MostCommon(arr() As String, ByVal n As Long) As String
    Dim i As Long, j As Long, c As Long, best As Long, pick As String
    For i = 0 To n - 1
        c = 0
        For j = 0 To n - 1
            If arr(j) = arr(i) Then c = c + 1
        Next
        If c > best Or (c = best And StrComp(arr(i), pick, vbBinaryCompare) < 0) Then
            best = c
            pick = arr(i)
        End If
    Next
    MostCommon = pick
End Function

'============================================ step two: the ship-to list

' Loads the ship-to list (.xlsx or .csv) into the ShipTos table. The list is
' kept: a later list adds new ship-tos and updates the ones it has, so a later
' list only needs the new ones. A ship-to whose address changed loses its typed map code.
Public Function LoadShipToFile(ByVal path As String) As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset, v As Variant, msg As String, head As Long
    Dim r As Long, c As Long, i As Long, hk As String, col(1 To 14) As Long, keys As Variant
    Dim acct As String, shipTo As String, added As Long, changed As Long, same As Long, missing As String
    Dim a1 As String, a2 As String, city As String, moved As Long, before As String, after As String
    Dim item As Variant

    If LCase$(Right$(path, 4)) = ".csv" Then v = ReadCsvGrid(path, msg) Else v = ReadSheet(path, msg)
    If IsEmpty(v) Then Say msg, vbExclamation: Exit Function

    ' the header row: the first of the top ten rows naming City and Address 1
    For r = 1 To UBound(v, 1)
        If r > 10 Then Exit For
        hk = ""
        For c = 1 To UBound(v, 2)
            hk = hk & "|" & HeaderKey(CellText(v(r, c)))
        Next
        If InStr(hk & "|", "|city|") > 0 And InStr(hk & "|", "|address1|") > 0 Then head = r: Exit For
    Next
    If head = 0 Then
        Say "Couldn't find the column names in that file. It needs at least Acct #, Ship-To #, Address 1, " & _
            "Address 2 and City. Nothing changed.", vbExclamation
        Exit Function
    End If
    keys = Array("acct", "shipto", "customername", "shiptoname", "address1", "address2", "city", "state", _
        "county", "tax", "taxexemption", "producttypes", "customerstatus", "shiptostatus")
    For c = 1 To UBound(v, 2)
        hk = HeaderKey(CellText(v(head, c)))
        For i = 0 To 13
            If hk = keys(i) And col(i + 1) = 0 Then col(i + 1) = c
        Next
    Next
    For Each item In Array(Array(1, "Acct #"), Array(2, "Ship-To #"), Array(5, "Address 1"), Array(6, "Address 2"), _
            Array(7, "City"))
        If col(item(0)) = 0 Then missing = missing & ", " & item(1)
    Next
    If missing <> "" Then
        Say "That list is missing " & Mid$(missing, 3) & ". Nothing changed.", vbExclamation
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset("ShipTos", dbOpenTable)
    rs.Index = "PrimaryKey"
    For r = head + 1 To UBound(v, 1)
        acct = ListCell(v, r, col(1))
        shipTo = ListCell(v, r, col(2))
        If acct <> "" Or shipTo <> "" Then
            a1 = ListCell(v, r, col(5))
            a2 = ListCell(v, r, col(6))
            city = ListCell(v, r, col(7))
            rs.Seek "=", acct, shipTo
            If rs.NoMatch Then
                rs.AddNew
                rs!AcctNo = acct
                rs!ShipToNo = shipTo
                rs!FirstSeen = Now
                added = added + 1
            Else
                before = UCase$(Nz(rs!Address1, "") & "|" & Nz(rs!Address2, "") & "|" & Nz(rs!city, ""))
                after = UCase$(a1 & "|" & a2 & "|" & city)
                rs.Edit
                If before <> after Then
                    changed = changed + 1
                    If Nz(rs!PinnedCode, "") <> "" Then
                        rs!PinnedCode = Null
                        rs!PinnedOn = Null
                        moved = moved + 1
                    End If
                Else
                    same = same + 1
                End If
            End If
            rs!CustomerName = Left$(ListCell(v, r, col(3)), 120)
            rs!ShipToName = Left$(ListCell(v, r, col(4)), 120)
            rs!Address1 = Left$(a1, 150)
            rs!Address2 = Left$(a2, 150)
            rs!city = Left$(city, 60)
            rs!State = Left$(ListCell(v, r, col(8)), 20)
            rs!County = Left$(ListCell(v, r, col(9)), 60)
            If col(10) > 0 Then rs!Tax = TaxValue(v(r, col(10))) Else rs!Tax = Null
            rs!TaxExemption = Left$(ListCell(v, r, col(11)), 60)
            rs!ProductTypes = Left$(ListCell(v, r, col(12)), 60)
            rs!CustomerStatus = Left$(ListCell(v, r, col(13)), 20)
            rs!ShipToStatus = Left$(ListCell(v, r, col(14)), 20)
            rs!LastSeen = Now
            rs.Update
        End If
    Next
    rs.Close
    msg = "Loaded the ship-to list: " & Plural(added, "new ship-to", "new ship-tos") & ", " & _
        Plural(changed, "with a changed address", "with a changed address") & ", " & _
        Plural(same, "already here and the same", "already here and the same") & "."
    If moved > 0 Then msg = msg & " " & Plural(moved, "ship-to's", "ship-tos'") & " typed map code was cleared " & _
        "because the address changed."
    If col(10) = 0 Then msg = msg & " The list has no Tax column, so rates can't be compared."
    SetOpt "ListLoadedOn", Format$(Now, "yyyy-mm-dd hh:nn")
    SetOpt "ListFile", FileName(path)
    gLastMessage = msg
    LoadShipToFile = True
End Function

' A cell from the list as text: trimmed, NULL means empty.
Private Function ListCell(v As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim s As String
    If c = 0 Then Exit Function
    s = Trim$(CellText(v(r, c)))
    If UCase$(s) = "NULL" Then s = ""
    ListCell = s
End Function

' "Ship-To #" -> "shipto"
Private Function HeaderKey(ByVal h As String) As String
    Dim i As Long, ch As String, out As String
    h = LCase$(h)
    For i = 1 To Len(h)
        ch = Mid$(h, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next
    HeaderKey = out
End Function

' 0.08475, "0.08475", "8.475%" or 8.475 -> 0.08475. Empty or NULL -> Null.
Private Function TaxValue(ByVal v As Variant) As Variant
    Dim s As String, x As Double, pct As Boolean
    TaxValue = Null
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(v)
        If s = "" Or UCase$(s) = "NULL" Then Exit Function
        If Right$(s, 1) = "%" Then pct = True: s = Trim$(Left$(s, Len(s) - 1))
        If Not IsNumeric(s) Then Exit Function
        x = Val(s)
        If pct Then x = x / 100
    Else
        x = CDbl(v)
    End If
    If x > 0.5 Then x = x / 100
    TaxValue = Round(x, 6)
End Function

' A CSV file as a grid like ReadSheet's (text kept as written).
Private Function ReadCsvGrid(ByVal path As String, ByRef msg As String) As Variant
    Dim f As Integer, s As String, lines() As String, n As Long, r As Long, c As Long, mx As Long
    Dim parts() As String, grid() As Variant
    On Error GoTo Fail
    f = FreeFile
    Open path For Input Access Read Shared As #f
    ReDim lines(0 To 255)
    Do While Not EOF(f)
        Line Input #f, s
        If n = 0 Then s = NoMark(s)
        If n > UBound(lines) Then ReDim Preserve lines(0 To 2 * n)
        lines(n) = s
        n = n + 1
    Loop
    Close #f
    If n = 0 Then msg = "That file is empty. Nothing changed.": ReadCsvGrid = Empty: Exit Function
    For r = 0 To n - 1
        parts = SplitQuoted(lines(r))
        If UBound(parts) + 1 > mx Then mx = UBound(parts) + 1
    Next
    ReDim grid(1 To n, 1 To mx)
    For r = 0 To n - 1
        parts = SplitQuoted(lines(r))
        For c = 0 To UBound(parts)
            grid(r + 1, c + 1) = parts(c)
        Next
    Next
    ReadCsvGrid = grid
    Exit Function
Fail:
    msg = "Couldn't read " & FileName(path) & ": " & Err.Description & ". Nothing changed."
    ReadCsvGrid = Empty
End Function

' Matches every ship-to against this quarter and last quarter and sorts each
' into Fix (the rate in the system isn't the state's), Map (couldn't be placed
' for sure; look it up and type the code), OK, or Skipped (inactive).
Public Function CheckShipTos() As Boolean
    Dim db As DAO.Database, rs As DAO.Recordset, nowQ As String, lastQ As String, hasLast As Boolean
    Dim dbN As DAO.Database, dbL As DAO.Database
    Dim cityN As Cursor, countyN As Cursor, cityL As Cursor, countyL As Cursor
    Dim useN As Collection, salesN As Collection, useL As Collection, total As Long, done As Long
    Dim res As String, why As String, codeNow As String, howNow As String, whereNow As String, lineNow As String
    Dim useNow As Variant, codeLast As String, useLast As Variant

    Set db = CurrentDb
    FindQuarters nowQ, lastQ
    If nowQ = "" Then Exit Function
    total = DCount("*", "ShipTos")
    If total = 0 Then Exit Function
    hasLast = (lastQ <> "")
    InitWords
    LoadCountyNames
    Set dbN = DBEngine.OpenDatabase(QuarterFile(nowQ), False, True)
    Set useN = RateMap(dbN)
    Set salesN = SalesMap(dbN)
    OpenCursor cityN, dbN, "idxCity", "CityKey"
    OpenCursor countyN, dbN, "idxCounty", "County"
    If hasLast Then
        Set dbL = DBEngine.OpenDatabase(QuarterFile(lastQ), False, True)
        Set useL = RateMap(dbL)
        OpenCursor cityL, dbL, "idxCity", "CityKey"
        OpenCursor countyL, dbL, "idxCounty", "County"
    End If

    SysCmd acSysCmdInitMeter, "Checking the ship-tos", 100
    Set rs = db.OpenRecordset("ShipTos", dbOpenTable)
    Do While Not rs.EOF
        CheckOne rs, cityN, countyN, cityL, countyL, hasLast, useN, salesN, useL, DateText2(nowQ), _
            res, why, codeNow, howNow, whereNow, lineNow, useNow, codeLast, useLast
        rs.Edit
        rs!Result = res
        rs!why = Left$(why, 255)
        rs!codeNow = codeNow
        rs!howNow = howNow
        rs!whereNow = whereNow
        rs!lineNow = lineNow
        rs!useNow = useNow
        rs!codeLast = codeLast
        rs!useLast = useLast
        rs!Exempt = IIf(Nz(rs!TaxExemption, "") <> "", "Yes", "")
        rs!CheckedOn = Now
        If Nz(rs!PinnedCode, "") <> "" And IsNull(rs!PinnedOn) Then rs!PinnedOn = Now
        rs.Update
        done = done + 1
        If done Mod 200 = 0 Then
            SysCmd acSysCmdUpdateMeter, Int(100# * done / total)
            DoEvents
        End If
        rs.MoveNext
    Loop
    rs.Close
    cityN.rs.Close: countyN.rs.Close
    dbN.Close
    If hasLast Then cityL.rs.Close: countyL.rs.Close: dbL.Close
    SysCmd acSysCmdRemoveMeter

    ' this quarter's answers go in the history
    db.Execute "DELETE FROM Checks WHERE Quarter='" & nowQ & "'", dbFailOnError
    db.Execute "INSERT INTO Checks (Quarter, CheckedOn, AcctNo, ShipToNo, Code, How, UseRate, Tax, Result, Why) " & _
        "SELECT '" & nowQ & "', CheckedOn, AcctNo, ShipToNo, CodeNow, HowNow, UseNow, Tax, Result, Why " & _
        "FROM ShipTos WHERE Result<>'Skipped'", dbFailOnError
    SetOpt "ShipToCount", CStr(total)
    SetOpt "ShipToSkipped", CStr(DCount("*", "ShipTos", "Result='Skipped'"))
    SetOpt "ShipToFix", CStr(DCount("*", "ShipTos", "Result='Fix'"))
    SetOpt "ShipToMap", CStr(DCount("*", "ShipTos", "Result='Map'"))
    SetOpt "ShipToOK", CStr(DCount("*", "ShipTos", "Result='OK'"))
    SetOpt "ShipToByCounty", CStr(DCount("*", "ShipTos", "WhereNow='county' AND (Result='OK' OR Result='Fix')"))
    SetOpt "ShipToCheckedFor", nowQ
    CheckShipTos = True
End Function

Private Function SalesMap(db As DAO.Database) As Collection
    Dim rs As DAO.Recordset, c As New Collection
    Set rs = db.OpenRecordset("SELECT JCode, SalesRate FROM Rates", dbOpenSnapshot)
    Do While Not rs.EOF
        If Not IsNull(rs!SalesRate) Then c.Add CDbl(rs!SalesRate), CStr(rs!JCode)
        rs.MoveNext
    Loop
    rs.Close
    Set SalesMap = c
End Function

' "20261001" -> "Oct 1"
Private Function DateText2(ByVal q As String) As String
    If Len(q) < 8 Then Exit Function
    DateText2 = Format$(DateSerial(Val(Left$(q, 4)), Val(Mid$(q, 5, 2)), Val(Mid$(q, 7, 2))), "mmm d")
End Function

Private Sub CheckOne(rs As DAO.Recordset, cityN As Cursor, countyN As Cursor, _
        cityL As Cursor, countyL As Cursor, ByVal hasLast As Boolean, useN As Collection, _
        salesN As Collection, useL As Collection, ByVal qStart As String, res As String, why As String, _
        codeNow As String, howNow As String, whereNow As String, lineNow As String, useNow As Variant, _
        codeLast As String, useLast As Variant)
    Dim a1 As String, a2 As String, city As String, county As String, pinned As String, tax As Variant
    Dim how As String, code As String, note As String, where As String, ln As String, exempt As Boolean
    Dim salesNow As Variant, h2 As String, c2 As String, n2 As String, w2 As String, l2 As String

    res = "": why = "": codeNow = "": howNow = "": whereNow = "": lineNow = ""
    useNow = Null: codeLast = "": useLast = Null
    exempt = (Nz(rs!TaxExemption, "") <> "")
    If UCase$(Nz(rs!CustomerStatus, "")) = "INACTIVE" Or UCase$(Nz(rs!ShipToStatus, "")) = "INACTIVE" Then
        res = "Skipped": why = "inactive"
        Exit Sub
    End If
    If Nz(rs!State, "") <> "" And UCase$(Nz(rs!State, "")) <> "MO" And UCase$(Nz(rs!State, "")) <> "MISSOURI" Then
        res = "Skipped": why = "not a Missouri address"
        Exit Sub
    End If
    a1 = Nz(rs!Address1, ""): a2 = Nz(rs!Address2, ""): city = Nz(rs!city, ""): county = Nz(rs!County, "")
    pinned = Nz(rs!PinnedCode, "")
    tax = rs!Tax

    CombineLines cityN, countyN, a2, a1, city, county, how, code, note, where, ln
    howNow = how: whereNow = where: lineNow = ln
    If pinned <> "" Then
        If IsPlaced(how) And code <> pinned Then
            res = "Map": codeNow = code
            why = "you typed " & pinned & ", but the state's file now puts it in " & code
            Exit Sub
        End If
        code = pinned: howNow = "PINNED": lineNow = "typed"
    End If
    If code = "" Then
        res = "Map": why = note
        Exit Sub
    End If
    codeNow = code
    useNow = UseOf(useN, code)
    If IsNull(useNow) Then
        res = "Map": why = "the state's rate file has no use rate for " & code
        Exit Sub
    End If
    salesNow = UseOf(salesN, code)
    If hasLast Then
        If pinned <> "" Then
            codeLast = pinned
        Else
            CombineLines cityL, countyL, a2, a1, city, county, h2, c2, n2, w2, l2
            If IsPlaced(h2) Then codeLast = c2
        End If
        If codeLast <> "" Then useLast = UseOf(useL, codeLast)
    End If

    If IsNull(tax) Then
        If exempt Then
            res = "OK": why = "exempt, no rate set up"
        Else
            res = "Fix": why = "no rate set up"
        End If
    ElseIf Abs(tax - useNow) < 0.000001 Then
        res = "OK": why = ""
    ElseIf Not IsNull(salesNow) And Abs(tax - Nz(salesNow, 0)) < 0.000001 Then
        res = "Fix": why = "has the sales rate, not the use rate"
    ElseIf Not IsNull(useLast) And Abs(tax - Nz(useLast, 0)) < 0.000001 Then
        If codeLast <> "" And codeLast <> code Then
            res = "Fix": why = "moved from " & codeLast & " to " & code & " on " & qStart
        Else
            res = "Fix": why = "new rate on " & qStart
        End If
    Else
        res = "Fix": why = "doesn't match the state's rate"
    End If
End Sub

' Placed means street-exact. A street-only guess (the house number isn't in
' the state's ranges) goes to the map, its likely code in the note.
Private Function IsPlaced(ByVal how As String) As Boolean
    IsPlaced = (how = "EXACT" Or how = "CLOSE")
End Function

' Both address lines, Address 2 first (it usually holds the street, but
' people use either). One places: that's it. Both place on the same code:
' fine. Both place, different codes, or neither places: the map.
Private Sub CombineLines(rsCity As Cursor, rsCounty As Cursor, ByVal a2 As String, _
        ByVal a1 As String, ByVal city As String, ByVal county As String, how As String, code As String, _
        note As String, where As String, lineName As String)
    Dim h(0 To 1) As String, c(0 To 1) As String, n(0 To 1) As String, w(0 To 1) As String, ln(0 To 1) As String
    Dim lines(0 To 1) As String, names(0 To 1) As String, i As Long, nt As Long, p1 As Long, np As Long, pick As Long
    lines(0) = a2: names(0) = "Address 2"
    lines(1) = a1: names(1) = "Address 1"
    For i = 0 To 1
        If Len(Trim$(lines(i))) > 0 Then
            MatchShipTo rsCity, rsCounty, lines(i), city, county, h(nt), c(nt), n(nt), w(nt)
            ln(nt) = names(i)
            nt = nt + 1
        End If
    Next
    how = "NONE": code = "": note = "": where = "": lineName = ""
    p1 = -1
    For i = 0 To nt - 1
        If IsPlaced(h(i)) Then
            np = np + 1
            If p1 = -1 Then p1 = i
        End If
    Next
    If np = 2 Then
        If c(0) = c(1) Then
            how = h(0): code = c(0): note = n(0): where = w(0): lineName = "both"
        Else
            how = "MIXED": note = "Address 2 is in " & c(0) & " but Address 1 is in " & c(1)
        End If
        Exit Sub
    End If
    If np = 1 Then
        how = h(p1): code = c(p1): note = n(p1): where = w(p1): lineName = ln(p1)
        Exit Sub
    End If
    If nt = 0 Then
        note = "no address"
        Exit Sub
    End If
    pick = 0
    For i = 0 To nt - 1
        If h(i) = "MIXED" Then pick = i: Exit For
    Next
    how = h(pick): note = n(pick): where = w(pick)
End Sub

' One line, plus the county cross-check: if the list's county names a
' different county than the state's row, it goes to the map.
Private Sub MatchShipTo(rsCity As Cursor, rsCounty As Cursor, ByVal addr As String, _
        ByVal city As String, ByVal county As String, how As String, code As String, note As String, where As String)
    MatchLine rsCity, rsCounty, addr, city, county, how, code, note, where
    If code <> "" And county <> "" Then
        If Not CountyAgrees(county, Mid$(code, 7, 3)) Then
            note = "the list says " & county & ", but the state's row is in " & CountyName(Mid$(code, 7, 3)) & _
                " (" & code & ")"
            how = "MIXED"
            code = ""
        End If
    End If
End Sub

' One address line against one quarter: the street within the city on the
' mail (the state's City Name is the postal city), then within the county.
' how: EXACT (as written, number in range), CLOSE (a word left off or spelled
' another way), STREET (number not in the state's ranges, but the whole street
' is one code), MIXED (the state's rows disagree), NONE.
Private Sub MatchLine(rsCity As Cursor, rsCounty As Cursor, ByVal addr As String, _
        ByVal city As String, ByVal county As String, how As String, code As String, note As String, where As String)
    Dim numS As String, numV As Double, hasNum As Boolean, toks As String
    Dim rdPre(0 To 7) As String, rdName(0 To 7) As String, rdSuf(0 To 7) As String, rdPost(0 To 7) As String
    Dim nRd As Long, names(0 To 7) As String, nNames As Long, i As Long, k As Long
    Dim rows As MatchRows, ck As String, cc As String, how2 As String, code2 As String, note2 As String

    how = "NONE": code = "": note = "": where = ""
    SplitStreet addr, numS, numV, hasNum, toks
    If toks = "" Then note = "no street on that line": Exit Sub
    Readings toks, rdPre, rdName, rdSuf, rdPost, nRd
    For i = 0 To nRd - 1
        For k = 0 To nNames - 1
            If names(k) = rdName(i) Then Exit For
        Next
        If k = nNames Then names(nNames) = rdName(i): nNames = nNames + 1
    Next

    ck = CityKeyOf(city)
    InitMatchRows rows
    If ck <> "" Then
        For i = 0 To nNames - 1
            ReadNamed rsCity, ck, names(i), rows
        Next
    End If
    MatchSteps rows, numS, numV, hasNum, toks, rdPre, rdName, rdSuf, rdPost, nRd, how, code, note
    If how = "EXACT" Or how = "CLOSE" Or how = "MIXED" Then where = "city": Exit Sub

    cc = CountyCodeOf(county)
    If cc <> "" Then
        rows.n = 0
        For i = 0 To nNames - 1
            ReadNamed rsCounty, cc, names(i), rows
        Next
        MatchSteps rows, numS, numV, hasNum, toks, rdPre, rdName, rdSuf, rdPost, nRd, how2, code2, note2
        If how = "STREET" Then
            ' the house number isn't on that street under that city: a house the
            ' file doesn't have yet, or a city name that's another town's. The
            ' county has to agree.
            If how2 = "NONE" Then where = "city": Exit Sub
            If code2 = code Then
                If how2 = "EXACT" Or how2 = "CLOSE" Then
                    how = how2: note = note2: where = "county"
                Else
                    where = "city"
                End If
                Exit Sub
            End If
            note = "that street under that city is " & code & ", but the county's rows say " & _
                IIf(code2 = "", "more than one code", code2)
            how = "MIXED": code = "": where = "county"
            Exit Sub
        End If
        If how2 <> "NONE" Then
            how = how2: code = code2: note = note2: where = "county"
            Exit Sub
        End If
    End If
    If how = "STREET" Then where = "city"
End Sub

' '1234 N Main St Ste 5' -> number '1234', words 'N MAIN ST'.
Private Sub SplitStreet(ByVal addr As String, numS As String, numV As Double, hasNum As Boolean, toks As String)
    Dim t As String, p As Long, first As String, digits As String, i As Long, ch As String
    numS = "": numV = 0: hasNum = False
    t = Tokens(addr)
    If t <> "" Then
        If IsDigits(Left$(t, 1)) Then
            p = InStr(1, t, " ")
            If p > 0 Then first = Left$(t, p - 1): t = Mid$(t, p + 1) Else first = t: t = ""
            For i = 1 To Len(first)
                ch = Mid$(first, i, 1)
                If ch < "0" Or ch > "9" Then Exit For
                digits = digits & ch
            Next
            numS = first
            numV = CDbl(digits)
            hasNum = True
            ' '123 1/2 Main': the half became the words 1 and 2
            If Left$(t & " ", 4) = "1 2 " Then t = Mid$(t, 5)
        End If
    End If
    toks = Canon(CutUnit(t))
End Sub

' Every sensible way to split the words into (direction before, name, suffix,
' direction after). 'Henley Lake' may be a name or 'Henley' plus Lake.
Private Sub Readings(ByVal toks As String, rdPre() As String, rdName() As String, rdSuf() As String, _
        rdPost() As String, nRd As Long)
    Dim t() As String, n As Long, up As Long, uo As Long, us As Long, a As Long, z As Long, k As Long
    Dim pre As String, post As String, suf As String, nm As String, ok As Boolean
    t = Split(toks, " ")
    n = UBound(t) + 1
    nRd = 0
    For up = 1 To 0 Step -1
        For uo = 1 To 0 Step -1
            For us = 1 To 0 Step -1
                a = 0: z = n: pre = "": post = "": suf = "": ok = True
                If up = 1 Then
                    If z - a > 1 And InSet(mDir, t(a)) Then pre = t(a): a = a + 1 Else ok = False
                End If
                If ok And uo = 1 Then
                    If z - a > 1 And InSet(mDir, t(z - 1)) Then post = t(z - 1): z = z - 1 Else ok = False
                End If
                If ok And us = 1 Then
                    ok = False
                    If z - a > 1 Then
                        If InSet(mSuffix, t(z - 1)) Then
                            If Not (z - a = 2 And (t(a) = "HWY" Or t(a) = "CR" Or t(a) = "FR" Or t(a) = "I")) Then
                                suf = t(z - 1): z = z - 1: ok = True
                            End If
                        End If
                    End If
                End If
                If ok Then
                    nm = ""
                    For k = a To z - 1
                        nm = nm & " " & t(k)
                    Next
                    rdPre(nRd) = pre: rdName(nRd) = Mid$(nm, 2): rdSuf(nRd) = suf: rdPost(nRd) = post
                    nRd = nRd + 1
                End If
            Next
        Next
    Next
End Sub

Private Sub InitMatchRows(rows As MatchRows)
    rows.n = 0
    ReDim rows.num(0 To 31): ReDim rows.lo(0 To 31): ReDim rows.hi(0 To 31)
    ReDim rows.loTxt(0 To 31): ReDim rows.hiTxt(0 To 31): ReDim rows.side(0 To 31)
    ReDim rows.pre(0 To 31): ReDim rows.nm(0 To 31): ReDim rows.suf(0 To 31): ReDim rows.post(0 To 31)
    ReDim rows.full(0 To 31): ReDim rows.code(0 To 31): ReDim rows.street(0 To 31): ReDim rows.city(0 To 31)
End Sub

Private Sub GrowMatchRows(rows As MatchRows)
    Dim m As Long
    m = 2 * (UBound(rows.lo) + 1) - 1
    ReDim Preserve rows.num(0 To m): ReDim Preserve rows.lo(0 To m): ReDim Preserve rows.hi(0 To m)
    ReDim Preserve rows.loTxt(0 To m): ReDim Preserve rows.hiTxt(0 To m): ReDim Preserve rows.side(0 To m)
    ReDim Preserve rows.pre(0 To m): ReDim Preserve rows.nm(0 To m): ReDim Preserve rows.suf(0 To m)
    ReDim Preserve rows.post(0 To m): ReDim Preserve rows.full(0 To m): ReDim Preserve rows.code(0 To m)
    ReDim Preserve rows.street(0 To m): ReDim Preserve rows.city(0 To m)
End Sub

' Adds the rows filed under (key, street name) to the list.
Private Sub ReadNamed(cur As Cursor, ByVal key As String, ByVal nm As String, rows As MatchRows)
    cur.rs.Seek "=", key, nm
    If cur.rs.NoMatch Then Exit Sub
    Do While Not cur.rs.EOF
        If cur.key.Value & "" <> key Or cur.nm.Value & "" <> nm Then Exit Do
        If rows.n > UBound(rows.lo) Then GrowMatchRows rows
        If IsNull(cur.lo.Value) Then
            rows.num(rows.n) = False
            rows.loTxt(rows.n) = cur.loTxt.Value & ""
            rows.hiTxt(rows.n) = cur.hiTxt.Value & ""
        Else
            rows.num(rows.n) = True
            rows.lo(rows.n) = cur.lo.Value
            rows.hi(rows.n) = cur.hi.Value
        End If
        If IsNull(cur.side.Value) Then rows.side(rows.n) = "B" Else rows.side(rows.n) = cur.side.Value
        rows.pre(rows.n) = cur.pre.Value & ""
        rows.nm(rows.n) = nm
        rows.suf(rows.n) = cur.suf.Value & ""
        rows.post(rows.n) = cur.post.Value & ""
        rows.full(rows.n) = cur.full.Value & ""
        rows.code(rows.n) = cur.code.Value & ""
        rows.street(rows.n) = cur.street.Value & ""
        rows.city(rows.n) = cur.city.Value & ""
        rows.n = rows.n + 1
        cur.rs.MoveNext
    Loop
End Sub

' Opens the Boundary table of one quarter's file on an index, with its fields
' at hand. keyField is the index's first field.
Private Sub OpenCursor(cur As Cursor, db As DAO.Database, ByVal idx As String, ByVal keyField As String)
    Set cur.rs = db.OpenRecordset("Boundary", dbOpenTable)
    cur.rs.Index = idx
    With cur.rs.Fields
        Set cur.key = .Item(keyField)
        Set cur.zip = .Item("Zip")
        Set cur.full = .Item("FullKey")
        Set cur.nm = .Item("NameKey")
        Set cur.pre = .Item("PreKey")
        Set cur.suf = .Item("SufKey")
        Set cur.post = .Item("PostKey")
        Set cur.lo = .Item("Lo")
        Set cur.hi = .Item("Hi")
        Set cur.loTxt = .Item("LoTxt")
        Set cur.hiTxt = .Item("HiTxt")
        Set cur.side = .Item("Side")
        Set cur.code = .Item("JCode")
        Set cur.street = .Item("Street")
        Set cur.city = .Item("City")
    End With
End Sub

' The street steps: as written, then close, then the whole street.
Private Sub MatchSteps(rows As MatchRows, ByVal numS As String, ByVal numV As Double, ByVal hasNum As Boolean, _
        ByVal toks As String, rdPre() As String, rdName() As String, rdSuf() As String, rdPost() As String, _
        ByVal nRd As Long, how As String, code As String, note As String)
    Dim hit() As Boolean, onStreet() As Boolean, i As Long, k As Long, codes() As String, nc As Long, first As Long
    Dim nOn As Long
    how = "NONE": code = "": note = ""
    If rows.n = 0 Then note = "street not found there": Exit Sub
    ReDim hit(0 To rows.n - 1): ReDim onStreet(0 To rows.n - 1): ReDim codes(0 To rows.n)

    ' 1. the street as written
    For i = 0 To rows.n - 1
        hit(i) = (rows.full(i) = toks) And RowHas(rows, i, numS, numV, hasNum)
    Next
    Narrowest rows, hit
    nc = 0: first = -1
    For i = 0 To rows.n - 1
        If hit(i) Then
            AddUnique codes, nc, rows.code(i)
            If first = -1 Then first = i
        End If
    Next
    If nc = 1 Then how = "EXACT": code = codes(0): note = SayRow(rows, first): Exit Sub
    If nc > 1 Then
        SortStrings codes, nc
        how = "MIXED": note = "the state's file has more than one code here: " & JoinN(codes, nc, ", ")
        Exit Sub
    End If

    ' 2. close: a reading of the words fits the row
    For i = 0 To rows.n - 1
        onStreet(i) = False
        For k = 0 To nRd - 1
            If Fits(rows, i, rdPre(k), rdName(k), rdSuf(k), rdPost(k)) Then onStreet(i) = True: Exit For
        Next
        If onStreet(i) Then nOn = nOn + 1
        hit(i) = onStreet(i) And RowHas(rows, i, numS, numV, hasNum)
    Next
    Narrowest rows, hit
    nc = 0: first = -1
    For i = 0 To rows.n - 1
        If hit(i) Then
            AddUnique codes, nc, rows.code(i)
            If first = -1 Then first = i
        End If
    Next
    If nc = 1 Then how = "CLOSE": code = codes(0): note = SayRow(rows, first): Exit Sub
    If nc > 1 Then
        SortStrings codes, nc
        how = "MIXED": note = "more than one street fits what was written: " & JoinN(codes, nc, ", ")
        Exit Sub
    End If

    ' 3. the right street, a number the state doesn't list, one code for the whole street
    nc = 0
    For i = 0 To rows.n - 1
        If onStreet(i) Then AddUnique codes, nc, rows.code(i)
    Next
    If nc = 1 And hasNum Then
        how = "STREET": code = codes(0)
        note = "that house number isn't in the state's file; the rest of the street is " & code
        Exit Sub
    End If
    If nOn > 0 And Not hasNum Then
        note = "no house number on that line"
    ElseIf nOn > 0 Then
        note = "the street is there, but not that house number"
    Else
        note = "street not found there"
    End If
End Sub

' Does the reading name the same street as the row, with nothing it says
' contradicting the row (a direction or suffix left out is fine)?
Private Function Fits(rows As MatchRows, ByVal i As Long, ByVal pre As String, ByVal nm As String, _
        ByVal suf As String, ByVal post As String) As Boolean
    If nm <> rows.nm(i) Then Exit Function
    If suf <> "" And suf <> rows.suf(i) Then Exit Function
    If (pre <> "" Or post <> "") And (rows.pre(i) <> "" Or rows.post(i) <> "") Then
        If Not ((pre <> "" And (pre = rows.pre(i) Or pre = rows.post(i))) Or _
                (post <> "" And (post = rows.pre(i) Or post = rows.post(i)))) Then Exit Function
    End If
    Fits = True
End Function

' Is the house number in the row's range, on the right side of the street?
Private Function RowHas(rows As MatchRows, ByVal i As Long, ByVal numS As String, ByVal numV As Double, _
        ByVal hasNum As Boolean) As Boolean
    Dim odd As Boolean
    If Not hasNum Then Exit Function
    If rows.num(i) Then
        If numV < rows.lo(i) Or numV > rows.hi(i) Then Exit Function
        odd = (numV - 2 * Int(numV / 2) = 1)
        If rows.side(i) = "O" And Not odd Then Exit Function
        If rows.side(i) = "E" And odd Then Exit Function
        RowHas = True
    Else
        RowHas = (numS = rows.loTxt(i) Or numS = rows.hiTxt(i))
    End If
End Function

' On one street, the narrowest rows win (a carved-out house beats the range
' around it); two different streets never knock each other out.
Private Sub Narrowest(rows As MatchRows, hit() As Boolean)
    Dim i As Long, j As Long, keep() As Boolean
    ReDim keep(0 To rows.n - 1)
    For i = 0 To rows.n - 1
        If hit(i) Then
            keep(i) = True
            For j = 0 To rows.n - 1
                If hit(j) And j <> i Then
                    If rows.full(j) = rows.full(i) And RowWidth(rows, j) < RowWidth(rows, i) Then keep(i) = False: Exit For
                End If
            Next
        End If
    Next
    For i = 0 To rows.n - 1
        hit(i) = keep(i)
    Next
End Sub

Private Function RowWidth(rows As MatchRows, ByVal i As Long) As Long
    If rows.num(i) Then RowWidth = rows.hi(i) - rows.lo(i)
End Function

Private Function SayRow(rows As MatchRows, ByVal i As Long) As String
    If rows.num(i) Then
        SayRow = rows.lo(i) & "-" & rows.hi(i) & " " & rows.street(i) & ", " & rows.city(i)
    Else
        SayRow = rows.loTxt(i) & "-" & rows.hiTxt(i) & " " & rows.street(i) & ", " & rows.city(i)
    End If
End Function

' The list's county column -> the DOR's three-digit county code; "" if unknown.
Public Function CountyCodeOf(ByVal nm As String) As String
    Dim k As String, v As String
    k = CountyKeyOf(nm)
    If k = "" Then Exit Function
    If IsDigits(k) And Len(k) <= 3 Then CountyCodeOf = Right$("00" & k, 3): Exit Function
    If mCountyKeys Is Nothing Then LoadCountyKeys
    On Error Resume Next
    v = mCountyKeys(k)
    CountyCodeOf = v
End Function

Private Sub LoadCountyKeys()
    Dim rs As DAO.Recordset, k As String
    Set mCountyKeys = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT County, CountyName FROM Counties WHERE County<>'186'", dbOpenSnapshot)
    On Error Resume Next
    Do While Not rs.EOF
        k = CountyKeyOf(rs!CountyName)
        mCountyKeys.Add CStr(rs!County), k
        rs.MoveNext
    Loop
    mCountyKeys.Add "193", "STGENEVIEVE"
    mCountyKeys.Add "189", "STLOUISCO"
    rs.Close
End Sub

' 'St. Louis County', 'ST LOUIS CO', 'Saint Louis' -> 'STLOUIS'; 'St. Louis City' -> 'STLOUISCITY'.
Public Function CountyKeyOf(ByVal nm As String) As String
    Dim i As Long, c As Long, buf As String, pos As Long, parts() As String, out As String, w As String
    If Len(nm) = 0 Then Exit Function
    buf = Space$(Len(nm))
    For i = 1 To Len(nm)
        c = AscW(Mid$(nm, i, 1))
        If c = 46 Or c = 39 Then
        ElseIf c >= 97 And c <= 122 Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c - 32)
        ElseIf (c >= 65 And c <= 90) Or (c >= 48 And c <= 57) Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c)
        Else
            pos = pos + 1
        End If
    Next
    If pos = 0 Then Exit Function
    parts = Split(Left$(buf, pos), " ")
    For i = 0 To UBound(parts)
        w = parts(i)
        If w = "SAINT" Then w = "ST"
        If w = "SAINTE" Then w = "STE"
        If w <> "COUNTY" And w <> "CO" And w <> "CNTY" Then out = out & w
    Next
    CountyKeyOf = out
End Function

' Does the list's county fit the state's county code? A county name that
' can't be read, and a plain "St. Louis" (city or county), pass.
Private Function CountyAgrees(ByVal nm As String, ByVal code As String) As Boolean
    Dim cc As String
    cc = CountyCodeOf(nm)
    If cc = "" Then CountyAgrees = True: Exit Function
    If CountyKeyOf(nm) = "STLOUIS" And (code = "189" Or code = "510") Then CountyAgrees = True: Exit Function
    CountyAgrees = (cc = code)
End Function

'============================================== text for the form and report

Public Function ChangesTitle() As String
    Dim nowQ As String, lastQ As String
    FindQuarters nowQ, lastQ
    ChangesTitle = "Use tax changes, " & QLabel(lastQ) & " to " & QLabel(nowQ)
End Function

Public Function ChangesSummary() As String
    Dim s As String, nowQ As String, lastQ As String, places As Long, unused As Long, streets As Long, zips As Long
    ReadChanges nowQ, lastQ, places, unused, streets, zips
    s = "New rates start " & DateText(nowQ) & ". Only the use rate counts here; " & _
        "sales tax only changes are left out." & vbCrLf
    s = s & Plural(places, "place has", "places have") & " a new use rate"
    If unused > 0 Then s = s & " (" & unused & " more did, but no " & _
        "address in the state's file is in them, so they're left out)"
    s = s & ". On " & Plural(streets, "street", "streets") & " (in " & Plural(zips, "zip", "zips") & _
        "), some house numbers moved to a place with a different use rate." & vbCrLf
    s = s & "By county. Each street line shows the house numbers that moved and the city on their mail."
    ChangesSummary = s
End Function

Public Function CountyLine(ByVal places As Variant, ByVal streets As Variant) As String
    Dim s As String
    If Nz(places, 0) > 0 Then s = Plural(Nz(places, 0), "place with a new rate", "places with a new rate")
    If Nz(streets, 0) > 0 Then
        If s <> "" Then s = s & ", "
        s = s & Plural(Nz(streets, 0), "street", "streets")
    End If
    CountyLine = s
End Function

Public Function ChangeText(ByVal oldUse As Variant, ByVal newUse As Variant) As String
    If IsNull(oldUse) Or IsNull(newUse) Then Exit Function
    ChangeText = Format$((newUse - oldUse) * 100, "+0.000;-0.000;0.000") & "%"
End Function

Public Function RateText(ByVal v As Variant) As String
    If IsNull(v) Then
        RateText = "none"
    Else
        RateText = Format$(v * 100, "0.000") & "%"
    End If
End Function

Private Function Plural(ByVal n As Double, ByVal one As String, ByVal many As String) As String
    If n = 1 Then Plural = "1 " & one Else Plural = Format$(n, "#,##0") & " " & many
End Function

'=============================================================== quarters

' "20261001" -> "Q4 2026"
Public Function QLabel(ByVal q As String) As String
    If Len(q) < 8 Then Exit Function
    QLabel = "Q" & ((Val(Mid$(q, 5, 2)) - 1) \ 3 + 1) & " " & Left$(q, 4)
End Function

' "20261001" -> "Q4-2026", as in the DOR's file names
Private Function QFileTag(ByVal q As String) As String
    QFileTag = Replace(QLabel(q), " ", "-")
End Function

' "20261001" -> "Oct 1, 2026"
Private Function DateText(ByVal q As String) As String
    If Len(q) < 8 Then Exit Function
    DateText = Format$(DateSerial(Val(Left$(q, 4)), Val(Mid$(q, 5, 2)), Val(Mid$(q, 7, 2))), "mmm d, yyyy")
End Function

' "dor-jurisdiction-total-rate-q3-2026.xlsx" -> "20260701"; "" if the name doesn't say.
Private Function QuarterInName(ByVal path As String) As String
    Dim s As String, i As Long, qn As Long
    s = LCase$(FileName(path))
    For i = 1 To Len(s) - 6
        If Mid$(s, i, 1) = "q" And Mid$(s, i + 2, 1) = "-" Then
            qn = Val(Mid$(s, i + 1, 1))
            If qn >= 1 And qn <= 4 And IsDigits(Mid$(s, i + 3, 4)) Then
                QuarterInName = Mid$(s, i + 3, 4) & Format$(3 * qn - 2, "00") & "01"
                Exit Function
            End If
        End If
    Next
End Function

'================================================== the street spelling rules

' The same rules as the reference matcher: words to their standard spelling
' (AVENUE -> AVE), highways written one way (HWY 50), units dropped.
Private Sub InitWords()
    Dim s As String, p As Variant, kv() As String
    If Not mSyn Is Nothing Then Exit Sub
    Set mSyn = New Collection
    s = "NORTH N|SOUTH S|EAST E|WEST W|NORTHEAST NE|NORTHWEST NW|SOUTHEAST SE|SOUTHWEST SW|"
    s = s & "ALLEY ALY|ALLEE ALY|ALLY ALY|ANNEX ANX|ANEX ANX|ARCADE ARC|AVENUE AVE|AV AVE|AVEN AVE|"
    s = s & "AVENU AVE|AVN AVE|AVNUE AVE|BAYOU BYU|BEACH BCH|BEND BND|BLUFF BLF|BLUF BLF|BOTTOM BTM|"
    s = s & "BOULEVARD BLVD|BOUL BLVD|BOULV BLVD|BRANCH BR|BRIDGE BRG|BROOK BRK|BYPASS BYP|CAMP CP|"
    s = s & "CANYON CYN|CAPE CPE|CAUSEWAY CSWY|CENTER CTR|CENTRE CTR|CENT CTR|CNTR CTR|CNTER CTR|"
    s = s & "CIRCLE CIR|CIRC CIR|CIRCL CIR|CRCL CIR|CRCLE CIR|CLIFF CLF|CLUB CLB|COMMON CMN|CORNER COR|"
    s = s & "CORNERS CORS|COURSE CRSE|COURT CT|CRT CT|COVE CV|CREEK CRK|CRESCENT CRES|CREST CRST|"
    s = s & "CROSSING XING|CRSSNG XING|CROSSROAD XRD|CURVE CURV|DALE DL|DAM DM|DIVIDE DV|DRIVE DR|"
    s = s & "DRIV DR|DRV DR|ESTATE EST|ESTATES EST|ESTS EST|EXPRESSWAY EXPY|EXPR EXPY|EXPRESS EXPY|"
    s = s & "EXPW EXPY|EXTENSION EXT|EXTN EXT|EXTNSN EXT|FALLS FLS|FERRY FRY|FIELD FLD|FIELDS FLDS|"
    s = s & "FLAT FLT|FOREST FRST|FORGE FRG|FORK FRK|FORT FT|FREEWAY FWY|GARDEN GDN|GARDENS GDNS|"
    s = s & "GATEWAY GTWY|GLEN GLN|GREEN GRN|GROVE GRV|HARBOR HBR|HAVEN HVN|HEIGHTS HTS|HEIGHT HTS|"
    s = s & "HT HTS|HIGHWAY HWY|HIGHWY HWY|HIWAY HWY|HIWY HWY|HWAY HWY|HILL HL|HILLS HLS|HOLLOW HOLW|"
    s = s & "HOLLOWS HOLW|HLLW HOLW|INLET INLT|ISLAND IS|JUNCTION JCT|JCTION JCT|KNOLL KNL|LAKE LK|"
    s = s & "LAKES LKS|LANDING LNDG|LNDNG LNDG|LANE LN|MANOR MNR|MEADOW MDW|MEADOWS MDWS|MILL ML|"
    s = s & "MILLS MLS|MOUNT MT|MOUNTAIN MTN|ORCHARD ORCH|PARKS PARK|PARKWAY PKWY|PARKWY PKWY|"
    s = s & "PKWAY PKWY|PKY PKWY|PINE PNE|PINES PNES|PLACE PL|PLAIN PLN|PLAINS PLNS|PLAZA PLZ|PLZA PLZ|"
    s = s & "POINT PT|POINTS PTS|PORT PRT|PRAIRIE PR|RANCH RNCH|RIDGE RDG|RDGE RDG|RIVER RIV|ROAD RD|"
    s = s & "ROUTE RTE|RT RTE|SHORE SHR|SHORES SHRS|SPRING SPG|SPRINGS SPGS|SQUARE SQ|SQR SQ|SQU SQ|"
    s = s & "STATION STA|STREET ST|STR ST|STRT ST|SUMMIT SMT|TERRACE TER|TERR TER|TRACE TRCE|TRACK TRAK|"
    s = s & "TRAFFICWAY TRFY|TRAIL TRL|TRAILS TRL|TRLS TRL|TUNNEL TUNL|TURNPIKE TPKE|VALLEY VLY|VIEW VW|"
    s = s & "VILLAGE VLG|VILLE VL|VISTA VIS|WELL WL|WELLS WLS|SAINT ST|INTERSTATE I|IH I|COUNTY CO|"
    s = s & "BUSINESS BUS|PRIVATE PVT|MISSOURI MO"
    For Each p In Split(s, "|")
        kv = Split(p, " ")
        mSyn.Add kv(1), kv(0)
    Next
    Set mSuffix = New Collection
    s = "ALY|ANX|ARC|AVE|BCH|BLF|BLVD|BND|BR|BRG|BRK|BTM|BYP|BYU|CIR|CLB|CLF|CMN|COR|CORS|CP|CPE|CRES|CRK|"
    s = s & "CRSE|CRST|CSWY|CT|CTR|CURV|CV|CYN|DL|DM|DR|DV|EST|EXPY|EXT|FLD|FLDS|FLS|FLT|FRG|FRK|FRST|FRY|FT|FWY|"
    s = s & "GDN|GDNS|GLN|GRN|GRV|GTWY|HBR|HL|HLS|HOLW|HTS|HVN|HWY|INLT|IS|JCT|KNL|LK|LKS|LN|LNDG|LOOP|MALL|MDW|"
    s = s & "MDWS|ML|MLS|MNR|MT|MTN|ORCH|OVAL|PARK|PASS|PATH|PIKE|PKWY|PL|PLN|PLNS|PLZ|PNE|PNES|PR|PRT|PT|PTS|RD|"
    s = s & "RDG|RIV|RNCH|ROW|RTE|RUN|SHR|SHRS|SMT|SPG|SPGS|SPUR|SQ|ST|STA|TER|TPKE|TRAK|TRCE|TRFY|TRL|TUNL|VIS|"
    s = s & "VL|VLG|VLY|VW|WALK|WAY|WL|WLS|XING|XRD"
    For Each p In Split(s, "|")
        mSuffix.Add p, p
    Next
    Set mUnit = New Collection
    For Each p In Split("# APARTMENT APT BLDG BSMT BUILDING DEPT FL FLOOR FRNT LBBY LOT LOWR NO OFC OFFICE " & _
            "REAR RM ROOM SPACE SPC STE SUITE TRAILER TRLR UNIT UPPR", " ")
        mUnit.Add p, p
    Next
    Set mDir = New Collection
    For Each p In Split("N S E W NE NW SE SW", " ")
        mDir.Add p, p
    Next
End Sub

Private Function InSet(c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    If Len(k) = 0 Then Exit Function
    On Error Resume Next
    v = c(k)
    InSet = (Err.Number = 0)
End Function

Private Function Syn(ByVal w As String) As String
    On Error Resume Next
    Syn = w
    Syn = mSyn(w)
End Function

' Uppercase; drop periods and apostrophes; anything else that isn't a letter,
' digit or # becomes a space; each word to its standard spelling. Returns the
' words joined by single spaces.
Public Function Tokens(ByVal text As String) As String
    Dim i As Long, n As Long, c As Long, buf As String, pos As Long, parts() As String, out As String
    InitWords
    n = Len(text)
    If n = 0 Then Exit Function
    buf = Space$(3 * n + 1)
    For i = 1 To n
        c = AscW(Mid$(text, i, 1))
        If c = 46 Or c = 39 Then
            ' dropped
        ElseIf c >= 97 And c <= 122 Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c - 32)
        ElseIf (c >= 65 And c <= 90) Or (c >= 48 And c <= 57) Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c)
        ElseIf c = 35 Then
            Mid$(buf, pos + 1, 3) = " # "
            pos = pos + 3
        Else
            pos = pos + 1
        End If
    Next
    If pos = 0 Then Exit Function
    parts = Split(Left$(buf, pos), " ")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then out = out & " " & Syn(parts(i))
    Next
    Tokens = Mid$(out, 2)
End Function

' A highway id: a number (50, 13A) or one to three letters (K, CC, AB).
Private Function IsRouteId(ByVal w As String) As Boolean
    Dim i As Long, c As Long
    If Len(w) = 0 Then Exit Function
    c = AscW(w)
    If c >= 48 And c <= 57 Then
        For i = 1 To Len(w)
            c = AscW(Mid$(w, i, 1))
            If Not ((c >= 48 And c <= 57) Or (c >= 65 And c <= 90)) Then Exit Function
        Next
        IsRouteId = True
    Else
        If Len(w) > 3 Then Exit Function
        For i = 1 To Len(w)
            c = AscW(Mid$(w, i, 1))
            If c < 65 Or c > 90 Then Exit Function
        Next
        If InSet(mSuffix, w) Or InSet(mDir, w) Then Exit Function
        IsRouteId = True
    End If
End Function

' STATE HWY K, ROUTE K, MO 13, US HWY 61, US 61 -> HWY <id>; CO RD 12 -> CR 12; FARM RD 12 -> FR 12.
Public Function Canon(ByVal toks As String) As String
    Dim t() As String, n As Long, i As Long, w As String, nx As String, nx2 As String, out As String
    If Len(toks) = 0 Then Exit Function
    InitWords
    t = Split(toks, " ")
    n = UBound(t) + 1
    Do While i < n
        w = t(i)
        nx = "": nx2 = ""
        If i + 1 < n Then nx = t(i + 1)
        If i + 2 < n Then nx2 = t(i + 2)
        If (w = "STATE" Or w = "MO" Or w = "US") And (nx = "HWY" Or nx = "RTE" Or nx = "RD") And IsRouteId(nx2) Then
            out = out & " HWY " & nx2
            i = i + 3
        ElseIf (w = "MO" Or w = "US") And IsRouteId(nx) And IsDigits(Left$(nx, 1)) Then
            out = out & " HWY " & nx
            i = i + 2
        ElseIf (w = "HWY" Or w = "RTE") And IsRouteId(nx) Then
            out = out & " HWY " & nx
            i = i + 2
        ElseIf w = "CO" And nx = "RD" And IsRouteId(nx2) Then
            out = out & " CR " & nx2
            i = i + 3
        ElseIf w = "FARM" And nx = "RD" And IsRouteId(nx2) Then
            out = out & " FR " & nx2
            i = i + 3
        Else
            out = out & " " & w
            i = i + 1
        End If
    Loop
    Canon = Mid$(out, 2)
End Function

' Drop a unit (STE 5, LOT 4, REAR) and everything after it; never the first word.
Public Function CutUnit(ByVal toks As String) As String
    Dim t() As String, k As Long, out As String
    If Len(toks) = 0 Then Exit Function
    InitWords
    t = Split(toks, " ")
    out = t(0)
    For k = 1 To UBound(t)
        If InSet(mUnit, t(k)) Then Exit For
        out = out & " " & t(k)
    Next
    CutUnit = out
End Function

' A direction or suffix field, standardized: its first word.
Private Function OneWord(ByVal w As String) As String
    Dim t As String, p As Long
    t = Tokens(w)
    p = InStr(1, t, " ")
    If p > 0 Then OneWord = Left$(t, p - 1) Else OneWord = t
End Function

' The state's four street fields -> "name|pre|suf|post|full", standardized.
Public Function StreetPieces(ByVal pre As String, ByVal nm As String, ByVal suf As String, ByVal post As String) As String
    Dim k As String, p As String, s As String, q As String, n As String, v As String
    k = pre & "|" & nm & "|" & suf & "|" & post
    If Not mKeys Is Nothing Then
        On Error Resume Next
        v = mKeys(k)
        If Err.Number = 0 Then
            StreetPieces = v
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    p = OneWord(pre)
    s = OneWord(suf)
    q = OneWord(post)
    n = Canon(CutUnit(Tokens(nm)))
    v = n & "|" & p & "|" & s & "|" & q & "|" & Canon(JoinWords(p, n, s, q))
    If Not mKeys Is Nothing Then mKeys.Add v, k
    StreetPieces = v
End Function

' The postal city, standardized so the state's own spellings meet:
' SAINT LOUIS = ST. LOUIS, O'FALLON = O FALLON.
Public Function CityKeyOf(ByVal city As String) As String
    Dim i As Long, c As Long, buf As String, pos As Long, parts() As String, out As String, w As String
    If Len(city) = 0 Then Exit Function
    buf = Space$(Len(city))
    For i = 1 To Len(city)
        c = AscW(Mid$(city, i, 1))
        If c = 46 Or c = 39 Then
        ElseIf c >= 97 And c <= 122 Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c - 32)
        ElseIf (c >= 65 And c <= 90) Or (c >= 48 And c <= 57) Then
            pos = pos + 1
            Mid$(buf, pos, 1) = ChrW$(c)
        Else
            pos = pos + 1
        End If
    Next
    If pos = 0 Then Exit Function
    parts = Split(Left$(buf, pos), " ")
    For i = 0 To UBound(parts)
        w = parts(i)
        Select Case w
            Case "SAINT": w = "ST"
            Case "SAINTE": w = "STE"
            Case "MOUNT": w = "MT"
            Case "FORT": w = "FT"
        End Select
        out = out & w
    Next
    CityKeyOf = out
End Function

Private Function CachedCityKey(ByVal city As String) As String
    Dim v As String
    On Error Resume Next
    v = mCityKeys(city)
    If Err.Number <> 0 Then
        Err.Clear
        v = CityKeyOf(city)
        mCityKeys.Add v, city
    End If
    CachedCityKey = v
End Function

'=========================================================== the data folder

' Where the quarters' files live: a folder on this PC's own drive, never the
' server. Settings, DataFolder; %LOCALAPPDATA% and the like are filled in.
Public Function DataFolder() As String
    Dim p As String, a As Long, b As Long
    p = GetOpt("DataFolder")
    If p = "" Then p = DATA_DEFAULT
    a = InStr(1, p, "%")
    Do While a > 0
        b = InStr(a + 1, p, "%")
        If b = 0 Then Exit Do
        p = Left$(p, a - 1) & Environ$(Mid$(p, a + 1, b - a - 1)) & Mid$(p, b + 1)
        a = InStr(a, p, "%")
    Loop
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    DataFolder = p
End Function

' "20261001" -> ...\Quarter 2026-10-01.accdb
Public Function QuarterFile(ByVal q As String) As String
    QuarterFile = DataFolder() & "\Quarter " & Left$(q, 4) & "-" & Mid$(q, 5, 2) & "-" & Mid$(q, 7, 2) & ".accdb"
End Function

' The two newest quarters in the data folder ("" if there's none).
Private Sub FindQuarters(nowQ As String, lastQ As String)
    Dim f As String, q As String
    nowQ = "": lastQ = ""
    On Error Resume Next
    f = Dir$(DataFolder() & "\Quarter ????-??-??.accdb")
    Do While Len(f) > 0
        q = Replace(Mid$(f, 9, 10), "-", "")
        If Len(q) = 8 And IsDigits(q) Then
            If q > nowQ Then
                lastQ = nowQ
                nowQ = q
            ElseIf q > lastQ And q <> nowQ Then
                lastQ = q
            End If
        End If
        f = Dir$()
    Loop
End Sub

Public Function NowQuarter() As String
    Dim lastQ As String
    FindQuarters NowQuarter, lastQ
End Function

Public Function LastQuarter() As String
    Dim nowQ As String
    FindQuarters nowQ, LastQuarter
End Function

' The counts saved with the comparison of the two quarters on hand. False if
' they haven't been compared.
Private Function ReadChanges(nowQ As String, lastQ As String, places As Long, unused As Long, _
        streets As Long, zips As Long) As Boolean
    Dim db As DAO.Database
    FindQuarters nowQ, lastQ
    If lastQ = "" Then Exit Function
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(QuarterFile(nowQ), False, True)
    If db Is Nothing Then Exit Function
    If InfoGet(db, "ComparedWith") = lastQ Then
        places = Val(InfoGet(db, "RateChangeCount"))
        unused = Val(InfoGet(db, "RateChangeUnused"))
        streets = Val(InfoGet(db, "MoveStreetCount"))
        zips = Val(InfoGet(db, "MoveZipCount"))
        ReadChanges = (Err.Number = 0)
    End If
    db.Close
End Function

Private Function InfoGet(db As DAO.Database, ByVal k As String) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT InfoValue FROM Info WHERE InfoName='" & k & "'", dbOpenSnapshot)
    If Not rs.EOF Then InfoGet = Nz(rs!InfoValue, "")
    rs.Close
End Function

Private Sub InfoSet(db As DAO.Database, ByVal k As String, ByVal v As String)
    db.Execute "DELETE FROM Info WHERE InfoName='" & k & "'", dbFailOnError
    db.Execute "INSERT INTO Info (InfoName, InfoValue) VALUES ('" & k & "','" & _
        Replace(Left$(v, 255), "'", "''") & "')", dbFailOnError
End Sub

Private Sub MakeFolder(ByVal p As String)
    Dim i As Long
    If Len(Dir$(p, vbDirectory)) > 0 Then Exit Sub
    On Error Resume Next
    For i = 4 To Len(p)
        If Mid$(p, i, 1) = "\" Then MkDir Left$(p, i - 1)
    Next
    MkDir p
    On Error GoTo 0
    If Len(Dir$(p, vbDirectory)) = 0 Then Err.Raise 76, , "Couldn't make the folder " & p
End Sub

' Deletes a file if it's there; an error if it can't.
Private Sub KillFile(ByVal p As String)
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) = 0 Then Exit Sub
    On Error Resume Next
    Kill p
    On Error GoTo 0
    If Len(Dir$(p)) > 0 Then Err.Raise 75, , "Couldn't delete " & p & ". Close Access, open it again and retry."
End Sub

' Puts the text on the clipboard, the same as selecting it and pressing Ctrl+C.
Private Function ToClipboard(ByVal s As String) As Boolean
    Dim h As LongPtr, p As LongPtr, bytes As Long, i As Long
    bytes = (Len(s) + 1) * 2
    For i = 1 To 10
        If OpenClipboard(0) <> 0 Then Exit For
        DoEvents
    Next
    If i > 10 Then Exit Function
    EmptyClipboard
    h = GlobalAlloc(&H2, bytes)
    If h <> 0 Then
        p = GlobalLock(h)
        If p <> 0 Then
            CopyMemory p, StrPtr(s), bytes
            GlobalUnlock h
            If SetClipboardData(13, h) <> 0 Then ToClipboard = True
        End If
        If Not ToClipboard Then GlobalFree h
    End If
    CloseClipboard
End Function

'=================================================================== helpers

Private Function JoinWords(ParamArray w() As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(w)
        If Len(w(i)) > 0 Then s = s & " " & w(i)
    Next
    JoinWords = Mid$(s, 2)
End Function

' Some quarters' files start with the three bytes of a UTF-8 mark (Q3 2026 did,
' Q4 2026 didn't); drop it.
Private Function NoMark(ByVal s As String) As String
    If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then NoMark = Mid$(s, 4) Else NoMark = s
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next
    IsDigits = True
End Function

' 37000-051-000
Private Function LooksLikeCode(ByVal s As String) As Boolean
    If Len(s) <> 13 Then Exit Function
    If Mid$(s, 6, 1) <> "-" Or Mid$(s, 10, 1) <> "-" Then Exit Function
    LooksLikeCode = IsDigits(Left$(s, 5)) And IsDigits(Mid$(s, 7, 3)) And IsDigits(Right$(s, 3))
End Function

' The first comma-separated field in s that looks like a code; "" if none.
' Every code has a dash, so only the dashes are looked at.
Private Function CodeIn(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "-")
    Do While p > 0
        If p > 5 Then
            If LooksLikeCode(Mid$(s, p - 5, 13)) Then
                If p = 6 Or Mid$(s, p - 6, 1) = "," Then
                    If p + 8 > Len(s) Or Mid$(s, p + 8, 1) = "," Then CodeIn = Mid$(s, p - 5, 13): Exit Function
                End If
            End If
        End If
        p = InStr(p + 1, s, "-")
    Loop
End Function

' Split a CSV line that has quotes in it.
Private Function SplitQuoted(ByVal s As String) As String()
    Dim out() As String, n As Long, i As Long, ch As String, cur As String, inQ As Boolean
    ReDim out(0 To 127)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = "," Then
            If n > UBound(out) Then ReDim Preserve out(0 To 2 * n)
            out(n) = cur
            n = n + 1
            cur = ""
        Else
            cur = cur & ch
        End If
    Next
    If n > UBound(out) Then ReDim Preserve out(0 To n)
    out(n) = cur
    ReDim Preserve out(0 To n)
    SplitQuoted = out
End Function

Private Sub AddUnique(arr() As String, n As Long, ByVal v As String)
    Dim i As Long
    For i = 0 To n - 1
        If arr(i) = v Then Exit Sub
    Next
    If n > UBound(arr) Then ReDim Preserve arr(0 To 2 * n + 1)
    arr(n) = v
    n = n + 1
End Sub

Private Sub SortStrings(arr() As String, ByVal n As Long)
    Dim i As Long, j As Long, t As String
    For i = 1 To n - 1
        t = arr(i)
        j = i - 1
        Do While j >= 0
            If StrComp(arr(j), t, vbBinaryCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = t
    Next
End Sub

Private Sub SortLongs(arr() As Long, ByVal n As Long)
    If n > 1 Then QuickLongs arr, 0, n - 1
End Sub

Private Sub QuickLongs(arr() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Long, t As Long
    i = lo: j = hi
    p = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < p: i = i + 1: Loop
        Do While arr(j) > p: j = j - 1: Loop
        If i <= j Then
            t = arr(i): arr(i) = arr(j): arr(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickLongs arr, lo, j
    If i < hi Then QuickLongs arr, i, hi
End Sub

Private Function JoinN(arr() As String, ByVal n As Long, ByVal sep As String) As String
    Dim i As Long, s As String
    For i = 0 To n - 1
        If i > 0 Then s = s & sep
        s = s & arr(i)
    Next
    JoinN = s
End Function

Private Function FileName(ByVal path As String) As String
    FileName = Mid$(path, InStrRev(path, "\") + 1)
End Function

Private Function PickFile(ByVal title As String, ByVal what As String, ByVal pattern As String, _
        ByVal startIn As String) As String
    Dim fd As Object
    If gQuiet Then Exit Function
    Set fd = Application.FileDialog(3)
    With fd
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add what, pattern
        If Len(startIn) > 0 Then .InitialFileName = startIn
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function Say(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    gLastMessage = msg
    If gQuiet Then
        Say = vbOK
    Else
        Say = MsgBox(msg, style, APP_NAME)
    End If
End Function

Private Sub TidyUp()
    On Error Resume Next
    SysCmd acSysCmdRemoveMeter
    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
End Sub

Private Function GetOpt(ByVal k As String) As String
    GetOpt = Nz(DLookup("SettingValue", "Settings", "SettingName='" & k & "'"), "")
End Function

Private Sub SetOpt(ByVal k As String, ByVal v As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "DELETE FROM Settings WHERE SettingName='" & k & "'", dbFailOnError
    db.Execute "INSERT INTO Settings (SettingName, SettingValue) VALUES ('" & k & "','" & _
        Replace(Left$(v, 255), "'", "''") & "')", dbFailOnError
End Sub

Private Function CountSql(db As DAO.Database, ByVal sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    CountSql = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

Private Function TableExists(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    db.TableDefs.Refresh
    For Each td In db.TableDefs
        If td.Name = nm Then
            TableExists = True
            Exit Function
        End If
    Next
End Function

Private Sub DropTable(db As DAO.Database, ByVal nm As String)
    If TableExists(db, nm) Then db.TableDefs.Delete nm
    db.TableDefs.Refresh
End Sub

' spec like "Name:T50;Count:L;Rate:D;Notes:M;When:W" (T text, L whole number, D decimal,
' M long text, W date)
Private Sub NewTable(db As DAO.Database, ByVal nm As String, ByVal spec As String)
    Dim td As DAO.TableDef, fld As DAO.Field, part As Variant, bits() As String
    Set td = db.CreateTableDef(nm)
    For Each part In Split(spec, ";")
        bits = Split(part, ":")
        Select Case Left$(bits(1), 1)
            Case "T"
                Set fld = td.CreateField(bits(0), dbText, CInt(Mid$(bits(1), 2)))
                fld.AllowZeroLength = True
            Case "L"
                Set fld = td.CreateField(bits(0), dbLong)
            Case "D"
                Set fld = td.CreateField(bits(0), dbDouble)
            Case "M"
                Set fld = td.CreateField(bits(0), dbMemo)
                fld.AllowZeroLength = True
            Case "W"
                Set fld = td.CreateField(bits(0), dbDate)
        End Select
        td.Fields.Append fld
    Next
    db.TableDefs.Append td
    db.TableDefs.Refresh
End Sub

Private Sub DeleteObject(ByVal kind As AcObjectType, ByVal nm As String)
    On Error Resume Next
    DoCmd.Close kind, nm, acSaveNo
    DoCmd.DeleteObject kind, nm
End Sub
