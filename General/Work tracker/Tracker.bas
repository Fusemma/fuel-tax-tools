' Tracker: everything you're keeping track of at work, in one list, with a dated log of
' every update and reminders that pop up while the file is open.
'
' Every row on the Items sheet is one thing: a return, a reminder, or anything else (Kind
' says which, and it's only a label). Any row can repeat and any row can remind:
'   Due      the date it's due, with a time when it has one.
'   Repeat   blank for once. Or daily, weekdays, weekly, every 2 weeks, monthly, quarterly,
'            semi-annual, annual, or every 3 days, every 6 weeks, every 2 months.
'   Until    the last date it repeats. Blank means it keeps going.
'   Remind   a time of day for a popup, like 9:00 AM. Blank means no popup.
' Closing a row that repeats makes the next one by itself, due one step on. If you fell
' behind, it steps on until the next one is due today or later.
'
' Reminders: while this file is open, it checks every minute. When a row's Remind time comes
' on the day it's due, a box pops up. Yes means it's done (and the next one gets made if it
' repeats), No snoozes it (15 minutes unless you say different), Cancel keeps it quiet until
' tomorrow. A late row pops once each morning. Nothing pops while the file is closed, so
' leave it open all day.
'
' Who: who has the next move, blank if it's you. On a return, who files it; someone else's
' returns stay on the list as the record, move on by themselves, and show gray on Month.
'
' Every call, email or anything else that happens gets its own dated line on the Log sheet.
' The Today sheet is the home page: late, today, the next 7 days, waiting on them, and gone
' quiet (no update in two weeks). The Month sheet shows it all as a calendar. Both redraw
' themselves, so type on Items.
'
' The buttons (for Update, Close and History, click the row first, on Today, Items or Log):
'   New item     asks what it is, the kind, when it's due, whether it repeats and reminds.
'   Update item  adds a dated line to the Log and lets you move the due date. Type done in
'                the date box to close it.
'   Close item   marks it done. On a return it asks for the confirmation number.
'   History      shows only that row's lines on the Log.
'   Import old   brings everything over from the old tracker. Open the old file first.
'
' Dates: 10/15, 10/15 2pm, tomorrow 9am, 30 (days), 2w, 3m, today. A count of days, weeks or
' months that lands on a weekend moves to the Monday after.
'
' Setup: paste this whole module into a new module in a blank workbook, run TrackerSetup
' (Alt F8), and save the workbook as .xlsm. To update, paste a newer copy over this one
' and run TrackerSetup again. Your rows live on the sheets, so updates never touch them.
' Nothing here saves by itself; save as you would any file.
'
' To have it open every morning: in File Explorer, right-click the saved Tracker file and
' pick Copy. Press the Windows key and R together, type shell:startup and press Enter.
' Right-click in the folder that opens and pick Paste shortcut (on Windows 11 it's under
' Show more options). It opens when you sign in.

Option Explicit

Private Const TODAY_SHEET As String = "Today"
Private Const ITEMS_SHEET As String = "Items"
Private Const LOG_SHEET As String = "Log"
Private Const MONTH_SHEET As String = "Month"
Private Const ITEMS_TABLE As String = "Items"
Private Const LOG_TABLE As String = "UpdateLog"
Private Const HEAD_ROW As Long = 4             ' the tables start here; buttons and help sit above
Private Const SPAN As Long = 5000              ' rows under the headers that get colors and dropdowns
Private Const TODAY_TOP As Long = 5            ' Today's column headers; the lists start under them
Private Const QUIET_DAYS As Long = 14          ' an open row not updated this long has gone quiet
Private Const LEAD_DAYS As Long = 7            ' how far ahead Today's next-days list looks
Private Const HEAD_WORDS As Long = 12          ' a row's head on Today is cut to this many words
Private Const MONTH_TOP As Long = 4            ' the Month sheet's day names; the weeks start under them
Private Const DAY_MAX As Long = 20             ' lines a day's box shows before "and 3 more"
Private Const WEEK_HEIGHT As Long = 70         ' the least a week gets, so an empty one still looks like a calendar
Private Const TICK_SECONDS As Long = 60        ' how often the reminders get checked
Private Const SNOOZE As String = "15"          ' minutes, already filled in when you snooze

' Columns are found by name, so a column you add or move is fine. Sort is hidden; it keeps
' closed rows at the bottom.
Private Const ITEM_COLS As String = "ID|Kind|What|State|Who|Status|Due|Repeat|Until|Remind|Quiet until|" & _
    "Last updated|Latest note|Reference|Opened|Closed|Sort"
Private Const LOG_COLS As String = "When|ID|Item|Note"
Private Const TODAY_COLS As String = "#|What|Done when|Yours|Kind|Latest note"

Private Const ST_THEM As String = "Waiting on them"
Private Const ST_ME As String = "My move"
Private Const ST_DONE As String = "Done"
Private Const KIND_ITEM As String = "Item"
Private Const KIND_REMINDER As String = "Reminder"
Private Const KIND_RETURN As String = "Return"
Private Const REPEAT_LIST As String = "daily,weekdays,weekly,every 2 weeks,monthly,quarterly,semi-annual,annual"
Private Const STATE_LIST As String = "AL,AK,AZ,AR,CA,CO,CT,DE,DC,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS," & _
    "MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY,Federal"
Private Const FIRST_DUE As String = "7"        ' days, already filled in when a new item asks

' Colors, written back to front the way Excel wants them (blue, green, red)
Private Const RED_FILL As Long = &HCEC7FF
Private Const RED_TEXT As Long = &H6009C
Private Const YELLOW_FILL As Long = &H9CEBFF
Private Const YELLOW_TEXT As Long = &H579C&
Private Const ORANGE_FILL As Long = &HD6E4FC
Private Const QUIET_FILL As Long = &HECDFE4
Private Const SOON_FILL As Long = &HDAEFE2
Private Const WAIT_FILL As Long = &HF7EBDD
Private Const HEAD_FILL As Long = &H794E1F
Private Const GRAY_TEXT As Long = &H808080
Private Const HELP_TEXT As Long = &H595959
Private Const BLUE_TEXT As Long = &HB35900
Private Const WEEKEND_FILL As Long = &HF5F5F5
Private Const AWAY_FILL As Long = &HE6E6E6      ' days from the months before and after
Private Const LINE_COLOR As Long = &HBFBFBF

Private mItems As ListObject     ' the two tables, found fresh by Ready each time
Private mLog As ListObject
Private mNext As Date            ' when the next reminder check is set for, 0 when none is
Private mBusy As Boolean         ' a reminder box is up, so the next check waits
Private mFeed As Collection      ' test runs only: answers for the boxes, in order
Private mSaid As String          ' test runs only: what the message boxes would have said
Private mTesting As Boolean      ' test runs only: no real timer, no real boxes
Private mClock As Double         ' test runs only: the time the reminder check sees
Private mRealTimer As Boolean    ' test runs only: run the real timer anyway
Private mTicks As Long           ' test runs only: how many times the timer fired


' ---- the buttons ---------------------------------------------------------

Public Sub TrackerSetup()
    Dim wsI As Worksheet, wsL As Worksheet, wsM As Worksheet, wsT As Worksheet, msg As String
    On Error GoTo Oops
    If IsOldTracker(ThisWorkbook) Then
        Say "This workbook holds the old tracker, so setup left it alone. Paste this macro into a new " & _
            "blank workbook instead and run TrackerSetup there. Then, with this one still open, press " & _
            "Import old on the new one's Items sheet."
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set wsI = EnsureSheet(ITEMS_SHEET, Nothing)
    Set wsL = EnsureSheet(LOG_SHEET, wsI)
    Set wsM = EnsureSheet(MONTH_SHEET, wsL)
    Set wsT = EnsureSheet(TODAY_SHEET, wsM)
    If ThisWorkbook.Worksheets(1).Name <> TODAY_SHEET Then wsT.Move Before:=ThisWorkbook.Worksheets(1)
    Set wsT = ThisWorkbook.Worksheets(TODAY_SHEET)
    Set mItems = EnsureTable(wsI, ITEMS_TABLE, ITEM_COLS)
    Set mLog = EnsureTable(wsL, LOG_TABLE, LOG_COLS)
    DressLog wsL, mLog
    DressItems wsI, mItems
    DressMonth wsM
    DressToday wsT
    Rework Date
    StartTicking
    Application.ScreenUpdating = True
    GoToday
    msg = "Tracker is set up. Today is the home page, Items holds everything you're tracking, Log " & _
          "holds a line for every update, and Month shows it all as a calendar. Reminders pop up " & _
          "while this file is open."
    If ThisWorkbook.FileFormat <> 52 Then
        msg = msg & vbLf & vbLf & "Now save it: File, Save As, and pick Excel Macro-Enabled Workbook (*.xlsm)."
    End If
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Setup stopped: " & Err.Description
End Sub

Public Sub NewItem()
    Dim what As Variant, kind As Variant, due As Variant, rep As Variant, til As Variant, rem_ As Variant
    Dim st As Variant, who As Variant, a As Variant, preset As String, canon As String
    Dim unit As String, every As Long, anchor As String, n As Long, id As Long, remAt As Double
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    Tidy
    what = Ask("What is it? One short line.", "")
    If VarType(what) = vbBoolean Then Exit Sub
    what = Trim$(what)
    If what = "" Then Exit Sub
    Do
        a = Ask("What kind is it: Item, Reminder or Return?", KIND_ITEM)
        If VarType(a) = vbBoolean Then Exit Sub
        kind = KindName(a)
        If kind <> "" Then Exit Do
        Say "Type Item, Reminder or Return."
    Loop
    If kind = KIND_REMINDER Then preset = "today" Else preset = FIRST_DUE
    Do
        a = Ask("When is it due? A date (10/15), days (30), weeks (2w) or months (3m)," & vbLf & _
                "with a time if it has one (10/15 2pm, tomorrow 9am). Type none for no date.", preset)
        If VarType(a) = vbBoolean Then Exit Sub
        If LCase$(Trim$(a)) = "none" Then
            due = Empty
            Exit Do
        End If
        If Trim$(a) = "" Then a = preset
        due = WhenFrom(a, Date)
        If Not IsEmpty(due) Then Exit Do
        Say "Couldn't read """ & a & """ as a date. Try 10/15, 10/15 2pm, 30, 2w or 3m."
    Loop
    Do
        rep = Ask("Does it repeat? Leave it blank if not." & vbLf & _
                  "daily, weekdays, weekly, every 2 weeks, monthly, quarterly, annual, or every 3 days.", "")
        If VarType(rep) = vbBoolean Then Exit Sub
        canon = ReadRepeat(rep, unit, every, anchor)
        If canon <> "?" Then Exit Do
        Say "Couldn't read """ & rep & """. Try daily, weekdays, weekly, monthly, or every 3 days."
    Loop
    If canon <> "" Then
        Do
            til = Ask("Last date it repeats? Leave it blank if it keeps going.", "")
            If VarType(til) = vbBoolean Then Exit Sub
            If Trim$(til) = "" Then Exit Do
            til = DayFrom(LCase$(Trim$(til)), Date)
            If Not IsEmpty(til) Then Exit Do
            Say "Couldn't read that as a date. Try 12/31."
        Loop
    End If
    If Not IsEmpty(due) Then
        If HasTime(due) Then preset = Format$(due, "h:mm AM/PM") Else preset = ""
    Else
        preset = ""
    End If
    Do
        rem_ = Ask("Pop up a reminder at what time? Like 9am or 2:30pm." & vbLf & _
                   "Leave it blank for no popup.", preset)
        If VarType(rem_) = vbBoolean Then Exit Sub
        remAt = TimeTyped(rem_)
        If remAt >= 0 Or Trim$(rem_) = "" Or LCase$(Trim$(rem_)) = "none" Then Exit Do
        Say "Couldn't read """ & rem_ & """ as a time. Try 9am or 2:30pm."
    Loop
    st = Ask("Which state? Leave it blank if none.", "")
    If VarType(st) = vbBoolean Then Exit Sub
    who = Ask("Who has the next move? Leave it blank if it's you." & vbLf & _
              "On a return, who files it.", "")
    If VarType(who) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    id = NextId()
    n = NewRow(mItems)
    Cel(mItems, n, "ID").Value = id
    Cel(mItems, n, "Kind").Value = kind
    Cel(mItems, n, "What").Value = SafeText(what)
    Cel(mItems, n, "State").Value = CleanState(st)
    Cel(mItems, n, "Who").Value = SafeText(Trim$(who))
    Cel(mItems, n, "Status").Value = StatusFor(who)
    If Not IsEmpty(due) Then Cel(mItems, n, "Due").Value = due
    If canon <> "" Then Cel(mItems, n, "Repeat").Value = canon
    If canon <> "" And Not IsEmpty(til) Then Cel(mItems, n, "Until").Value = til
    If remAt >= 0 Then Cel(mItems, n, "Remind").Value = CDate(remAt)
    Cel(mItems, n, "Opened").Value = Date
    Rework MonthShown()
    Application.ScreenUpdating = True
    SelectItem id, "What"
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "New item stopped: " & Err.Description
End Sub

Public Sub UpdateItem()
    Dim id As Variant, n As Long, note As Variant, back As Variant, due As Variant, cameFrom As String, old As Double
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    Tidy
    cameFrom = ActiveSheet.Name
    id = PickedId()
    If IsEmpty(id) Then Exit Sub
    n = FindItem(id)
    If n = 0 Then
        Say "There's no item number " & id & "."
        Exit Sub
    End If
    note = Ask("What happened with #" & id & ", " & Short(Cel(mItems, n, "What").Value) & "?" & vbLf & _
               "One line, like: called, they said 30 to 60 days.", "")
    If VarType(note) = vbBoolean Then Exit Sub
    note = Trim$(note)
    If note = "" Then Exit Sub
    old = DateOf(Cel(mItems, n, "Due").Value)
    Do
        back = Ask("When is it due now? Leave it blank to keep it (" & DueText(old) & ")." & vbLf & _
                   "A date (10/15), days (30), weeks (2w) or months (3m). Type done to close it.", "")
        If VarType(back) = vbBoolean Then back = ""       ' the note still gets logged
        back = LCase$(Trim$(back))
        If back = "" Or back = "done" Then Exit Do
        due = WhenFrom(back, Date)
        If Not IsEmpty(due) Then Exit Do
        Say "Couldn't read """ & back & """ as a date. Try 10/15, 10/15 2pm, 30, 2w or 3m."
    Loop
    Application.ScreenUpdating = False
    If back = "done" Then
        CloseRow n, note
    Else
        AddLogLine id, note
        If Not IsEmpty(due) Then
            ' A new day with no time keeps the old time
            If Not HasTime(due) And HasTime(old) Then due = CDate(CDbl(due) + (old - Int(old)))
            Cel(mItems, n, "Due").Value = due
            Cel(mItems, n, "Quiet until").ClearContents
        End If
    End If
    Rework MonthShown()
    Application.ScreenUpdating = True
    Select Case cameFrom
        Case LOG_SHEET: ShowHistoryFor id
        Case TODAY_SHEET: GoToday id
        Case Else: SelectItem id, "Latest note"
    End Select
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Update item stopped: " & Err.Description
End Sub

Public Sub CloseItem()
    Dim id As Variant, n As Long, why As Variant, cameFrom As String
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    Tidy
    cameFrom = ActiveSheet.Name
    id = PickedId()
    If IsEmpty(id) Then Exit Sub
    n = FindItem(id)
    If n = 0 Then
        Say "There's no item number " & id & "."
        Exit Sub
    End If
    If Txt(Cel(mItems, n, "Status").Value) = ST_DONE Then
        Say "#" & id & " is already closed."
        Exit Sub
    End If
    If Not Confirm("Close #" & id & ", " & Short(Cel(mItems, n, "What").Value) & "?") Then Exit Sub
    If Txt(Cel(mItems, n, "Kind").Value) = KIND_RETURN Then
        why = AskReference(n)
        Application.ScreenUpdating = False
        If why = "" Then CloseRow n, "Closed." Else CloseRow n, "Closed, confirmation " & why & "."
    Else
        why = Ask("Why is it closed? One line, or leave it blank.", "")
        If VarType(why) = vbBoolean Then why = ""
        why = Trim$(why)
        Application.ScreenUpdating = False
        If why = "" Then CloseRow n, "Closed." Else CloseRow n, "Closed: " & why
    End If
    Rework MonthShown()
    Application.ScreenUpdating = True
    If cameFrom = TODAY_SHEET Then GoToday Else SelectItem id, "Status"
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Close item stopped: " & Err.Description
End Sub

Public Sub ShowHistory()
    Dim id As Variant
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    Tidy
    id = PickedId()
    If IsEmpty(id) Then Exit Sub
    If FindItem(id) = 0 Then
        Say "There's no item number " & id & "."
        Exit Sub
    End If
    ShowHistoryFor id
    Exit Sub
Oops:
    Say "History stopped: " & Err.Description
End Sub

Public Sub BackToItems()
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    ClearFilter mLog
    mItems.Parent.Activate
    Exit Sub
Oops:
    Say "Back to items stopped: " & Err.Description
End Sub

Public Sub RefreshToday()
    ' Also brings the Month sheet back to this month
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    Application.ScreenUpdating = False
    Rework Date
    Application.ScreenUpdating = True
    GoToday
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Refresh stopped: " & Err.Description
End Sub

Public Sub ShowMonth()
    MonthButton 0
End Sub

Public Sub PreviousMonth()
    MonthButton -1
End Sub

Public Sub NextMonth()
    MonthButton 1
End Sub

Public Sub ImportOld()
    ' Brings everything over from an open copy of the old tracker: its items, its log, and each
    ' Calendar row as a Return that repeats. The old file isn't changed.
    Dim wb As Workbook, src As Workbook, oldI As ListObject, oldL As ListObject, oldC As ListObject
    Dim i As Long, n As Long, k As Long, offset As Long, oldId As Long, id As Long, ni As Long, nl As Long, nc As Long
    Dim what As String, who As String, myName As String, q As String, months As Long, dueDay As String
    Dim due As Long, rep As String, mine As Boolean
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If IsOldTracker(wb) Then
                Set src = wb
                Exit For
            End If
        End If
    Next
    If src Is Nothing Then
        Say "Open the old Tracker first (File, Open), come back to this one, and press Import old again."
        Exit Sub
    End If
    Set oldI = TableIn(src, ITEMS_TABLE)
    Set oldL = TableIn(src, LOG_TABLE)
    If oldL Is Nothing Then Set oldL = TableIn(src, "TouchLog")      ' the name the oldest copies used
    Set oldC = TableIn(src, "Calendar")
    If Not oldI Is Nothing Then ni = FilledRows(oldI, "What")
    If Not oldL Is Nothing Then nl = FilledRows(oldL, "Note")
    If Not oldC Is Nothing Then nc = FilledRows(oldC, "What")
    q = "Bring in " & ni & " items, " & nc & " calendar rows and " & nl & " log lines from " & src.Name & "?"
    If FilledRows(mItems, "What") > 0 Then
        q = q & vbLf & vbLf & "This tracker already has rows. If they came from an earlier import, you'll get doubles."
    End If
    If Not Confirm(q) Then Exit Sub
    Application.ScreenUpdating = False
    ' Old numbers stay the same in an empty tracker, so #1 is still #1
    offset = NextId() - 1
    If Not oldI Is Nothing Then
        For i = 1 To oldI.ListRows.Count
            what = Txt(OldValue(oldI, i, "What"))
            If what <> "" Then
                oldId = CLng(Val(Txt(OldValue(oldI, i, "ID"))))
                If oldId > 0 Then id = oldId + offset Else id = NextId()
                n = NewRow(mItems)
                Cel(mItems, n, "ID").Value = id
                Cel(mItems, n, "What").Value = SafeText(what)
                Cel(mItems, n, "State").Value = CleanState(Txt(OldValue(oldI, i, "State")))
                Cel(mItems, n, "Who").Value = SafeText(Txt(OldValue(oldI, i, "Who")))
                Cel(mItems, n, "Status").Value = Txt(OldValue(oldI, i, "Status"))
                due = DayNo(OldValue(oldI, i, "Due date"))
                If due > 0 Then
                    Cel(mItems, n, "Kind").Value = KIND_RETURN
                    Cel(mItems, n, "Due").Value = CDate(due)
                Else
                    Cel(mItems, n, "Kind").Value = KIND_ITEM
                    If DayNo(OldValue(oldI, i, "Check back")) > 0 Then Cel(mItems, n, "Due").Value = OldValue(oldI, i, "Check back")
                End If
                Cel(mItems, n, "Reference").Value = SafeText(Txt(OldValue(oldI, i, "Case number")) & _
                                                             Txt(OldValue(oldI, i, "Ref number")))
                If DayNo(OldValue(oldI, i, "Opened")) > 0 Then Cel(mItems, n, "Opened").Value = OldValue(oldI, i, "Opened")
                If DayNo(OldValue(oldI, i, "Closed")) > 0 Then Cel(mItems, n, "Closed").Value = OldValue(oldI, i, "Closed")
                k = k + 1
            End If
        Next
    End If
    If Not oldL Is Nothing Then
        ClearFilter mLog
        For i = 1 To oldL.ListRows.Count
            If Txt(OldValue(oldL, i, "Note")) <> "" Then
                n = NewRow(mLog)
                Cel(mLog, n, "When").Value = OldValue(oldL, i, "When")
                Cel(mLog, n, "ID").Value = CLng(Val(Txt(OldValue(oldL, i, "ID")))) + offset
                Cel(mLog, n, "Note").Value = SafeText(Txt(OldValue(oldL, i, "Note")))
            End If
        Next
    End If
    If Not oldC Is Nothing Then
        On Error Resume Next
        myName = LCase$(Txt(src.Names("YourName").RefersToRange.Value))
        On Error GoTo Oops
        For i = 1 To oldC.ListRows.Count
            what = Txt(OldValue(oldC, i, "What"))
            If what <> "" Then
                who = Txt(OldValue(oldC, i, "Who files it"))
                mine = (who = "" Or LCase$(who) = "me" Or LCase$(who) = myName)
                months = OldMonths(Txt(OldValue(oldC, i, "How often")))
                due = DayNo(OldValue(oldC, i, "Next due"))
                rep = ""
                If months > 0 Then
                    rep = RepeatName("m", months)
                    dueDay = LCase$(Txt(OldValue(oldC, i, "Due day")))
                    If dueDay = "last" Then
                        rep = rep & " on the last day"
                    ElseIf Val(dueDay) >= 29 And Val(dueDay) <= 31 Then
                        rep = rep & " on the " & Ordinal(CLng(Val(dueDay)))
                    End If
                End If
                id = NextId()
                n = NewRow(mItems)
                Cel(mItems, n, "ID").Value = id
                Cel(mItems, n, "Kind").Value = KIND_RETURN
                Cel(mItems, n, "What").Value = SafeText(what)
                Cel(mItems, n, "State").Value = CleanState(Txt(OldValue(oldC, i, "State")))
                If Not mine Then Cel(mItems, n, "Who").Value = SafeText(who)
                Cel(mItems, n, "Status").Value = IIf(mine, ST_ME, ST_THEM)
                If due > 0 Then Cel(mItems, n, "Due").Value = CDate(due)
                Cel(mItems, n, "Repeat").Value = rep
                Cel(mItems, n, "Opened").Value = Date
                AddLogLine id, "From the old calendar." & EndDot(OldValue(oldC, i, "How it's filed")) & _
                    EndDot(OldValue(oldC, i, "Notes"))
            End If
        Next
    End If
    Rework Date
    Application.ScreenUpdating = True
    mItems.Parent.Activate
    Say "Brought in " & k & " items, " & nc & " returns from the old calendar, and " & nl & _
        " log lines from " & src.Name & ". Look them over, then save this file."
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Import old stopped: " & Err.Description
End Sub

Public Sub Auto_Open()
    ' Excel runs this when the workbook opens: a fresh Today, the Month on this month, and the
    ' reminders start
    Dim lo As ListObject
    Set lo = GetTable(ITEMS_TABLE)
    If lo Is Nothing Then Exit Sub                       ' TrackerSetup hasn't run yet
    If Not HasColumns(lo, ITEM_COLS) Then Exit Sub        ' or hasn't run since this copy was pasted in
    RefreshToday
End Sub

Public Sub Auto_Close()
    ' Excel runs this when the workbook closes. A check left waiting would open the file again.
    StopTicking
End Sub

Public Sub TrackerTick()
    ' The reminder check, every minute while the file is open
    mTicks = mTicks + 1
    mNext = 0
    StartTicking
    If mTesting And mFeed Is Nothing Then Exit Sub        ' a test never shows a real box
    CheckReminders
End Sub


' ---- reminders ---------------------------------------------------------------

Private Sub StartTicking()
    StopTicking
    mNext = Now + TimeSerial(0, 0, TICK_SECONDS)
    If mTesting And Not mRealTimer Then Exit Sub
    On Error Resume Next
    Application.OnTime mNext, TickName()
End Sub

Private Sub StopTicking()
    If mNext = 0 Then Exit Sub
    If Not mTesting Or mRealTimer Then
        On Error Resume Next
        Application.OnTime mNext, TickName(), , False
    End If
    mNext = 0
End Sub

Private Function TickName() As String
    TickName = "'" & ThisWorkbook.Name & "'!TrackerTick"
End Function

Private Function NowT() As Double
    ' The time the reminders go by. A test can set it.
    If mClock > 0 Then NowT = mClock Else NowT = CDbl(Now)
End Function

Private Sub CheckReminders()
    ' Every open row whose Remind time has come on the day it's due gets a box, unless it's
    ' snoozed or quiet until later. A late row pops once a day, on the first check.
    Dim t As Double, today As Long, i As Long, due As Double, at As Double, quiet As Double
    Dim ids As New Collection, id As Variant, n As Long, lo As ListObject
    If mBusy Then Exit Sub
    Set lo = GetTable(ITEMS_TABLE)
    If lo Is Nothing Or GetTable(LOG_TABLE) Is Nothing Then Exit Sub
    If Not HasColumns(lo, ITEM_COLS) Then Exit Sub
    Set mItems = lo
    Set mLog = GetTable(LOG_TABLE)
    mBusy = True
    On Error GoTo Done
    t = NowT()
    today = Int(t)
    If SheetExists(TODAY_SHEET) Then
        ' The first check after midnight redraws Today for the new day
        If DayNo(ThisWorkbook.Worksheets(TODAY_SHEET).Range("A3").Value) < today Then
            Application.ScreenUpdating = False
            Rework MonthShown(), True
            Application.ScreenUpdating = True
        End If
    End If
    For i = 1 To mItems.ListRows.Count
        If IsOpenRow(i) Then
            at = RemindOf(Cel(mItems, i, "Remind").Value)
            If at >= 0 And Not IsTheirs(i) Then
                due = DateOf(Cel(mItems, i, "Due").Value)
                If due = 0 Then
                    at = today + at                ' no due date: every day at its time
                ElseIf Int(due) = today Then
                    at = today + at
                ElseIf Int(due) < today Then
                    at = today                     ' late: the first check of the day
                Else
                    at = -1                        ' not due yet
                End If
                quiet = DateOf(Cel(mItems, i, "Quiet until").Value)
                If at >= 0 And t >= at And t >= quiet Then ids.Add CLng(Cel(mItems, i, "ID").Value)
            End If
        End If
    Next
    For Each id In ids
        n = FindItem(id)
        If n > 0 Then
            If IsOpenRow(n) Then PopFor n, t
        End If
    Next
Done:
    Application.ScreenUpdating = True
    mBusy = False
End Sub

Private Sub PopFor(ByVal n As Long, ByVal t As Double)
    ' One reminder box. Yes closes the row, No snoozes it, Cancel keeps it quiet until tomorrow.
    Dim id As Long, due As Double, msg As String, a As Variant, mins As Double, ref As String
    id = CLng(Cel(mItems, n, "ID").Value)
    due = DateOf(Cel(mItems, n, "Due").Value)
    msg = "#" & id & "  " & Head(n) & vbLf
    If due = 0 Then
        msg = msg & "no due date"
    ElseIf Int(due) < Int(t) Then
        msg = msg & "was due " & DueText(due)
    Else
        msg = msg & "due " & DueText(due)
    End If
    msg = msg & vbLf & vbLf & "Yes: it's done.    No: snooze it.    Cancel: quiet until tomorrow."
    Select Case PopChoice(msg)
        Case vbYes
            If Txt(Cel(mItems, n, "Kind").Value) = KIND_RETURN Then ref = AskReference(n)
            Application.ScreenUpdating = False
            If ref = "" Then CloseRow n, "Done, from the reminder." Else CloseRow n, "Done, from the reminder, confirmation " & ref & "."
            Rework MonthShown(), True
            Application.ScreenUpdating = True
        Case vbNo
            a = Ask("Snooze #" & id & " how many minutes?", SNOOZE)
            If VarType(a) = vbBoolean Then a = SNOOZE
            mins = Val(a)
            If mins < 1 Then mins = Val(SNOOZE)
            Cel(mItems, n, "Quiet until").Value = CDate(t + mins / 1440)
        Case Else
            Cel(mItems, n, "Quiet until").Value = CDate(Int(t) + 1)
    End Select
End Sub

Private Function PopChoice(ByVal msg As String) As VbMsgBoxResult
    Dim a As String
    If Not mFeed Is Nothing Then
        mSaid = mSaid & "POP " & Replace(msg, vbLf, " / ") & vbLf
        If mFeed.Count > 0 Then
            a = LCase$(mFeed(1))
            mFeed.Remove 1
        End If
        Select Case a
            Case "yes": PopChoice = vbYes
            Case "no": PopChoice = vbNo
            Case Else: PopChoice = vbCancel
        End Select
        Exit Function
    End If
    ' On top of every window, so it shows even when Excel is behind something else
    PopChoice = MsgBox(msg, vbYesNoCancel + vbExclamation + vbSystemModal + vbMsgBoxSetForeground, "Tracker reminder")
End Function


' ---- working with the two tables -------------------------------------------

Private Function Ready() As Boolean
    Set mItems = GetTable(ITEMS_TABLE)
    Set mLog = GetTable(LOG_TABLE)
    If mItems Is Nothing Or mLog Is Nothing Then
        Say "Run TrackerSetup first: Alt F8, pick TrackerSetup, Run."
        Exit Function
    End If
    If Not HasColumns(mItems, ITEM_COLS) Then
        ' A newer copy was pasted in and setup hasn't brought the sheets up to date yet
        Say "Run TrackerSetup first: Alt F8, pick TrackerSetup, Run. It brings the sheets up to date."
        Exit Function
    End If
    Freshen
    If mNext = 0 Then StartTicking            ' a button press brings the reminders back if they stopped
    Ready = True
End Function

Private Function GetTable(ByVal nm As String) As ListObject
    Set GetTable = TableIn(ThisWorkbook, nm)
End Function

Private Function TableIn(ByVal wb As Workbook, ByVal nm As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nm Then
                Set TableIn = lo
                Exit Function
            End If
        Next
    Next
End Function

Private Function IsOldTracker(ByVal wb As Workbook) As Boolean
    ' The old tracker had a Calendar table and a Check back column on Items
    Dim lo As ListObject
    If Not TableIn(wb, "Calendar") Is Nothing Then
        IsOldTracker = True
        Exit Function
    End If
    Set lo = TableIn(wb, ITEMS_TABLE)
    If Not lo Is Nothing Then IsOldTracker = HasColumn(lo, "Check back")
End Function

Private Function Cel(ByVal lo As ListObject, ByVal rowNo As Long, ByVal colName As String) As Range
    Set Cel = lo.DataBodyRange.Cells(rowNo, lo.ListColumns(colName).Index)
End Function

Private Function OldValue(ByVal lo As ListObject, ByVal rowNo As Long, ByVal colName As String) As Variant
    ' A cell's value from the old tracker, or Empty when it has no such column
    If HasColumn(lo, colName) Then OldValue = Cel(lo, rowNo, colName).Value
End Function

Private Function FilledRows(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If Txt(Cel(lo, i, colName).Value) <> "" Then FilledRows = FilledRows + 1
    Next
End Function

Private Function NewRow(ByVal lo As ListObject) As Long
    ' Uses the table's empty last row if it has one, or grows the table over the blank row
    ' under it. Inserting a row instead would shove the colors and dropdowns down the sheet.
    If lo.ListRows.Count = 0 Then
        ' Every row was deleted by hand, so there's nothing to grow into. Adding one knocks the
        ' colors off the top row, so they get laid again.
        lo.ListRows.Add
        If lo.Name = ITEMS_TABLE Then AddColors lo.Parent, lo
    ElseIf Not RowIsBlank(lo, lo.ListRows.Count) Then
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + 1)
    End If
    NewRow = lo.ListRows.Count
End Function

Private Function RowIsBlank(ByVal lo As ListObject, ByVal n As Long) As Boolean
    Dim c As Range
    For Each c In lo.ListRows(n).Range.Cells
        If Not c.HasFormula Then
            If Txt(c.Value) <> "" Then Exit Function
        End If
    Next
    RowIsBlank = True
End Function

Private Function NextId() As Long
    ' One past the highest number on either sheet, so a deleted item's number is never
    ' reused and its old log lines can't land on a new item
    Dim top As Double
    With Application.WorksheetFunction
        If mItems.ListRows.Count > 0 Then top = .Max(mItems.ListColumns("ID").DataBodyRange)
        If mLog.ListRows.Count > 0 Then top = .Max(top, .Max(mLog.ListColumns("ID").DataBodyRange))
    End With
    NextId = top + 1
End Function

Private Function FindItem(ByVal id As Variant) As Long
    Dim hit As Variant
    If mItems.ListRows.Count = 0 Or IsEmpty(id) Then Exit Function
    hit = Application.Match(CDbl(id), mItems.ListColumns("ID").DataBodyRange, 0)
    If Not IsError(hit) Then FindItem = CLng(hit)
End Function

Private Function PickedId() As Variant
    ' The item whose row is selected, on Today, Items or Log. Otherwise it asks for the number.
    Dim lo As ListObject, c As Range, v As Variant, a As Variant
    PickedId = Empty
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set c = ActiveCell
        If ActiveSheet.Name = TODAY_SHEET And c.Row > TODAY_TOP Then
            v = ActiveSheet.Cells(c.Row, 1).Value            ' Today's # column
            If IsNumber(Txt(v)) Then
                PickedId = CLng(v)
                Exit Function
            End If
        End If
        On Error Resume Next
        Set lo = c.ListObject
        On Error GoTo 0
        If Not lo Is Nothing Then
            If lo.Name = ITEMS_TABLE Or lo.Name = LOG_TABLE Then
                If Not lo.DataBodyRange Is Nothing Then
                    If Not Intersect(c, lo.DataBodyRange) Is Nothing Then
                        v = Cel(lo, c.Row - lo.DataBodyRange.Row + 1, "ID").Value
                        If IsNumber(Txt(v)) Then
                            PickedId = CLng(v)
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    a = Ask("Which item number?" & vbLf & "(Click an item's row before the button to skip this.)", "")
    If VarType(a) = vbBoolean Then Exit Function
    a = Trim$(a)
    If IsNumber(a) Then
        PickedId = CLng(a)
    ElseIf a <> "" Then
        Say "That isn't an item number."
    End If
End Function

Private Sub SelectItem(ByVal id As Long, ByVal colName As String)
    Dim n As Long
    n = FindItem(id)
    If n = 0 Then Exit Sub
    mItems.Parent.Activate
    Cel(mItems, n, colName).Select
End Sub

Private Sub Rework(ByVal monthDay As Date, Optional ByVal quiet As Boolean = False)
    ' Everything after a change: tidy the rows, the worked-out columns, the order, Today and Month
    Dim problems As String
    problems = Tidy()
    EnsureFormulas
    SortItems mItems
    NoCorners
    BuildToday
    BuildMonth monthDay
    If problems <> "" And Not quiet Then
        Say "Couldn't read Repeat on these, so they won't repeat until it's fixed. Try daily, weekdays, " & _
            "weekly, monthly, or every 3 days:" & problems
    End If
End Sub

Private Function Tidy() As String
    ' Fills in what a hand-typed row is missing and puts Kind and Repeat in the list's words.
    ' A row copied from another one comes with its number, so the lower of the two gets a new
    ' one. Someone else's returns move on by themselves. Returns the rows whose Repeat it
    ' couldn't read.
    Dim lo As ListObject, i As Long, nextNo As Long, st As String, id As String, seen As String, k As String
    Dim rep As String, canon As String, unit As String, every As Long, anchor As String, due As Double
    Dim fresh As Boolean, problems As String
    Set lo = mItems
    If lo.ListRows.Count = 0 Then Exit Function
    nextNo = NextId()
    seen = "|"
    For i = 1 To lo.ListRows.Count
        If Txt(Cel(lo, i, "What").Value) <> "" Then
            id = Txt(Cel(lo, i, "ID").Value)
            fresh = (id = "")
            If id = "" Or InStr(seen, "|" & id & "|") > 0 Then
                Cel(lo, i, "ID").Value = nextNo
                id = CStr(nextNo)
                nextNo = nextNo + 1
            End If
            seen = seen & id & "|"
            If Txt(Cel(lo, i, "Opened").Value) = "" Then Cel(lo, i, "Opened").Value = Date
            k = KindName(Txt(Cel(lo, i, "Kind").Value))
            If k <> "" And k <> Txt(Cel(lo, i, "Kind").Value) Then Cel(lo, i, "Kind").Value = k
            st = Txt(Cel(lo, i, "Status").Value)
            If st = "" Then
                st = StatusFor(Cel(lo, i, "Who").Value)
                Cel(lo, i, "Status").Value = st
            End If
            If st = ST_DONE Then
                If Txt(Cel(lo, i, "Closed").Value) = "" Then Cel(lo, i, "Closed").Value = Date
            ElseIf Txt(Cel(lo, i, "Closed").Value) <> "" Then
                Cel(lo, i, "Closed").ClearContents
            End If
            due = DateOf(Cel(lo, i, "Due").Value)
            ' A hand-typed row with a time on its due date reminds at that time
            If fresh And due > 0 And HasTime(due) And Txt(Cel(lo, i, "Remind").Value) = "" Then
                Cel(lo, i, "Remind").Value = CDate(due - Int(due))
            End If
            rep = Txt(Cel(lo, i, "Repeat").Value)
            If rep <> "" Then
                canon = ReadRepeat(rep, unit, every, anchor)
                If canon = "?" Then
                    problems = problems & vbLf & "  #" & id & ", " & Short(Cel(lo, i, "What").Value) & ": " & rep
                Else
                    If unit = "m" And due > 0 Then canon = RepeatName(unit, every) & AnchorWords(KeepAnchor(Int(due), anchor))
                    If canon <> rep Then Cel(lo, i, "Repeat").Value = canon
                    If st <> ST_DONE And IsTheirs(i) Then RollTheirs i
                End If
            End If
        End If
    Next
    Tidy = problems
End Function

Private Sub NoCorners()
    ' Words like "every 3 days" aren't on the Repeat dropdown, which Excel marks with a green
    ' corner. The mark belongs to the cell, not the row, so this runs after the sort.
    Dim i As Long, c As Range
    For i = 1 To mItems.ListRows.Count
        Set c = Cel(mItems, i, "Repeat")
        If Txt(c.Value) <> "" Then
            c.Errors(xlListDataValidation).Ignore = (InStr("," & REPEAT_LIST & ",", "," & Txt(c.Value) & ",") = 0)
        End If
    Next
End Sub

Private Sub RollTheirs(ByVal n As Long)
    ' Someone else's return that repeats moves on to its next due date once a date passes,
    ' the way the old Calendar did, so it never sits late on your list
    Dim unit As String, every As Long, anchor As String, due As Double, nxt As Double, untilDay As Long, k As Long
    due = DateOf(Cel(mItems, n, "Due").Value)
    If due = 0 Or Int(due) >= CLng(Date) Then Exit Sub
    If ReadRepeat(Txt(Cel(mItems, n, "Repeat").Value), unit, every, anchor) = "" Then Exit Sub
    If anchor = "" Then anchor = CStr(Day(CDate(Int(due))))
    untilDay = DayNo(Cel(mItems, n, "Until").Value)
    nxt = due
    Do While Int(nxt) < CLng(Date) And k < 5000
        due = StepDue(nxt, unit, every, anchor)
        If untilDay > 0 And Int(due) > untilDay Then Exit Do
        nxt = due
        k = k + 1
    Loop
    If nxt <> DateOf(Cel(mItems, n, "Due").Value) Then Cel(mItems, n, "Due").Value = CDate(nxt)
End Sub

Private Sub EnsureFormulas()
    ' Rewrites the worked-out columns every time, so a typed-over formula heals itself.
    ' AGGREGATE skips a When that isn't a date, so one bad cell on the Log can't blank
    ' out every item's note.
    Dim newest As String
    If Not mItems.DataBodyRange Is Nothing Then
        newest = "AGGREGATE(14,6," & LOG_TABLE & "[When]/(" & LOG_TABLE & "[ID]=[@ID]),1)"
        mItems.ListColumns("Last updated").DataBodyRange.Formula = _
            "=IF([@ID]="""","""",MAX(N([@Opened]),IFERROR(" & newest & ",0)))"
        ' The note on the item's newest log line, even if older lines were typed in later
        mItems.ListColumns("Latest note").DataBodyRange.Formula = _
            "=IFERROR(LOOKUP(2,1/((" & LOG_TABLE & "[ID]=[@ID])*(" & LOG_TABLE & "[When]=" & newest & "))," & _
            LOG_TABLE & "[Note]),"""")"
        mItems.ListColumns("Sort").DataBodyRange.Formula = "=IF([@Status]=""" & ST_DONE & """,1,0)"
    End If
    If Not mLog.DataBodyRange Is Nothing Then
        mLog.ListColumns("Item").DataBodyRange.Formula = _
            "=IFERROR(INDEX(" & ITEMS_TABLE & "[What],MATCH([@ID]," & ITEMS_TABLE & "[ID],0)),"""")"
    End If
End Sub

Private Sub SortItems(ByVal lo As ListObject)
    ' Soonest due first, no date after that, closed rows at the bottom
    If lo.ListRows.Count < 2 Then Exit Sub
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns("Sort").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=lo.ListColumns("Due").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=lo.ListColumns("ID").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
End Sub

Private Function CloseRow(ByVal n As Long, ByVal note As String) As Long
    ' Marks row n done and logs the note. One that repeats gets its next row made, and the
    ' note says which number that is. Returns the new row's number, 0 when there's none.
    Dim id As Long, tail As String
    id = CLng(Cel(mItems, n, "ID").Value)
    CloseRow = MakeNext(n, tail)
    Cel(mItems, n, "Status").Value = ST_DONE
    Cel(mItems, n, "Closed").Value = Date
    Cel(mItems, n, "Quiet until").ClearContents
    AddLogLine id, note & tail
End Function

Private Function MakeNext(ByVal n As Long, tail As String) As Long
    ' The next row of one that repeats: same words, due one step on from this one's due date,
    ' stepped again until it's due today or later. None once it's past Until.
    Dim unit As String, every As Long, anchor As String, canon As String
    Dim id As Long, newId As Long, m As Long, due As Double, nxt As Double, untilDay As Long, skipped As Long
    tail = ""
    canon = ReadRepeat(Txt(Cel(mItems, n, "Repeat").Value), unit, every, anchor)
    If canon = "" Or canon = "?" Then Exit Function
    id = CLng(Cel(mItems, n, "ID").Value)
    due = DateOf(Cel(mItems, n, "Due").Value)
    If due = 0 Then due = CDbl(Date)
    If anchor = "" Then anchor = CStr(Day(CDate(Int(due))))
    nxt = StepDue(due, unit, every, anchor)
    Do While Int(nxt) < CLng(Date) And skipped < 5000
        nxt = StepDue(nxt, unit, every, anchor)
        skipped = skipped + 1
    Loop
    untilDay = DayNo(Cel(mItems, n, "Until").Value)
    If untilDay > 0 And Int(nxt) > untilDay Then
        tail = " That was the last one; it repeated until " & Format$(CDate(untilDay), "ddd m/d") & "."
        Exit Function
    End If
    newId = NextId()
    m = NewRow(mItems)
    Cel(mItems, m, "ID").Value = newId
    Cel(mItems, m, "Kind").Value = Cel(mItems, n, "Kind").Value
    Cel(mItems, m, "What").Value = SafeText(Txt(Cel(mItems, n, "What").Value))
    Cel(mItems, m, "State").Value = Cel(mItems, n, "State").Value
    Cel(mItems, m, "Who").Value = SafeText(Txt(Cel(mItems, n, "Who").Value))
    Cel(mItems, m, "Status").Value = StatusFor(Cel(mItems, n, "Who").Value)
    Cel(mItems, m, "Due").Value = CDate(nxt)
    Cel(mItems, m, "Repeat").Value = Cel(mItems, n, "Repeat").Value
    Cel(mItems, m, "Until").Value = Cel(mItems, n, "Until").Value
    Cel(mItems, m, "Remind").Value = Cel(mItems, n, "Remind").Value
    Cel(mItems, m, "Opened").Value = Date
    If skipped > 0 Then
        AddLogLine newId, "Repeats from #" & id & ", next due " & DueText(nxt) & ", skipped " & skipped & "."
    Else
        AddLogLine newId, "Repeats from #" & id & ", next due " & DueText(nxt) & "."
    End If
    tail = " Next one is #" & newId & "."
    MakeNext = newId
End Function

Private Function AskReference(ByVal n As Long) As String
    ' A return's confirmation number, kept in Reference
    Dim a As Variant
    a = Ask("Confirmation number? Leave it blank if there isn't one.", Txt(Cel(mItems, n, "Reference").Value))
    If VarType(a) = vbBoolean Then a = ""
    a = Trim$(a)
    If a <> "" Then Cel(mItems, n, "Reference").Value = SafeText(a)
    AskReference = a
End Function

Private Sub AddLogLine(ByVal id As Long, ByVal note As String)
    Dim n As Long
    ClearFilter mLog
    n = NewRow(mLog)
    Cel(mLog, n, "When").Value = Now
    Cel(mLog, n, "ID").Value = id
    Cel(mLog, n, "Note").Value = SafeText(note)
End Sub

Private Sub ShowHistoryFor(ByVal id As Long)
    Dim hits As Long, r As Long
    ClearFilter mLog
    If Not mLog.DataBodyRange Is Nothing Then
        hits = Application.WorksheetFunction.CountIf(mLog.ListColumns("ID").DataBodyRange, id)
    End If
    If hits = 0 Then
        Say "#" & id & " has no log lines yet."
        Exit Sub
    End If
    mLog.Range.AutoFilter Field:=mLog.ListColumns("ID").Index, Criteria1:="=" & id
    mLog.Parent.Activate
    For r = mLog.ListRows.Count To 1 Step -1
        If Not mLog.ListRows(r).Range.EntireRow.Hidden Then
            Cel(mLog, r, "Note").Select
            Exit For
        End If
    Next
End Sub

Private Sub ClearFilter(ByVal lo As ListObject)
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
End Sub

Private Function IsOpenRow(ByVal n As Long) As Boolean
    IsOpenRow = Txt(Cel(mItems, n, "What").Value) <> "" And Txt(Cel(mItems, n, "Status").Value) <> ST_DONE
End Function

Private Function IsTheirs(ByVal n As Long) As Boolean
    ' Someone else's return: a Return with a name under Who
    Dim w As String
    w = LCase$(Txt(Cel(mItems, n, "Who").Value))
    IsTheirs = Txt(Cel(mItems, n, "Kind").Value) = KIND_RETURN And w <> "" And w <> "me"
End Function

Private Function StatusFor(ByVal who As Variant) As String
    ' Nobody named means the next move is yours
    Dim w As String
    w = LCase$(Txt(who))
    If w = "" Or w = "me" Then StatusFor = ST_ME Else StatusFor = ST_THEM
End Function

Private Function Head(ByVal n As Long) As String
    ' A row as Today and the reminder box name it: the state, then what it is, twelve words at most
    Dim words As Variant, k As Long, s As String
    s = StateFirst(Txt(Cel(mItems, n, "What").Value), Txt(Cel(mItems, n, "State").Value))
    words = Split(Application.WorksheetFunction.Trim(s), " ")
    If UBound(words) + 1 > HEAD_WORDS Then
        s = ""
        For k = 0 To HEAD_WORDS - 1
            s = s & words(k) & " "
        Next
        s = Trim$(s) & "..."
    End If
    Head = s
End Function

Private Function StateFirst(ByVal what As String, ByVal st As String) As String
    ' The state, then what it is, unless that already starts with the state
    If st = "" Or LCase$(Left$(what, Len(st) + 1)) = LCase$(st & " ") Then
        StateFirst = what
    Else
        StateFirst = st & " " & what
    End If
End Function


' ---- repeats -------------------------------------------------------------------

Private Function ReadRepeat(ByVal s As String, unit As String, every As Long, anchor As String) As String
    ' What a Repeat cell says, in the list's words: "" for none, "?" when it can't tell. Sets
    ' unit (d days, wd weekdays, w weeks, m months), every (how many), and anchor (the day of
    ' the month for a monthly one, when it says "on the 31st" or "on the last day").
    Dim at As Long, tail As String, words As Variant, num As Long, u As String
    unit = ""
    every = 0
    anchor = ""
    s = LCase$(Trim$(Replace(Replace(s, "-", " "), ",", " ")))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    at = InStr(s, " on ")
    If at > 0 Then
        tail = Trim$(Mid$(s, at + 4))
        s = Trim$(Left$(s, at - 1))
        If Left$(tail, 4) = "the " Then tail = Mid$(tail, 5)
        If tail = "last" Or tail = "last day" Then
            anchor = "last"
        ElseIf Val(tail) >= 1 And Val(tail) <= 31 And tail Like "#*" Then
            anchor = CStr(CLng(Val(tail)))
        Else
            ReadRepeat = "?"
            Exit Function
        End If
    End If
    Select Case s
        Case "", "none", "no", "never", "once"
        Case "daily", "every day", "each day", "everyday", "day"
            unit = "d": every = 1
        Case "weekdays", "weekday", "every weekday", "workdays", "every workday", "business days", _
             "every business day", "mon fri", "monday friday", "monday to friday", "monday thru friday"
            unit = "wd": every = 1
        Case "weekly", "every week", "each week", "once a week", "week"
            unit = "w": every = 1
        Case "biweekly", "bi weekly", "every other week", "every two weeks", "fortnightly"
            unit = "w": every = 2
        Case "monthly", "every month", "each month", "once a month", "month"
            unit = "m": every = 1
        Case "quarterly", "every quarter", "once a quarter", "quarter"
            unit = "m": every = 3
        Case "semi annual", "semiannual", "semi annually", "semiannually", "twice a year"
            unit = "m": every = 6
        Case "annual", "annually", "yearly", "every year", "once a year", "year"
            unit = "m": every = 12
        Case Else
            ' every 3 days, every 6 weeks, every 2 months, every 2 years, or the same without "every"
            words = Split(s, " ")
            If words(0) = "every" Then s = Trim$(Mid$(s, 7))
            words = Split(s, " ")
            If UBound(words) = 1 Then
                num = NumberWord(words(0))
                u = words(1)
                If Right$(u, 1) = "s" Then u = Left$(u, Len(u) - 1)
                Select Case u
                    Case "day": unit = "d"
                    Case "week", "wk": unit = "w"
                    Case "month", "mo": unit = "m"
                    Case "year", "yr": unit = "m": num = num * 12
                End Select
                If num < 1 Or num > 999 Then unit = ""
                every = num
            End If
            If unit = "" Then
                ReadRepeat = "?"
                Exit Function
            End If
    End Select
    If unit = "" Then
        If anchor <> "" Then ReadRepeat = "?"
        Exit Function
    End If
    If anchor <> "" And unit <> "m" Then          ' only a monthly one can have a day of the month
        ReadRepeat = "?"
        Exit Function
    End If
    ReadRepeat = RepeatName(unit, every) & AnchorWords(anchor)
End Function

Private Function NumberWord(ByVal w As String) As Long
    Select Case w
        Case "other", "two": NumberWord = 2
        Case "three": NumberWord = 3
        Case "four": NumberWord = 4
        Case "five": NumberWord = 5
        Case "six": NumberWord = 6
        Case Else
            If IsNumber(w) Then NumberWord = CLng(w)
    End Select
End Function

Private Function RepeatName(ByVal unit As String, ByVal every As Long) As String
    Select Case unit
        Case "d": If every = 1 Then RepeatName = "daily" Else RepeatName = "every " & every & " days"
        Case "wd": RepeatName = "weekdays"
        Case "w": If every = 1 Then RepeatName = "weekly" Else RepeatName = "every " & every & " weeks"
        Case "m"
            Select Case every
                Case 1: RepeatName = "monthly"
                Case 3: RepeatName = "quarterly"
                Case 6: RepeatName = "semi-annual"
                Case 12: RepeatName = "annual"
                Case Else: RepeatName = "every " & every & " months"
            End Select
    End Select
End Function

Private Function AnchorWords(ByVal anchor As String) As String
    If anchor = "last" Then
        AnchorWords = " on the last day"
    ElseIf anchor <> "" Then
        AnchorWords = " on the " & Ordinal(CLng(anchor))
    End If
End Function

Private Function KeepAnchor(ByVal d As Long, ByVal anchor As String) As String
    ' The day a monthly row stays on. It follows Due whenever the two disagree, so a due date
    ' changed by hand moves it. It's only written out for the 29th to the 31st (or "last"),
    ' where a short month would otherwise move it for good: a 31st stays the 31st after
    ' February's 28th.
    If anchor = "" Or DueIn(CDate(d), anchor) <> d Then anchor = CStr(Day(CDate(d)))
    If anchor <> "last" Then
        If Val(anchor) <= 28 Then anchor = ""
    End If
    KeepAnchor = anchor
End Function

Private Function Ordinal(ByVal n As Long) As String
    Select Case n Mod 100
        Case 11, 12, 13: Ordinal = n & "th"
        Case Else
            Select Case n Mod 10
                Case 1: Ordinal = n & "st"
                Case 2: Ordinal = n & "nd"
                Case 3: Ordinal = n & "rd"
                Case Else: Ordinal = n & "th"
            End Select
    End Select
End Function

Private Function StepDue(ByVal d As Double, ByVal unit As String, ByVal every As Long, ByVal anchor As String) As Double
    ' One step on from d, keeping its time of day. Weekdays skip Saturday and Sunday; months
    ' land on the anchor day, or the month's last day when it's short.
    Dim dayNo As Long, t As Double
    dayNo = Int(d)
    t = d - dayNo
    Select Case unit
        Case "d": dayNo = dayNo + every
        Case "w": dayNo = dayNo + 7 * every
        Case "wd"
            dayNo = dayNo + 1
            Do While Weekday(CDate(dayNo), vbMonday) > 5
                dayNo = dayNo + 1
            Loop
        Case "m": dayNo = Advance(dayNo, every, anchor)
        Case Else: dayNo = dayNo + 1
    End Select
    StepDue = dayNo + t
End Function

Private Function Advance(ByVal d As Long, ByVal months As Long, ByVal dueDay As String) As Long
    ' The date that many months on, on the due day
    Advance = DueIn(DateSerial(Year(CDate(d)), Month(CDate(d)) + months, 1), dueDay)
End Function

Private Function DueIn(ByVal anyDay As Date, ByVal dueDay As String) As Long
    ' The due day's date in that day's month. "last", or a day the month doesn't have (the
    ' 31st in April), means the month's last day.
    Dim lastDay As Long, n As Long
    lastDay = Day(DateSerial(Year(anyDay), Month(anyDay) + 1, 0))
    If Val(dueDay) >= 1 And Val(dueDay) <= lastDay Then n = Val(dueDay) Else n = lastDay
    DueIn = CLng(DateSerial(Year(anyDay), Month(anyDay), n))
End Function

Private Function OldMonths(ByVal s As String) As Long
    ' The old Calendar's How often, as months
    Select Case LCase$(Replace(Trim$(s), "-", " "))
        Case "monthly", "month", "every month", "once a month": OldMonths = 1
        Case "quarterly", "quarter", "every quarter", "once a quarter": OldMonths = 3
        Case "semi annual", "semi annually", "semiannual", "semiannually", "twice a year", "every 6 months": OldMonths = 6
        Case "annual", "annually", "yearly", "year", "every year", "once a year": OldMonths = 12
    End Select
End Function

Private Function KindName(ByVal s As Variant) As String
    ' Kind in the list's words. Blank is an Item; "" when it can't tell.
    Select Case LCase$(Txt(s))
        Case "", "i", "item", "items", "thing": KindName = KIND_ITEM
        Case "rem", "remind", "reminder", "reminders": KindName = KIND_REMINDER
        Case "ret", "return", "returns", "filing": KindName = KIND_RETURN
    End Select
End Function


' ---- the month -------------------------------------------------------------------

Private Sub MonthButton(ByVal stepBy As Long)
    ' The Month sheet's buttons: this month (0), or a month back (-1) or on (1) from the one showing
    Dim shown As Date
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not SheetExists(MONTH_SHEET) Then
        Say "Run TrackerSetup first: Alt F8, pick TrackerSetup, Run. It adds the Month sheet."
        Exit Sub
    End If
    If stepBy = 0 Then shown = Date Else shown = DateAdd("m", stepBy, MonthShown())
    Application.ScreenUpdating = False
    Rework shown
    Application.ScreenUpdating = True
    GoMonth
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Month stopped: " & Err.Description
End Sub

Private Function MonthShown() As Date
    ' The first of the month the Month sheet shows, which its title keeps. This month if it has none.
    Dim d As Long
    If SheetExists(MONTH_SHEET) Then d = DayNo(ThisWorkbook.Worksheets(MONTH_SHEET).Range("A2").Value)
    If d = 0 Then d = CLng(Date)
    MonthShown = DateSerial(Year(CDate(d)), Month(CDate(d)), 1)
End Function

Private Sub BuildMonth(ByVal anyDay As Date)
    ' Draws the Month sheet fresh for anyDay's month: a box for each day, Sunday to Saturday.
    ' The weeks are whole, so the first and last show a few days of the months around it.
    Dim ws As Worksheet, boxes As New Collection, counts As Variant
    Dim first As Long, lastDay As Long, start As Long, finish As Long, d As Long
    If Not SheetExists(MONTH_SHEET) Then Exit Sub          ' TrackerSetup adds it
    Set ws = ThisWorkbook.Worksheets(MONTH_SHEET)
    first = CLng(DateSerial(Year(anyDay), Month(anyDay), 1))
    lastDay = CLng(DateSerial(Year(anyDay), Month(anyDay) + 1, 0))
    start = first - Weekday(first, vbSunday) + 1            ' the Sunday on or before the 1st
    finish = lastDay + 7 - Weekday(lastDay, vbSunday)       ' the Saturday on or after the last day
    For d = start To finish
        boxes.Add New Collection
    Next
    counts = Array(0, 0, 0, 0, 0)       ' 1 yours, 2 done, 3 late, 4 waiting on someone else
    FillBoxes boxes, start, first, lastDay, counts
    DrawMonth ws, boxes, start, first, lastDay
    ws.Range("A2").Value = CDate(first)
    ws.Range("B2").Value = counts(1) & " of yours (" & counts(2) & " done, " & counts(3) & " late), " & _
        counts(4) & " waiting on someone else."
    ws.Range("F2").Value = Now
End Sub

Private Sub FillBoxes(ByVal boxes As Collection, ByVal start As Long, ByVal first As Long, _
                      ByVal lastDay As Long, counts As Variant)
    ' Every row on its due date, and every open row that repeats on its next dates too, through
    ' Until. Yours are black, red once late, crossed out once done; anything waiting on someone
    ' else is gray with their name.
    Dim i As Long, n As Long, d As Long, k As Long, finish As Long, today As Long, untilDay As Long
    Dim due As Double, nxt As Double, what As String, st As String, who As String, style As String
    Dim unit As String, every As Long, anchor As String, closed As Boolean, theirs As Boolean
    finish = start + boxes.Count - 1
    today = CLng(Date)
    n = mItems.ListRows.Count
    For i = 1 To n
        what = Txt(Cel(mItems, i, "What").Value)
        due = DateOf(Cel(mItems, i, "Due").Value)
        If what <> "" And due > 0 Then
            st = Txt(Cel(mItems, i, "State").Value)
            who = Txt(Cel(mItems, i, "Who").Value)
            closed = Txt(Cel(mItems, i, "Status").Value) = ST_DONE
            theirs = Txt(Cel(mItems, i, "Status").Value) = ST_THEM Or IsTheirs(i)
            d = Int(due)
            If closed Then
                style = "F"
            ElseIf d < today And Not IsTheirs(i) Then
                style = "R"
            ElseIf theirs Then
                style = "G"
            Else
                style = "N"
            End If
            If d >= start And d <= finish Then AddEntry boxes(d - start + 1), style, due, MonthLabel(what, st, due, theirs, who)
            If d >= first And d <= lastDay Then Tally counts, style, theirs
            If Not closed Then
                If ReadRepeat(Txt(Cel(mItems, i, "Repeat").Value), unit, every, anchor) <> "" And unit <> "" Then
                    If anchor = "" Then anchor = CStr(Day(CDate(d)))
                    untilDay = DayNo(Cel(mItems, i, "Until").Value)
                    If theirs Then style = "G" Else style = "N"
                    nxt = StepDue(due, unit, every, anchor)
                    k = 0
                    Do While Int(nxt) <= finish And k < 3000
                        If untilDay > 0 And Int(nxt) > untilDay Then Exit Do
                        d = Int(nxt)
                        If d >= start And d >= today Then
                            AddEntry boxes(d - start + 1), style, nxt, MonthLabel(what, st, nxt, theirs, who)
                            If d >= first And d <= lastDay Then Tally counts, style, theirs
                        End If
                        nxt = StepDue(nxt, unit, every, anchor)
                        k = k + 1
                    Loop
                End If
            End If
        End If
    Next
End Sub

Private Function MonthLabel(ByVal what As String, ByVal st As String, ByVal due As Double, _
                            ByVal theirs As Boolean, ByVal who As String) As String
    ' A row as a box shows it: the time if it has one, the state, what it is, and whose it is
    MonthLabel = StateFirst(Short(what), st)
    If HasTime(due) Then MonthLabel = Format$(CDate(due), "h:mm AM/PM") & " " & MonthLabel
    If theirs And who <> "" Then MonthLabel = MonthLabel & " (" & who & ")"
End Function

Private Sub Tally(counts As Variant, ByVal style As String, ByVal theirs As Boolean)
    ' Counts a line in the month's summary
    If theirs Then
        counts(4) = counts(4) + 1
    Else
        counts(1) = counts(1) + 1
        If style = "F" Then counts(2) = counts(2) + 1
        If style = "R" Then counts(3) = counts(3) + 1
    End If
End Sub

Private Sub AddEntry(ByVal list As Collection, ByVal style As String, ByVal due As Double, ByVal words As String)
    ' Adds a line to a day's box, kept in order: late (R) first, then yours (N), waiting on
    ' someone else (G), and done (F); in each, no time first, then by time, then A to Z. An
    ' entry is its style letter, four sorting characters, then its words.
    Dim k As Long, e As String
    If HasTime(due) Then
        e = style & Format$(Round((due - Int(due)) * 1440), "0000") & words
    Else
        e = style & "----" & words
    End If
    For k = 1 To list.Count
        If EntryKey(e) < EntryKey(list(k)) Then
            list.Add e, Before:=k
            Exit Sub
        End If
    Next
    list.Add e
End Sub

Private Function EntryKey(ByVal entry As String) As String
    EntryKey = InStr("RNGF", Left$(entry, 1)) & LCase$(Mid$(entry, 2))
End Function

Private Sub DrawMonth(ByVal ws As Worksheet, ByVal boxes As Collection, ByVal start As Long, _
                      ByVal first As Long, ByVal lastDay As Long)
    ' Writes each day's box: the day on top in bold, then its lines, each in its own color
    Dim grid As Range, box As Range, list As Collection, entry As Variant
    Dim weeks As Long, j As Long, d As Long, k As Long, at As Long, head As String, body As String
    weeks = boxes.Count \ 7
    With ws.Range(ws.Cells(MONTH_TOP + 1, 1), ws.Cells(MONTH_TOP + 6, 7))
        .Clear
        .RowHeight = ws.StandardHeight
    End With
    Set grid = ws.Range(ws.Cells(MONTH_TOP + 1, 1), ws.Cells(MONTH_TOP + weeks, 7))
    With grid
        .NumberFormat = "@"                 ' text, so a day number or "Oct 1" never turns into a date
        .Font.Size = 10
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Color = LINE_COLOR
    End With
    For j = 1 To boxes.Count
        d = start + j - 1
        Set list = boxes(j)
        Set box = grid.Cells(1 + (j - 1) \ 7, 1 + (j - 1) Mod 7)
        If Day(CDate(d)) = 1 Then head = Format$(CDate(d), "mmm d") Else head = CStr(Day(CDate(d)))
        body = head
        k = 0
        For Each entry In list
            k = k + 1
            If k > DAY_MAX And list.Count > DAY_MAX + 1 Then
                body = body & vbLf & "and " & (list.Count - DAY_MAX) & " more"
                Exit For
            End If
            body = body & vbLf & Mid$(entry, 6)
        Next
        box.Value = body
        box.Errors(xlNumberAsText).Ignore = True     ' no green corner on a box that's just its day number
        box.Characters(1, Len(head)).Font.Bold = True
        at = Len(head) + 2
        k = 0
        For Each entry In list
            k = k + 1
            If k > DAY_MAX And list.Count > DAY_MAX + 1 Then
                StyleLine box, at, Len(body) - at + 1, "G"
                Exit For
            End If
            StyleLine box, at, Len(entry) - 5, Left$(entry, 1)
            at = at + Len(entry) - 4             ' the line and its line break
        Next
        If d = CLng(Date) Then
            box.Interior.Color = YELLOW_FILL
        ElseIf d < first Or d > lastDay Then
            box.Interior.Color = AWAY_FILL
            box.Characters(1, Len(head)).Font.Color = GRAY_TEXT
        ElseIf Weekday(CDate(d), vbMonday) > 5 Then
            box.Interior.Color = WEEKEND_FILL
        End If
    Next
    For k = 1 To weeks
        With grid.Rows(k)
            .EntireRow.AutoFit
            If .RowHeight < WEEK_HEIGHT Then .RowHeight = WEEK_HEIGHT
        End With
    Next
End Sub

Private Sub StyleLine(ByVal box As Range, ByVal at As Long, ByVal length As Long, ByVal style As String)
    ' Colors one line of a box by its letter: R red, F crossed out, G gray, N as is
    If length < 1 Then Exit Sub
    With box.Characters(at, length).Font
        Select Case style
            Case "R"
                .Color = RED_TEXT
                .Bold = True
            Case "F"
                .Color = GRAY_TEXT
                .Strikethrough = True
            Case "G"
                .Color = GRAY_TEXT
        End Select
    End With
End Sub


' ---- building the sheets -------------------------------------------------------

Private Function EnsureSheet(ByVal nm As String, ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        ' A brand-new workbook's empty sheet becomes the Items sheet instead of sitting there blank
        If after Is Nothing And ThisWorkbook.Worksheets.Count = 1 Then
            If Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets(1).Cells) = 0 And _
               ThisWorkbook.Worksheets(1).ListObjects.Count = 0 Then Set ws = ThisWorkbook.Worksheets(1)
        End If
        If ws Is Nothing Then
            If after Is Nothing Then
                Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
            Else
                Set ws = ThisWorkbook.Worksheets.Add(After:=after)
            End If
        End If
        ws.Name = nm
    End If
    Set EnsureSheet = ws
End Function

Private Function EnsureTable(ByVal ws As Worksheet, ByVal nm As String, ByVal cols As String) As ListObject
    Dim lo As ListObject, heads As Variant, i As Long
    heads = Split(cols, "|")
    Set lo = GetTable(nm)
    If lo Is Nothing Then
        For i = 0 To UBound(heads)
            ws.Cells(HEAD_ROW, i + 1).Value = heads(i)
        Next
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(HEAD_ROW, 1), ws.Cells(HEAD_ROW, UBound(heads) + 1)), , xlYes)
        lo.Name = nm
        If lo.ListRows.Count = 0 Then lo.ListRows.Add    ' a blank first row for the first item to fill
    Else
        ' A newer copy of this module can bring new columns. They go on the end.
        For i = 0 To UBound(heads)
            If Not HasColumn(lo, heads(i)) Then lo.ListColumns.Add.Name = heads(i)
        Next
    End If
    lo.TableStyle = "TableStyleMedium2"
    Set EnsureTable = lo
End Function

Private Function HasColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If LCase$(lc.Name) = LCase$(colName) Then
            HasColumn = True
            Exit Function
        End If
    Next
End Function

Private Function HasColumns(ByVal lo As ListObject, ByVal cols As String) As Boolean
    Dim c As Variant
    For Each c In Split(cols, "|")
        If Not HasColumn(lo, c) Then Exit Function
    Next
    HasColumns = True
End Function

Private Sub DressItems(ByVal ws As Worksheet, ByVal lo As ListObject)
    ws.Rows(1).RowHeight = 30
    ws.Range("A2").Value = "One row for everything you're tracking. Any row can repeat (Repeat) and pop up a " & _
        "reminder (Remind, a time of day) while this file is open. Who: who has the next move, blank if it's " & _
        "you; on a return, who files it. Red: late. Yellow: due today. Orange: no due date. Gray: done."
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = HELP_TEXT
    SetWidths ws, lo, "ID:6|Kind:10|What:42|State:8|Who:16|Status:16|Due:22|Repeat:18|Until:11|Remind:11|" & _
        "Quiet until:18|Last updated:13|Latest note:44|Reference:18|Opened:10|Closed:10"
    Freshen
    AddColors ws, lo
    AddButton ws, "trkNew", "New item", "NewItem", 0
    AddButton ws, "trkLog", "Update item", "UpdateItem", 1
    AddButton ws, "trkClose", "Close item", "CloseItem", 2
    AddButton ws, "trkHist", "History", "ShowHistory", 3
    AddButton ws, "trkToday", "Today", "RefreshToday", 4
    AddButton ws, "trkImport", "Import old", "ImportOld", 5
    FreezeTop ws, HEAD_ROW
End Sub

Private Sub DressLog(ByVal ws As Worksheet, ByVal lo As ListObject)
    ws.Rows(1).RowHeight = 30
    ws.Range("A2").Value = "One line for every call, email, or anything else that happens. Newest at the bottom."
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = HELP_TEXT
    SetWidths ws, lo, "When:17|ID:6|Item:42|Note:80"
    AddButton ws, "trkBack", "Back to items", "BackToItems", 0
    AddButton ws, "trkLog2", "Update item", "UpdateItem", 1
    FreezeTop ws, HEAD_ROW
End Sub

Private Sub DressMonth(ByVal ws As Worksheet)
    ' The month's name, a line of what it has, the Updated stamp, the key, and the day names.
    ' BuildMonth draws the weeks under them.
    Dim fc As FormatCondition, dayNames As Variant, k As Long
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 24
    ws.Range("A:G").ColumnWidth = 24
    With ws.Range("A2")
        .NumberFormat = "mmmm yyyy"
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Size = 16
    End With
    ws.Range("A2:G2").VerticalAlignment = xlCenter
    ws.Range("F2:G2").Merge
    With ws.Range("F2")
        .NumberFormat = """Updated ""dddd m/d h:mm AM/PM"
        .HorizontalAlignment = xlRight
        .Font.Italic = True
        .Font.Color = HELP_TEXT
        .FormatConditions.Delete
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=INT($F$2)<TODAY()")   ' drawn on an earlier day
        fc.Interior.Color = RED_FILL
        fc.Font.Color = RED_TEXT
    End With
    MonthKey ws.Range("A3")
    dayNames = Split("Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", ",")
    For k = 0 To 6
        ws.Cells(MONTH_TOP, k + 1).Value = dayNames(k)
    Next
    With ws.Range(ws.Cells(MONTH_TOP, 1), ws.Cells(MONTH_TOP, 7))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = HEAD_FILL
        .HorizontalAlignment = xlCenter
    End With
    AddButton ws, "trkPrev", "Previous month", "PreviousMonth", 0
    AddButton ws, "trkThis", "This month", "ShowMonth", 1
    AddButton ws, "trkNext", "Next month", "NextMonth", 2
    AddButton ws, "trkToday", "Today", "RefreshToday", 3
    FreezeTop ws, MONTH_TOP
    SetPrint ws
End Sub

Private Sub MonthKey(ByVal c As Range)
    ' The key over the day names, each sample written the way the boxes show it
    Dim parts As Variant, looks As Variant, k As Long, at As Long
    parts = Array("Key:", "Yours", "Waiting on someone (their name)", "Late", "Done", _
                  "This sheet redraws itself, so type on Items.")
    looks = Array("", "N", "G", "R", "F", "")
    c.Clear
    c.Value = Join(parts, "     ")
    c.Font.Size = 10
    c.Characters(1, Len(parts(0))).Font.Bold = True
    at = 1
    For k = 0 To UBound(parts)
        StyleLine c, at, Len(parts(k)), looks(k)
        at = at + Len(parts(k)) + 5
    Next
    With c.Characters(at - 5 - Len(parts(UBound(parts))), Len(parts(UBound(parts)))).Font
        .Italic = True
        .Color = HELP_TEXT
    End With
End Sub

Private Sub SetPrint(ByVal ws As Worksheet)
    ' The month prints on one page, sideways. A PC with no printer set up can't take page
    ' settings, so any trouble here is skipped.
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(2, 1), ws.Cells(MONTH_TOP + 6, 7)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
End Sub

Private Sub DressToday(ByVal ws As Worksheet)
    Dim fc As FormatCondition, widths As Variant, k As Long
    ws.Rows(1).RowHeight = 30
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12
    ws.Range("A3:B3").Merge
    With ws.Range("A3")
        .NumberFormat = """Updated ""dddd m/d h:mm AM/PM"
        .HorizontalAlignment = xlLeft
        .Font.Italic = True
        .Font.Color = HELP_TEXT
        .FormatConditions.Delete
        Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=INT($A$3)<TODAY()")   ' built on an earlier day
        fc.Interior.Color = RED_FILL
        fc.Font.Color = RED_TEXT
    End With
    ws.Range("C3").Value = "This sheet rebuilds itself, so type on Items. Reminders pop up while this file is open."
    ws.Range("C3").Font.Italic = True
    ws.Range("C3").Font.Color = HELP_TEXT
    widths = Array(6, 58, 20, 16, 10, 56)          ' one for each Today column
    For k = 0 To UBound(widths)
        ws.Columns(k + 1).ColumnWidth = widths(k)
    Next
    AddButton ws, "trkRefresh", "Refresh", "RefreshToday", 0
    AddButton ws, "trkNew", "New item", "NewItem", 1
    AddButton ws, "trkLog", "Update item", "UpdateItem", 2
    AddButton ws, "trkClose", "Close item", "CloseItem", 3
    AddButton ws, "trkHist", "History", "ShowHistory", 4
    AddButton ws, "trkMonth", "Month view", "ShowMonth", 5
    FreezeTop ws, TODAY_TOP
End Sub

Private Sub BuildToday()
    ' Writes the Today sheet fresh from Items. Each open row lands in one block: late, today,
    ' the next 7 days, gone quiet, waiting on them, or no date. Someone else's returns stay off.
    ' A row reads: its head, when it's due, then Yours in bold when the move is yours.
    Dim ws As Worksheet, cols As Variant, titles As Variant, fills As Variant, inks As Variant
    Dim lists(1 To 6) As Collection, order As Variant, idx As Variant, i As Long, g As Long, k As Long, r As Long
    Dim wide As Long, due As Long, lastDay As Long, today As Long, st As String, who As String
    If Not SheetExists(TODAY_SHEET) Then Exit Sub          ' TrackerSetup adds it
    Set ws = ThisWorkbook.Worksheets(TODAY_SHEET)
    cols = Split(TODAY_COLS, "|")
    wide = UBound(cols) + 1
    titles = Array("", "Late", "Today", "Next " & LEAD_DAYS & " days", "Gone quiet, no update in " & QUIET_DAYS & " days", _
                   "Waiting on them", "No date")
    fills = Array(0, RED_FILL, YELLOW_FILL, SOON_FILL, QUIET_FILL, WAIT_FILL, ORANGE_FILL)
    inks = Array(0, RED_TEXT, YELLOW_TEXT, vbBlack, vbBlack, vbBlack, vbBlack)
    order = Array(1, 2, 3, 5, 4, 6)                       ' waiting on them shows before gone quiet
    today = CLng(Date)
    For g = 1 To 6
        Set lists(g) = New Collection
    Next
    For i = 1 To mItems.ListRows.Count
        If IsOpenRow(i) And Not IsTheirs(i) Then
            due = DayNo(Cel(mItems, i, "Due").Value)
            lastDay = DayNo(Cel(mItems, i, "Last updated").Value)
            st = Txt(Cel(mItems, i, "Status").Value)
            If due > 0 And due < today Then
                g = 1
            ElseIf due = today Then
                g = 2
            ElseIf due > 0 And due <= today + LEAD_DAYS Then
                g = 3
            ElseIf lastDay > 0 And lastDay <= today - QUIET_DAYS Then
                g = 4
            ElseIf st = ST_THEM Then
                g = 5
            ElseIf due = 0 Then
                g = 6
            Else
                g = 0                                     ' due later, and recently updated
            End If
            If g > 0 Then lists(g).Add i
        End If
    Next

    ws.Range(ws.Cells(TODAY_TOP, 1), ws.Cells(ws.Rows.Count, wide)).Clear
    For k = 0 To UBound(cols)
        ws.Cells(TODAY_TOP, k + 1).Value = cols(k)
    Next
    With ws.Range(ws.Cells(TODAY_TOP, 1), ws.Cells(TODAY_TOP, wide))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = HEAD_FILL
    End With
    r = TODAY_TOP + 1
    For k = 0 To UBound(order)
        g = order(k)
        If lists(g).Count > 0 Then
            ws.Cells(r, 1).Value = titles(g) & ": " & lists(g).Count
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, wide))
                .Font.Bold = True
                .Font.Color = inks(g)
                .Interior.Color = fills(g)
            End With
            r = r + 1
            For Each idx In lists(g)
                ws.Cells(r, 1).Value = Cel(mItems, idx, "ID").Value
                ws.Cells(r, 2).Value = SafeText(Head(idx))
                ws.Cells(r, 3).Value = DueText(DateOf(Cel(mItems, idx, "Due").Value))
                st = Txt(Cel(mItems, idx, "Status").Value)
                who = Txt(Cel(mItems, idx, "Who").Value)
                If st = ST_ME Then
                    ws.Cells(r, 4).Value = "Yours"
                    ws.Cells(r, 4).Font.Bold = True
                ElseIf who <> "" Then
                    ws.Cells(r, 4).Value = SafeText(who)
                Else
                    ws.Cells(r, 4).Value = "Them"
                End If
                ws.Cells(r, 5).Value = Txt(Cel(mItems, idx, "Kind").Value)
                ws.Cells(r, 6).Value = SafeText(Short(Cel(mItems, idx, "Latest note").Value))
                ws.Range(ws.Cells(r, 1), ws.Cells(r, wide)).Font.Color = inks(g)
                r = r + 1
            Next
            r = r + 1
        End If
    Next
    If r = TODAY_TOP + 1 Then
        ws.Cells(r, 1).Value = "All clear. Nothing is late, due this week, waiting on anyone, or gone quiet."
    End If
    ws.Range(ws.Cells(TODAY_TOP + 1, 3), ws.Cells(r, 3)).NumberFormat = "@"
    ws.Range("A2").Value = Format$(Date, "dddd, mmmm d") & ": " & lists(1).Count & " late, " & lists(2).Count & _
        " today, " & lists(3).Count & " in the next " & LEAD_DAYS & " days, " & lists(5).Count & " waiting on them, " & _
        lists(4).Count & " gone quiet" & IIf(lists(6).Count > 0, ", " & lists(6).Count & " with no date", "") & "."
    ws.Range("A3").Value = Now
End Sub

Private Sub GoToday(Optional ByVal id As Long = 0)
    ' Shows the Today sheet, on the item's row when it's listed there
    Dim ws As Worksheet, hit As Variant
    If Not SheetExists(TODAY_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(TODAY_SHEET)
    ws.Activate
    If id > 0 Then
        hit = Application.Match(CDbl(id), ws.Range(ws.Cells(TODAY_TOP + 1, 1), ws.Cells(ws.Rows.Count, 1)), 0)
        If Not IsError(hit) Then
            ws.Cells(TODAY_TOP + hit, 2).Select
            Exit Sub
        End If
    End If
    ws.Cells(TODAY_TOP + 1, 2).Select
End Sub

Private Sub GoMonth()
    ' Shows the Month sheet, on today's box when this month is showing, or else the 1st
    Dim ws As Worksheet, first As Long, start As Long, d As Long
    If Not SheetExists(MONTH_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(MONTH_SHEET)
    ws.Activate
    first = CLng(MonthShown())
    start = first - Weekday(first, vbSunday) + 1
    d = CLng(Date)
    If Year(CDate(d)) <> Year(CDate(first)) Or Month(CDate(d)) <> Month(CDate(first)) Then d = first
    ws.Cells(MONTH_TOP + 1 + (d - start) \ 7, 1 + (d - start) Mod 7).Select
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Private Sub Freshen()
    ' Date formats and dropdowns, laid again on every button press so a row inserted by
    ' hand gets them too
    Dim ws As Worksheet
    Set ws = mItems.Parent
    ColSpan(ws, mItems, "Opened").NumberFormat = "m/d/yy"
    ColSpan(ws, mItems, "Due").NumberFormat = "ddd m/d/yy"      ' a time shows too, by a color rule
    ColSpan(ws, mItems, "Until").NumberFormat = "m/d/yy"
    ColSpan(ws, mItems, "Remind").NumberFormat = "h:mm AM/PM"
    ColSpan(ws, mItems, "Quiet until").NumberFormat = "ddd m/d h:mm AM/PM"
    ColSpan(ws, mItems, "Last updated").NumberFormat = "m/d/yy;;"
    ColSpan(ws, mItems, "Closed").NumberFormat = "m/d/yy"
    ColSpan(ws, mItems, "Reference").NumberFormat = "@"       ' as text, so long numbers and leading zeros stay whole
    AddDropdown ws, mItems, "Status", ST_THEM & "," & ST_ME & "," & ST_DONE, True
    AddDropdown ws, mItems, "Kind", KIND_ITEM & "," & KIND_REMINDER & "," & KIND_RETURN, True
    AddDropdown ws, mItems, "State", STATE_LIST, False
    AddDropdown ws, mItems, "Repeat", REPEAT_LIST, False
    NoCorners                                                  ' laying the dropdown again brings them back
    DateOnly ColSpan(ws, mItems, "Due"), "Type a date, like 10/15, or a date and time, like 10/15 2:30 PM."
    DateOnly ColSpan(ws, mItems, "Until"), "Type a date, like 12/31."
    With ColSpan(ws, mItems, "Remind").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="0.99999"
        .ErrorMessage = "Type a time of day, like 9:00 AM or 2:30 PM."
        .IgnoreBlank = True
    End With
    mItems.ListColumns("Sort").Range.EntireColumn.Hidden = True
    ColSpan(mLog.Parent, mLog, "When").NumberFormat = "m/d/yy h:mm AM/PM"
    DateOnly ColSpan(mLog.Parent, mLog, "When"), "Type a date, like 9/18, or a date and time, like 9/18 2:30 PM."
End Sub

Private Sub DateOnly(ByVal target As Range, ByVal msg As String)
    ' A date typed as text never counts, so the column turns text away as it's typed
    With target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="36526"
        .ErrorMessage = msg
        .IgnoreBlank = True
    End With
End Sub

Private Function ColSpan(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal colName As String) As Range
    ' The column from under its header down past where the table will ever reach
    Dim c As Long
    c = lo.ListColumns(colName).Range.Column
    Set ColSpan = ws.Range(ws.Cells(HEAD_ROW + 1, c), ws.Cells(HEAD_ROW + SPAN, c))
End Function

Private Sub SetWidths(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal spec As String)
    Dim part As Variant, bits As Variant
    For Each part In Split(spec, "|")
        bits = Split(part, ":")
        If HasColumn(lo, bits(0)) Then ws.Columns(lo.ListColumns(bits(0)).Range.Column).ColumnWidth = CDbl(bits(1))
    Next
End Sub

Private Sub AddDropdown(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal colName As String, _
                        ByVal choices As String, ByVal strict As Boolean)
    With ColSpan(ws, lo, colName).Validation
        .Delete
        If strict Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=choices
            .ErrorMessage = "Pick one from the list."
        Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=choices
        End If
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = strict
    End With
End Sub

Private Sub AddColors(ByVal ws As Worksheet, ByVal lo As ListObject)
    ' The rules start on the header row, which none of them ever colors, so a row inserted
    ' at the top of the table still falls inside them. The first rule shows the time on a
    ' due date that has one; it comes first because the color rules stop the ones after them.
    Dim area As Range, fc As FormatCondition, s As String, c As String, w As String, dueCol As Range
    Set area = ws.Range(ws.Cells(HEAD_ROW, lo.Range.Column), _
                        ws.Cells(HEAD_ROW + SPAN, lo.Range.Column + lo.ListColumns.Count - 1))
    area.FormatConditions.Delete
    ws.Activate
    area.Cells(1, 1).Select            ' Excel reads these formulas from the selected cell
    s = "$" & ColLetter(ws, lo.ListColumns("Status").Range.Column) & HEAD_ROW
    c = "$" & ColLetter(ws, lo.ListColumns("Due").Range.Column) & HEAD_ROW
    w = "$" & ColLetter(ws, lo.ListColumns("What").Range.Column) & HEAD_ROW
    Set dueCol = ws.Range(ws.Cells(HEAD_ROW, lo.ListColumns("Due").Range.Column), _
                          ws.Cells(HEAD_ROW + SPAN, lo.ListColumns("Due").Range.Column))
    ws.Cells(HEAD_ROW, lo.ListColumns("Due").Range.Column).Select
    Set fc = dueCol.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & c & "),MOD(" & c & ",1)>0)")
    fc.NumberFormat = "ddd m/d/yy h:mm AM/PM"
    fc.StopIfTrue = False
    area.Cells(1, 1).Select
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & s & "=""" & ST_DONE & """")
    fc.Font.Color = GRAY_TEXT
    fc.StopIfTrue = True
    fc.SetLastPriority
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & c & "),INT(" & c & ")<TODAY())")
    fc.Interior.Color = RED_FILL
    fc.Font.Color = RED_TEXT
    fc.StopIfTrue = True
    fc.SetLastPriority
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & c & "),INT(" & c & ")=TODAY())")
    fc.Interior.Color = YELLOW_FILL
    fc.Font.Color = YELLOW_TEXT
    fc.StopIfTrue = True
    fc.SetLastPriority
    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(" & w & "<>""""," & c & "="""")")
    fc.Interior.Color = ORANGE_FILL
    fc.SetLastPriority
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal key As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal slot As Long)
    Dim b As Object
    On Error Resume Next
    ws.Buttons(key).Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Range("A1").Left + 4 + slot * 100, ws.Range("A1").Top + 3, 94, 24)
    b.Name = key
    b.Caption = caption
    b.OnAction = macroName
End Sub

Private Sub FreezeTop(ByVal ws As Worksheet, ByVal rowsAbove As Long)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 0
        .SplitRow = rowsAbove
        .FreezePanes = True
        .DisplayGridlines = False
    End With
End Sub

Private Function ColLetter(ByVal ws As Worksheet, ByVal colNo As Long) As String
    ColLetter = Split(ws.Cells(1, colNo).Address(True, False), "$")(0)
End Function


' ---- reading what was typed ------------------------------------------------------

Private Function WhenFrom(ByVal s As String, ByVal base As Date) As Variant
    ' A due date, with a time when one is typed on the end: 10/15, 10/15 2pm, tomorrow 9am,
    ' 30 (days), 2w, 3m, today, or just 3pm (today, or tomorrow once it's past). Empty when
    ' it can't tell.
    Dim t As Double, rest As String, d As Variant
    s = Squeeze(s)
    If s = "" Then Exit Function
    t = SplitTime(s, rest)
    If rest = "" Then
        If t < 0 Then Exit Function
        d = base
        If CDbl(base) + t < CDbl(Now) Then d = base + 1
    Else
        d = DayFrom(rest, base)
        If IsEmpty(d) Then Exit Function
    End If
    If t >= 0 Then WhenFrom = CDate(CDbl(d) + t) Else WhenFrom = CDate(d)
End Function

Private Function DayFrom(ByVal s As String, ByVal base As Date) As Variant
    ' A day from 10/15, oct 15, 30 (days), 2w, 3m, today or tomorrow. Empty when it can't
    ' tell. A count that lands on a weekend moves to the Monday after. A date typed without
    ' a year that has already passed means next year's.
    Dim num As String, unit As String, d As Date, seps As Long
    s = LCase$(Trim$(s))
    If s = "" Then Exit Function
    If s = "today" Then
        DayFrom = base
        Exit Function
    End If
    If s = "tomorrow" Then
        DayFrom = base + 1
        Exit Function
    End If
    unit = Right$(s, 1)
    If unit Like "[dwm]" Then
        num = Trim$(Left$(s, Len(s) - 1))
    Else
        num = s
        unit = "d"
    End If
    If Len(num) > 0 And Len(num) <= 4 And IsNumber(num) Then
        Select Case unit
            Case "d": d = base + CLng(num)
            Case "w": d = base + 7 * CLng(num)
            Case "m": d = DateAdd("m", CLng(num), base)
        End Select
        DayFrom = OffWeekend(d)
        Exit Function
    End If
    If IsDate(s) Then
        d = DateValue(s)
        seps = Len(s) - Len(Replace(Replace(s, "/", ""), "-", ""))
        If seps < 2 And InStr(s, CStr(Year(d))) = 0 Then
            ' No year typed. VBA reads a day the month doesn't have as a year instead (9/31
            ' comes back 9/1/2031), so a date outside this year is a misread.
            If Year(d) <> Year(base) Then Exit Function
            If d < base Then d = DateSerial(Year(d) + 1, Month(d), Day(d))
        End If
        DayFrom = d
    End If
End Function

Private Function SplitTime(ByVal s As String, rest As String) As Double
    ' The time of day typed on the end of s (2pm, 2:30 pm, 14:30, noon), or -1 when there's
    ' none. rest gets what's left, with a trailing "at" taken off.
    Dim words As Variant, k As Long, j As Long, tok As String, t As Double
    rest = s
    SplitTime = -1
    If s = "" Then Exit Function
    words = Split(s, " ")
    k = UBound(words)
    tok = words(k)
    If tok = "noon" Then
        t = 0.5
    Else
        If k > 0 Then
            Select Case tok
                Case "am", "pm", "a.m.", "p.m."
                    tok = words(k - 1) & tok
                    k = k - 1
            End Select
        End If
        t = TimeFrom(tok)
        If t < 0 Then Exit Function
    End If
    rest = ""
    For j = 0 To k - 1
        rest = rest & " " & words(j)
    Next
    rest = Trim$(rest)
    If rest = "at" Then rest = ""
    If Right$(rest, 3) = " at" Then rest = Trim$(Left$(rest, Len(rest) - 3))
    SplitTime = t
End Function

Private Function TimeFrom(ByVal tok As String) As Double
    ' A time of day from 2pm, 2:30pm, 2:30, 14:30 or 9a, as a part of a day. -1 when it isn't
    ' one. A bare number is a count of days, not a time. Without am or pm, 1:00 to 6:59 is the
    ' afternoon.
    Dim ap As String, hs As String, ms As String, p As Long, h As Long, m As Long
    TimeFrom = -1
    tok = Replace(tok, ".", "")
    If Len(tok) > 2 Then
        If Right$(tok, 2) = "am" Or Right$(tok, 2) = "pm" Then
            ap = Right$(tok, 2)
            tok = Left$(tok, Len(tok) - 2)
        End If
    End If
    If ap = "" And Len(tok) > 1 Then
        If Right$(tok, 1) = "a" Or Right$(tok, 1) = "p" Then
            ap = Right$(tok, 1) & "m"
            tok = Left$(tok, Len(tok) - 1)
        End If
    End If
    p = InStr(tok, ":")
    If p > 0 Then
        hs = Left$(tok, p - 1)
        ms = Mid$(tok, p + 1)
        If Len(ms) <> 2 Or Not IsNumber(ms) Then Exit Function
    Else
        If ap = "" Then Exit Function
        hs = tok
        ms = "00"
    End If
    If Len(hs) < 1 Or Len(hs) > 2 Or Not IsNumber(hs) Then Exit Function
    h = CLng(hs)
    m = CLng(ms)
    If m > 59 Then Exit Function
    If ap <> "" Then
        If h < 1 Or h > 12 Then Exit Function
        If ap = "am" And h = 12 Then h = 0
        If ap = "pm" And h < 12 Then h = h + 12
    Else
        If h > 23 Then Exit Function
        If h >= 1 And h <= 6 Then h = h + 12
    End If
    TimeFrom = (h * 60 + m) / 1440
End Function

Private Function TimeTyped(ByVal s As String) As Double
    ' A time of day typed by itself, like 9am or 2:30 pm. -1 when it isn't one.
    Dim rest As String
    TimeTyped = SplitTime(Squeeze(s), rest)
    If rest <> "" Then TimeTyped = -1
End Function

Private Function Squeeze(ByVal s As String) As String
    ' Lowercase, trimmed, single spaces
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squeeze = s
End Function

Private Function OffWeekend(ByVal d As Date) As Date
    Select Case Weekday(d, vbMonday)
        Case 6: OffWeekend = d + 2
        Case 7: OffWeekend = d + 1
        Case Else: OffWeekend = d
    End Select
End Function

Private Function DateOf(ByVal v As Variant) As Double
    ' A cell's date and time as one number, 0 when it holds no date
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate, vbDouble, vbLong, vbInteger, vbCurrency, vbSingle
            If CDbl(v) >= 1 Then DateOf = CDbl(v)
    End Select
End Function

Private Function RemindOf(ByVal v As Variant) As Double
    ' A Remind cell's time of day as a part of a day, -1 when it holds none
    RemindOf = -1
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate, vbDouble, vbLong, vbInteger, vbCurrency, vbSingle
            If CDbl(v) >= 0 Then RemindOf = CDbl(v) - Int(CDbl(v))
        Case vbString
            RemindOf = TimeTyped(CStr(v))
    End Select
End Function

Private Function HasTime(ByVal v As Variant) As Boolean
    ' A date with a time of day on it
    Dim d As Double
    d = DateOf(v)
    HasTime = (d - Int(d)) > 0.0000001
End Function

Private Function DueText(ByVal d As Double) As String
    ' Tue 10/7, or Tue 10/7 10:00 AM when it has a time
    If d <= 0 Then
        DueText = "no date"
    Else
        DueText = Format$(CDate(d), "ddd m/d")
        If HasTime(d) Then DueText = DueText & " " & Format$(CDate(d), "h:mm AM/PM")
    End If
End Function

Private Function CleanState(ByVal st As String) As String
    st = Trim$(st)
    If Len(st) = 2 Then st = UCase$(st)
    CleanState = SafeText(st)
End Function

Private Function SafeText(ByVal s As String) As String
    ' A leading = + - or @ would make Excel read the text as a formula
    If Len(s) > 0 Then
        If InStr("=+-@", Left$(s, 1)) > 0 Then s = "'" & s
    End If
    SafeText = s
End Function

Private Function EndDot(ByVal v As Variant) As String
    ' Text to tack onto a note as its own sentence: a space first, a period last
    Dim s As String
    s = Txt(v)
    If s = "" Then Exit Function
    If InStr(".!?", Right$(s, 1)) = 0 Then s = s & "."
    EndDot = " " & s
End Function

Private Function Short(ByVal v As Variant) As String
    Short = Txt(v)
    If Len(Short) > 60 Then Short = Left$(Short, 57) & "..."
End Function

Private Function Txt(ByVal v As Variant) As String
    If Not IsError(v) Then Txt = Trim$(CStr(v))
End Function

Private Function DayNo(ByVal v As Variant) As Long
    ' The day a cell holds, as Excel's day number. 0 when it holds no date.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            DayNo = CLng(Int(CDbl(v)))
        Case vbDouble, vbLong, vbInteger, vbCurrency, vbSingle
            If v > 0 Then DayNo = CLng(Int(v))
    End Select
End Function

Private Function IsNumber(ByVal s As String) As Boolean
    ' Digits only: an item number or a count of days
    IsNumber = Len(s) > 0 And Len(s) <= 9 And s Like String(Len(s), "#")
End Function


' ---- the boxes -------------------------------------------------------------

Private Function Ask(ByVal prompt As String, ByVal preset As String) As Variant
    ' What was typed, or False if the box was cancelled
    If Not mFeed Is Nothing Then
        If mFeed.Count = 0 Then
            Ask = False
        Else
            Ask = mFeed(1)
            mFeed.Remove 1
            If Ask = "<cancel>" Then
                Ask = False
            ElseIf Ask = "<ok>" Then
                Ask = preset                     ' OK pressed on what was already filled in
            End If
        End If
        Exit Function
    End If
    Ask = Application.InputBox(prompt, "Tracker", preset, Type:=2)
End Function

Private Sub Say(ByVal msg As String)
    If Not mFeed Is Nothing Then
        mSaid = mSaid & msg & vbLf
    Else
        MsgBox msg, vbInformation, "Tracker"
    End If
End Sub

Private Function Confirm(ByVal question As String) As Boolean
    Dim a As Variant
    If Not mFeed Is Nothing Then
        If mFeed.Count > 0 Then
            a = mFeed(1)
            mFeed.Remove 1
        End If
        Confirm = (LCase$(CStr(a)) = "yes")
        Exit Function
    End If
    Confirm = (MsgBox(question, vbYesNo + vbQuestion, "Tracker") = vbYes)
End Function

Public Function TrackerTest(ByVal macroName As String, ByVal answers As String) As String
    ' For test runs: feeds answers to the boxes (split on |) and returns what the message
    ' boxes would have said. Not needed day to day.
    Dim a As Variant
    mTesting = True
    Set mFeed = New Collection
    mSaid = ""
    If answers <> "" Then
        For Each a In Split(answers, "|")
            If a = "<blank>" Then a = ""
            mFeed.Add CStr(a)
        Next
    End If
    Select Case macroName
        Case "TrackerSetup": TrackerSetup
        Case "NewItem": NewItem
        Case "UpdateItem": UpdateItem
        Case "CloseItem": CloseItem
        Case "ShowHistory": ShowHistory
        Case "BackToItems": BackToItems
        Case "RefreshToday": RefreshToday
        Case "ShowMonth": ShowMonth
        Case "PreviousMonth": PreviousMonth
        Case "NextMonth": NextMonth
        Case "ImportOld": ImportOld
        Case "Auto_Open": Auto_Open
        Case "Auto_Close": Auto_Close
        Case "Tick": CheckReminders
        Case "Clock"
            mClock = CDbl(answers)
            Set mFeed = New Collection
        Case "Ticking"
            If mNext > 0 Then mSaid = "ticking" & vbLf Else mSaid = "not ticking" & vbLf
            Set mFeed = New Collection
        Case "RealTimer"
            mRealTimer = True
            mTicks = 0
            StartTicking
            Set mFeed = New Collection
        Case "Ticks"
            mSaid = mTicks & vbLf
            Set mFeed = New Collection
        Case Else: mSaid = "no macro called " & macroName & vbLf
    End Select
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    TrackerTest = mSaid
    Set mFeed = Nothing
End Function
