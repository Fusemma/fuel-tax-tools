' Tracker: everything that lands on you at work, in one list, with Outlook doing the popups.
' To install, paste this whole module into a new module in a blank workbook, run TrackerSetup
' (Alt F8) and save it as a macro workbook (.xlsm). To update, paste a newer copy over this one
' and run TrackerSetup again; every row stays.
' How to use it is on the Today sheet, on the right.

Option Explicit

Private Const TODAY_SHEET As String = "Today"
Private Const LIST_SHEET As String = "List"
Private Const LOG_SHEET As String = "Log"
Private Const OFF_SHEET As String = "Days off"
Private Const PASTE_SHEET As String = "Paste here"
Private Const PULL_SHEET As String = "From Outlook"
Private Const PULL_HEAD As Long = 3        ' From Outlook: the column headers; its rows start under them
Private Const LIST_TABLE As String = "List"
Private Const LOG_TABLE As String = "Log"
Private Const OFF_TABLE As String = "DaysOff"

' Columns are found by name, so one you move is fine. The hidden ones are for the macro.
Private Const LIST_COLS As String = "#|What|Due|Time|Repeats|Waiting on|Came from|Off the list|Date off|" & _
    "Rule|Raw date|Email|Outlook task|Calendar"
Private Const LIST_HIDDEN As String = "Rule|Raw date|Email|Outlook task|Calendar"
Private Const LIST_FORMATS As String = "0|@|ddd m/d/yy|h:mm AM/PM|@|@|@|@|m/d/yy|@|m/d/yy|@|@|@"
Private Const LIST_WIDTHS As String = "6|50|13|10|24|20|13|12|10|20|10|20|20|20"
Private Const LOG_COLS As String = "When|#|What|What happened|Kind"
Private Const LOG_FORMATS As String = "ddd m/d/yy h:mm AM/PM|0|@|@|@"
Private Const LOG_WIDTHS As String = "22|6|44|80|10"
Private Const OFF_COLS As String = "Day off|What"
Private Const OFF_FORMATS As String = "ddd m/d/yy|@"
Private Const OFF_WIDTHS As String = "14|32"
Private Const TODAY_COLS As String = "#|When|What|Repeats|Latest note"
Private Const TODAY_FORMATS As String = "0|@|@|@|@"
Private Const TODAY_WIDTHS As String = "7|26|62|24|50"
Private Const TODAY_PT As Long = 12         ' Today's words, a size up from Excel's 11
Private Const TODAY_ROW_PT As Double = 21   ' and its rows, with room around them
Private Const TODAY_STRIPE As Long = &HF2F2F2   ' light gray, on every other thing under a heading
Private Const HOWTO_WIDTH As Double = 84    ' Today: the how-to's column, right of the list

Private Const TABLE_TOP As Long = 2        ' List, Log and Days off: a note in row 1, the table under it
Private Const COUNT_ROW As Long = 1        ' Today: the count line
Private Const SAID_ROW As Long = 4         ' Today: what the last button said
Private Const HEAD_ROW As Long = 5         ' Today: the column headers
Private Const FIRST_ROW As Long = 6        ' Today: the sections start here
Private Const CF_ROWS As Long = 20000      ' List rows that get crossed out when they're done

Private Const MORNING_AT As Long = 510     ' 8:30 AM, when the morning list comes
Private Const TICK_SECS As Long = 60       ' how often the tracker looks at Outlook while the file is open
Private Const TASK_PREFIX As String = "Tracker: "
Private Const TASK_MARK As String = "Tracker"                    ' in a popup's Mileage box, so the tracker knows its own
Private Const MORNING_MARK As String = "Tracker morning list"
Private Const TRACKER_FOLDER As String = "Tracker"               ' the popups' folder, under Tasks
Private Const FOLDER_SETTING As String = "TrackerTasksFolder"    ' a hidden name in the workbook: main means the main Tasks folder
Private Const GONE_MARK As String = "dismissed"                  ' a popup that was dismissed, so it isn't made again
Private Const OLD_FLAGS As String = "urn:tracker:old-flags"     ' the workbook's note of the emails flagged before its first look
' An email's internet message ID, which stays the same when it's filed, and its flag (2 flagged, 1 done)
Private Const MSG_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const FLAG_STATUS As String = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
Private Const NO_UNDO_FLAG As String = "Nothing to undo here: an email's flag turned into a check in Outlook, and " & _
    "Undo can't take that back. Drop it takes a row off."
Private Const NOW_CAP As Long = 5          ' things shown under Now before the rest fold away
Private Const SOON_CAP As Long = 20        ' and under Coming up
Private Const SOON_DAYS As Long = 10       ' how far ahead Coming up looks
Private Const WAIT_DAYS As Long = 7        ' a waiting thing with no date gets checked on in a week
Private Const NO_TIME As Long = -1

' Words the line reader knows
Private Const DAYS_ALT As String = "sun(?:day)?|mon(?:day)?|tue(?:sday|s)?|wed(?:nesday|s)?|thu(?:rsday|rs|r)?|fri(?:day)?|sat(?:urday)?"
Private Const DAY1 As String = "(" & DAYS_ALT & ")"
Private Const DAYNC As String = "(?:" & DAYS_ALT & ")"
' as a one-time date, "sat" and "sun" are too often ordinary words, so those two need spelling out
Private Const DAY_DATE_ALT As String = "sunday|mon(?:day)?|tue(?:sday|s)?|wed(?:nesday|s)?|thu(?:rsday|rs|r)?|fri(?:day)?|saturday"
Private Const DAY_DATE As String = "(" & DAY_DATE_ALT & ")"
' a day name alone makes a repeat as a plural, "mondays", "mons" or "fris"; "tues", "weds" and "thurs" are
' just short, and "sats" and "suns" are just words
Private Const DAYS_FULL As String = "(?:(?:sun|mon|tues|wednes|thurs|fri|satur)days|mons|fris)"
' words that end the name in "waiting on the county every monday"
Private Const REPEAT_WORDS As String = "(?:every|each|daily|weekdays|weekly|bi-?weekly|monthly|semi-?monthly|" & _
    "quarterly|yearly|annually)"
Private Const MON1 As String = "(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sept?(?:ember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?"
Private Const WORD_END As String = "(?=[\s,.;!?]|$)"
Private Const NTH_NC As String = "(?:first|1st|second|2nd|third|3rd|fourth|4th|last)"
Private Const SPEC As String = "(?:the\s)?(?:" & NTH_NC & "\s" & DAYNC & "|last\s(?:business|work|working)\s?day|last\sday|month[\s-]?end|end\sof\s(?:the\s)?month|\d{1,2}(?:st|nd|rd|th))"
Private Const SPEC_LIST As String = "(" & SPEC & "(?:\s?(?:,|and|&)\s(?:on\s)?" & SPEC & ")*)"
Private Const NUM_DATE As String = "(\d{1,2})/(\d{1,2})(?:/(\d{4}|\d{2}))?"
Private Const MON_DAY As String = MON1 & "\s(\d{1,2})(?:st|nd|rd|th)?"
Private Const LEAD As String = "(?:on|by|due|starting|before)"

Private mList As ListObject      ' the three tables, found fresh by Ready each time
Private mLogT As ListObject
Private mOffT As ListObject
Private mL As Variant            ' List's rows, one column of this per row: mL(column, row)
Private mN As Long               ' how many rows mL holds
Private mNCols As Long
Private mNextId As Long
Private cId As Long, cWhat As Long, cDue As Long, cTime As Long, cRep As Long, cWait As Long
Private cFrom As Long, cOff As Long, cOffOn As Long, cRule As Long, cRaw As Long, cTask As Long
Private cMail As Long, cCal As Long
Private mPend As Collection      ' log lines waiting to be written, oldest first
Private mOff As Object           ' days off: day number to name
Private mRx As Object            ' the line reader's patterns, made once
Private mToday As Long           ' the day and minute the buttons go by
Private mMin As Long
Private mAdded As Boolean        ' the last line read became a row
Private mReadDate As Boolean     ' the last line read had a date in it, or a repeat

' Undo: the last change to the rows, kept while the file is open. Each entry is Array(row number,
' the row as it was), or Array(row number, Empty) for a row the change added.
Private mUndo As Collection
Private mUndoLog As Collection   ' the Log lines that change wrote
Private mUndoName As String      ' what it was: Done on "File the Ohio return"
Private mUndoVerb As String      ' and for the Log: Done
Private mUndoWhy As String       ' why there's nothing to undo, after a change in Outlook
Private mKeep As Collection      ' the change being made now, the same way
Private mKeepName As String
Private mKeepVerb As String
Private mUndoing As Boolean

' Outlook. A row with a set time has one Outlook task with a reminder, its popup; the Outlook task
' column holds "the task's ID|the day and minute it pops", or "dismissed|..." once it's been dismissed.
Private mGone As Collection      ' popups of rows Undo took away, to delete
Private mMornKey As String       ' the next morning list, as it was last put in Outlook
Private mOlNote As String        ' what went wrong in Outlook, said with the next message
Private mDrawn As Long           ' the day Today was last drawn for
Private mTickAt As Double        ' when the timer comes next, 0 when it's off
Private mTicks As Long           ' how many times the timer has come
Private mTickSecs As Long        ' test runs only: the timer's seconds
Private mFeed As Collection      ' test runs only: answers for the boxes, in order
Private mSaid As String          ' test runs only: what the buttons said
Private mTesting As Boolean      ' test runs only
Private mClock As Double         ' test runs only: a pretend clock


' ---- the buttons ---------------------------------------------------------

Public Sub TrackerSetup()
    Dim wsT As Worksheet, wsL As Worksheet, wsG As Worksheet, wsD As Worksheet
    On Error GoTo Oops
    Application.ScreenUpdating = False
    Set wsT = EnsureSheet(TODAY_SHEET, Nothing)
    Set wsL = EnsureSheet(LIST_SHEET, wsT)
    Set wsG = EnsureSheet(LOG_SHEET, wsL)
    Set wsD = EnsureSheet(OFF_SHEET, wsG)
    DropBlankSheets
    Set mList = EnsureTable(wsL, LIST_TABLE, LIST_COLS, LIST_FORMATS, LIST_WIDTHS, _
        "Every row the tracker holds. Done rows stay here, crossed out. The buttons are on Today.")
    HideColumns mList, LIST_HIDDEN
    CrossOutDone wsL
    Set mLogT = EnsureTable(wsG, LOG_TABLE, LOG_COLS, LOG_FORMATS, LOG_WIDTHS, _
        "A dated line for everything that happens to a row, newest on top. # is the row's number on List.")
    HideColumns mLogT, "Kind"
    Set mOffT = EnsureTable(wsD, OFF_TABLE, OFF_COLS, OFF_FORMATS, OFF_WIDTHS, _
        "Your days off. Anything that lands on one shows the workday before, and tomorrow skips them, like weekends. " & _
        "Change them to match what the company gives.")
    If mOffT.ListRows.Count = 0 Then FillDaysOff Year(Int(NowT()))
    DressToday wsT
    Application.OnKey "^+a", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AddThing"
    If Not Ready() Then GoTo Done
    DrawToday
    wsT.Activate
    wsT.Range("A" & FIRST_ROW).Select
    Application.ScreenUpdating = True
    RunTick
    StartTimer
    Say "Tracker is set up. Press Add, or Ctrl+Shift+A, and type one line, like call Huntsville back 11/17 9am. " & _
        "Save this file as a macro workbook (.xlsm)."
Done:
    Application.ScreenUpdating = True
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Setup stopped: " & Err.Description
End Sub

Public Sub AddThing()
    ' One line in a box, and it's a row
    Dim a As Variant, msg As String
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    a = Ask("One line: what, and when if you like." & vbLf & "Like: call Huntsville back 11/17 9am", "Add")
    If VarType(a) = vbBoolean Then Exit Sub                ' the box was cancelled
    msg = AddLine(CStr(a), "typed")
    Finish
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Add stopped: " & Err.Description
End Sub

Public Sub DoneThing()
    ' The clicked row is done: off Today, crossed out on List. A repeat moves to its next date
    ' on the same row instead, and ends after its last one.
    Dim id As Long, k As Long, msg As String
    On Error GoTo Oops
    If Not OpenRow("Done", id, k) Then Exit Sub
    Keep k, "Done on """ & CStr(mL(cWhat, k)) & """", "Done"
    msg = DoneRow(k)
    Finish
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Done stopped: " & Err.Description
End Sub

Private Function DoneRow(ByVal k As Long) As String
    ' Row k is done, or a repeat moves to its next date. Returns what to say about it.
    Dim id As Long, what As String, rule As String, raw As Long, due As Long, tm As Long
    id = CLng(mL(cId, k))
    what = CStr(mL(cWhat, k))
    rule = Trim$(CellText(mL(cRule, k)))
    If rule <> "" Then
        raw = CLng(Int(NumOf(mL(cRaw, k))))
        If raw = 0 Then raw = CLng(Int(NumOf(mL(cDue, k))))
        raw = RuleNext(rule, raw)
        Do While raw <> 0
            If WorkdayOnOrBefore(raw) > mToday Then Exit Do
            raw = RuleNext(rule, raw)
        Loop
        If raw = 0 Or (RuleNum(rule, "until") > 0 And raw > RuleNum(rule, "until")) Then
            mL(cOff, k) = "Ended"
            mL(cOffOn, k) = CDbl(mToday)
            LogLine id, what, "Done. That was the last one.", "done"
            DoneRow = "Done: """ & what & """. That was the last one, so the repeat ended."
        Else
            due = WorkdayOnOrBefore(raw)
            tm = TimeOf(mL(cTime, k))
            mL(cRaw, k) = CDbl(raw)
            mL(cDue, k) = CDbl(due)
            LogLine id, what, "Done. Next one " & FmtDay(due) & ".", "done"
            DoneRow = "Done: """ & what & """. Same row, next one " & FmtDay(due) & _
                IIf(tm <> NO_TIME, " at " & FmtTime(tm), "") & "." & IIf(raw <> due, " " & MovedWhy(raw, due), "")
        End If
        Exit Function
    End If
    mL(cOff, k) = "Done"
    mL(cOffOn, k) = CDbl(mToday)
    LogLine id, what, "Done.", "done"
    DoneRow = "Done: """ & what & """. It's off Today. The List keeps it, crossed out."
End Function

Public Sub MoveTomorrow()
    ' The clicked row goes to the next workday. Something already due later stays put.
    Dim id As Long, k As Long, what As String, due As Long, nextDay As Long, tm As Long
    On Error GoTo Oops
    If Not OpenRow("Move to tomorrow", id, k) Then Exit Sub
    what = CStr(mL(cWhat, k))
    due = CLng(Int(NumOf(mL(cDue, k))))
    nextDay = NextWorkday(mToday)
    If due >= nextDay Then
        Say """" & what & """ is already due " & FmtDay(due) & ". Use Change date to move it somewhere else."
        Exit Sub
    End If
    tm = TimeOf(mL(cTime, k))
    Keep k, "Move to tomorrow on """ & what & """", "Move to tomorrow"
    mL(cDue, k) = CDbl(nextDay)
    LogLine id, what, "Moved to " & FmtDay(nextDay) & ".", "moved"
    Finish
    If tm <> NO_TIME Then
        Say "Moved """ & what & """ to " & FmtDay(nextDay) & " at " & FmtTime(tm) & ". Its Outlook popup moves with it."
    Else
        Say "Moved """ & what & """ to " & FmtDay(nextDay) & ". It'll be in that morning's " & FmtTime(MORNING_AT) & " list."
    End If
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Move to tomorrow stopped: " & Err.Description
End Sub

Public Sub ChangeDate()
    ' A new date for the clicked row, in the same plain words as Add. On a repeat only this
    ' one moves; the next one still steps from its usual date.
    Dim id As Long, k As Long, what As String, a As Variant, bad As String, oldDue As Long, due As Long, tm As Long
    Dim nWhat As String, nDue As Long, nTm As Long, who As String, notes As String, rule As String, raw As Long
    Dim msg As String
    On Error GoTo Oops
    If Not OpenRow("Change date", id, k) Then Exit Sub
    what = CStr(mL(cWhat, k))
    a = Ask("The new date for """ & what & """." & vbLf & "Like: fri 2pm, 11/20, next week", "Change date")
    If VarType(a) = vbBoolean Then Exit Sub
    bad = ReadLine(CStr(a), nWhat, nDue, nTm, who, notes, rule, raw)
    If bad <> "" Or rule <> "" Then
        Say "Change date moves this one only. To change how it repeats, Drop it and add it again."
        Exit Sub
    End If
    If Not mReadDate And nTm = NO_TIME Then
        Say "Couldn't read a date in """ & TrimAll(CStr(a)) & """. Try ""fri"", ""11/20"" or ""next week 2pm""."
        Exit Sub
    End If
    oldDue = CLng(Int(NumOf(mL(cDue, k))))
    tm = TimeOf(mL(cTime, k))
    If mReadDate Then
        due = nDue
    ElseIf nTm <= mMin And oldDue <= mToday Then
        due = NextWorkday(mToday)                    ' a time already gone today: the next workday
    Else
        due = oldDue                                 ' just a time: the day stays
        notes = ""
    End If
    If nTm <> NO_TIME Then tm = nTm
    Keep k, "Change date on """ & what & """", "Change date"
    mL(cDue, k) = CDbl(due)
    If tm <> NO_TIME Then mL(cTime, k) = tm / 1440
    LogLine id, what, "Moved to " & FmtDay(due) & IIf(tm <> NO_TIME, " " & FmtTime(tm), "") & ".", "moved"
    msg = "Moved """ & what & """ to " & FmtDay(due) & IIf(tm <> NO_TIME, ", popup at " & FmtTime(tm), "") & "."
    If Trim$(CellText(mL(cRule, k))) <> "" Then msg = msg & " Only this one moved. The next repeat keeps its usual date."
    If notes <> "" Then msg = msg & " " & notes
    Finish
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Change date stopped: " & Err.Description
End Sub

Public Sub NoteThing()
    ' A dated line on the Log for the clicked row. Today shows the newest one.
    Dim id As Long, k As Long, what As String, a As Variant, t As String
    On Error GoTo Oops
    If Not OpenRow("Note", id, k) Then Exit Sub
    what = CStr(mL(cWhat, k))
    a = Ask("A note on """ & what & """." & vbLf & "Like: left a voicemail, they'll call back Friday", "Note")
    If VarType(a) = vbBoolean Then Exit Sub
    t = TrimAll(CStr(a))
    If t = "" Then
        Say "The note was empty, so nothing was saved."
        Exit Sub
    End If
    Keep 0, "the note on """ & what & """", "the note"
    LogLine id, what, t, "note"
    Finish
    Say "Note saved on """ & what & """. It's a dated line on the Log, and the newest one shows on Today."
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Note stopped: " & Err.Description
End Sub

Public Sub DropIt()
    ' The clicked row comes off the list, or a repeat stops. List keeps it, crossed out.
    Dim id As Long, k As Long, what As String, isRep As Boolean
    On Error GoTo Oops
    If Not OpenRow("Drop it", id, k) Then Exit Sub
    what = CStr(mL(cWhat, k))
    isRep = Trim$(CellText(mL(cRule, k))) <> ""
    Keep k, "Drop it on """ & what & """", "Drop it"
    mL(cOffOn, k) = CDbl(mToday)
    If isRep Then
        mL(cOff, k) = "Stopped"
        LogLine id, what, "Stopped repeating.", "stopped"
    Else
        mL(cOff, k) = "Dropped"
        LogLine id, what, "Dropped.", "dropped"
    End If
    Finish
    If isRep Then
        Say """" & what & """ stopped repeating. The List keeps it, crossed out."
    Else
        Say "Dropped """ & what & """. The List keeps it, crossed out."
    End If
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Drop it stopped: " & Err.Description
End Sub

Public Sub UndoLast()
    ' Takes back the last change to the rows: they go back the way they were, and the Log lines
    ' that change wrote go. One step only, and nothing in Outlook.
    Dim i As Long, j As Long, k As Long, e As Variant, id As Long, what As String, ln As Variant
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If mUndo Is Nothing Then
        If mUndoWhy <> "" Then
            Say mUndoWhy
        Else
            Say "Nothing to undo. Undo goes back one change to the tracker's rows. Changes in Outlook, " & _
                "like taking things off your calendar, stay as they are."
        End If
        Exit Sub
    End If
    For i = mUndo.Count To 1 Step -1
        e = mUndo(i)
        k = RowOf(e(0))
        If k > 0 Then                                   ' a row deleted by hand since stays deleted
            what = CStr(mL(cWhat, k))
            If IsArray(e(1)) Then
                For j = 1 To mNCols                     ' its popup stays as Outlook has it; the row decides it next
                    If j <= UBound(e(1)) And j <> cTask Then mL(j, k) = e(1)(j)
                Next
                LogLine e(0), what, "Undid " & mUndoVerb & ".", "undo"
            Else
                If CellText(mL(cTask, k)) <> "" Then mGone.Add CellText(mL(cTask, k))
                DropRowAt k
                LogLine 0, what, "Undid adding it.", "undo"
            End If
        End If
    Next
    If mUndo.Count = 0 Then                             ' a note: no row changed
        For Each ln In mUndoLog
            LogLine ln(1), ln(2), "Undid " & mUndoVerb & ".", "undo"
        Next
    End If
    RemoveLogLines mUndoLog
    mUndoing = True
    Finish
    mUndoing = False
    Say "Undid " & mUndoName & ". Anything already done in Outlook stays done."
    Set mUndo = Nothing
    Set mUndoLog = Nothing
    Exit Sub
Oops:
    mUndoing = False
    Application.ScreenUpdating = True
    Say "Undo stopped: " & Err.Description
End Sub

Public Sub PasteList()
    ' A sheet to paste a list on, one thing per line
    Dim ws As Worksheet
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If SheetExists(PASTE_SHEET) Then
        ThisWorkbook.Worksheets(PASTE_SHEET).Activate
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = PASTE_SHEET
    ws.Columns(1).ColumnWidth = 100
    ws.Columns(1).NumberFormat = "@"
    ws.Rows(1).RowHeight = 30
    ActiveWindow.DisplayGridlines = False
    AddButton ws, "trkTake", "Take these", "TakeThese", 0, 0, 80
    AddButton ws, "trkCancel", "Cancel", "CancelPaste", 0, 84, 60
    ws.Range("A2").Value = "Paste one thing per line, starting in A3, then press Take these. Plain words work: " & _
        """fri 2pm"", ""11/17"", ""next week"", ""waiting on accounting""."
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Range("A3").Select
    Application.ScreenUpdating = True
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Paste a list stopped: " & Err.Description
End Sub

Public Sub TakeThese()
    ' Every line on Paste here becomes a row; the ones it can't read get named
    Dim lines As Collection, missed As String, took As Long, l As Variant, msg As String
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not SheetExists(PASTE_SHEET) Then
        Say "Press Paste a list first."
        Exit Sub
    End If
    Set lines = PastedLines(ThisWorkbook.Worksheets(PASTE_SHEET))
    If lines.Count = 0 Then
        Say "The paste box is empty."
        Exit Sub
    End If
    For Each l In lines
        AddLine CStr(l), "pasted"
        If mAdded Then
            took = took + 1
        Else
            missed = missed & IIf(missed = "", "", ", ") & """" & l & """"
        End If
    Next
    DropPasteSheet
    Keep 0, "Take these (" & took & IIf(took = 1, " row)", " rows)"), "adding it"
    Finish
    msg = "Took " & took & " of " & lines.Count & " lines. Each one is a row now, and the List sheet shows how it read them."
    If missed <> "" Then msg = msg & " Couldn't read the repeat in: " & missed & ". Fix those and paste them again."
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Take these stopped: " & Err.Description
End Sub

Public Sub CancelPaste()
    On Error GoTo Oops
    DropPasteSheet
    If SheetExists(TODAY_SHEET) Then ThisWorkbook.Worksheets(TODAY_SHEET).Activate
    Say "Nothing was taken from the paste."
    Exit Sub
Oops:
    Say "Cancel stopped: " & Err.Description
End Sub

Public Sub Auto_Open()
    ' Excel runs this when the workbook opens: a fresh Today, the shortcut key, every popup
    ' checked and the next morning list made, then the one-minute timer
    If FindTable(LIST_TABLE) Is Nothing Then Exit Sub          ' TrackerSetup hasn't run yet
    If Not HasColumns(FindTable(LIST_TABLE), LIST_COLS) Then Exit Sub
    Application.OnKey "^+a", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AddThing"
    If Not Ready() Then Exit Sub
    DrawToday
    ThisWorkbook.Worksheets(TODAY_SHEET).Activate
    RunTick
    StartTimer
End Sub

Public Sub Auto_Close()
    ' Excel runs this when the workbook closes. A shortcut or a timer left set would open the file again.
    Application.OnKey "^+a"
    StopTimer
End Sub

Public Sub TrackerTick()
    ' The one-minute timer: looks at Outlook, then sets itself for the next minute
    mTickAt = 0
    mTicks = mTicks + 1
    RunTick
    StartTimer
End Sub

Public Sub TrackerUseMainTasks()
    ' For a work PC where popups from the Tracker folder don't come up: they go in the main Tasks folder
    SwitchTasks True
End Sub

Public Sub TrackerUseTrackerFolder()
    ' Back to the Tracker folder under Tasks, where they start
    SwitchTasks False
End Sub

Public Sub OpenEmail()
    ' Opens the clicked row's email in Outlook, wherever it's been filed since
    Dim k As Long, link As String, what As String, ns As Object, m As Object
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    k = RowOf(PickedId())
    If k > 0 Then link = CellText(mL(cMail, k))
    If link = "" Then
        Say "Click a row that came from an email first."
        Exit Sub
    End If
    what = CellText(mL(cWhat, k))
    If Not OutlookOpen("Open email", ns) Then Exit Sub
    Set m = FindEmail(ns, link)
    If m Is Nothing Then
        Say "Couldn't find the email for """ & what & """ in Outlook. It may have been deleted."
        Exit Sub
    End If
    If Not mTesting Then m.Display
    Say "Opened the email for """ & what & """. It's in your " & m.Parent.Name & " folder."
    Exit Sub
Oops:
    Say "Open email stopped: " & Err.Description
End Sub

Public Sub PullFromOutlook()
    ' Lists the calendar's repeats and what's still to come on it, and the emails flagged before the
    ' tracker first looked, on a From Outlook sheet with a Yes or No for each
    Dim ns As Object, found As Collection, ws As Worksheet, v As Variant, i As Long, j As Long
    Dim nCal As Long, nMail As Long, nYes As Long, msg As String, last As Long
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not OutlookOpen("Pull from Outlook", ns) Then Exit Sub
    Set found = New Collection
    ListCalendar ns, found, nCal
    ListOldFlags ns, found, nMail
    DropPullSheet
    If found.Count = 0 Then
        Say "Nothing new in Outlook to pull."
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = PULL_SHEET
    ActiveWindow.DisplayGridlines = False
    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 50
    ws.Columns(3).ColumnWidth = 44
    ws.Columns(4).ColumnWidth = 60
    ws.Columns("B:E").NumberFormat = "@"
    ws.Rows(1).RowHeight = 30
    AddButton ws, "trkTakeYes", "Take the Yes ones", "TakeYes", 0, 0, 110
    AddButton ws, "trkCancelPull", "Cancel", "CancelPull", 0, 114, 60
    ws.Range("A" & PULL_HEAD & ":D" & PULL_HEAD).Value = Array("Take?", "What", "In Outlook", "Why")
    ws.Range("A" & PULL_HEAD & ":D" & PULL_HEAD).Font.Bold = True
    ReDim v(1 To found.Count, 1 To 5)
    For i = 1 To found.Count
        For j = 0 To 4
            v(i, j + 1) = found(i)(j)
        Next
        If found(i)(0) = "Yes" Then nYes = nYes + 1
    Next
    last = PULL_HEAD + found.Count
    ws.Range(ws.Cells(PULL_HEAD + 1, 1), ws.Cells(last, 5)).Value = v
    With ws.Range(ws.Cells(PULL_HEAD + 1, 1), ws.Cells(last, 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
    End With
    ws.Columns(5).Hidden = True                          ' where each one is in Outlook, for Take the Yes ones
    msg = "Found " & nCal & IIf(nCal = 1, " thing", " things") & " on your calendar and " & nMail & _
        IIf(nMail = 1, " email", " emails") & " you'd already flagged. " & nYes & " say Yes: the ones that repeat " & _
        "or take all day. Meetings, set-time appointments and old flags say No. Change any Yes or No, then press " & _
        "Take the Yes ones."
    ws.Range("A2").Value = msg
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Range("A" & PULL_HEAD + 1).Select
    Application.ScreenUpdating = True
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Pull from Outlook stopped: " & Err.Description
End Sub

Public Sub TakeYes()
    ' Every row on From Outlook that says Yes comes in: a calendar series keeps its schedule, a
    ' one-time thing its day, and an email's flag turns into a check
    Dim ws As Worksheet, ns As Object, r As Long, last As Long, link As String
    Dim nCal As Long, nMail As Long, ended As Long, cant As Long, msg As String, notes As String
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not SheetExists(PULL_SHEET) Then
        Say "Press Pull from Outlook first."
        Exit Sub
    End If
    If Not OutlookOpen("Take the Yes ones", ns) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(PULL_SHEET)
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = PULL_HEAD + 1 To last
        If LCase$(Trim$(CellText(ws.Cells(r, 1).Value))) = "yes" Then
            link = CellText(ws.Cells(r, 5).Value)
            If Left$(link, 4) = "cal|" Then
                Select Case TakeCalendar(ns, Mid$(link, 5), notes)
                    Case "took": nCal = nCal + 1
                    Case "ended": ended = ended + 1
                    Case "cant": cant = cant + 1
                End Select
            ElseIf Left$(link, 5) = "mail|" Then
                If TakeOldFlag(ns, Mid$(link, 6), notes) Then nMail = nMail + 1
            End If
        End If
    Next
    DropPullSheet
    ThisWorkbook.Worksheets(TODAY_SHEET).Activate
    Keep 0, "Take the Yes ones (" & (nCal + nMail) & IIf(nCal + nMail = 1, " row)", " rows)"), "adding it"
    Finish
    If nMail > 0 Then ForgetUndo NO_UNDO_FLAG
    msg = "Took " & (nCal + nMail) & " from Outlook and typed none of them."
    If notes <> "" Then msg = msg & " " & notes
    If ended > 0 Then msg = msg & " " & ended & " had already ended, so they stayed out."
    If cant > 0 Then msg = msg & " " & cant & " repeat in a way the tracker can't follow, so they stayed out."
    If nCal > 0 Then msg = msg & " The calendar ones are still on your calendar too, until you press Take them off my calendar."
    Say msg
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Take the Yes ones stopped: " & Err.Description
End Sub

Public Sub CancelPull()
    On Error GoTo Oops
    DropPullSheet
    If SheetExists(TODAY_SHEET) Then ThisWorkbook.Worksheets(TODAY_SHEET).Activate
    Say "Nothing taken from Outlook."
    Exit Sub
Oops:
    Say "Cancel stopped: " & Err.Description
End Sub

Public Sub TakeOffCalendar()
    ' Everything taken in from the calendar, one-time things too, goes to Deleted Items, so the
    ' calendar is down to meetings
    Dim ns As Object, calId As String, k As Long, c As Object, n As Long, linked As Boolean
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not OutlookOpen("Take them off my calendar", ns) Then Exit Sub
    calId = ns.GetDefaultFolder(9).EntryID
    For k = 1 To mN
        If CellText(mL(cCal, k)) <> "" Then
            Set c = ItemById(ns, CellText(mL(cCal, k)))
            If Not c Is Nothing Then
                If c.Parent.EntryID = calId Then c.Delete: n = n + 1
            End If
            Set c = Nothing
            mL(cCal, k) = Empty
            linked = True
        End If
    Next
    If linked Then SaveList
    If n = 0 Then
        Say "Nothing to take off. Pull from Outlook first."
        Exit Sub
    End If
    ForgetUndo "Nothing to undo here: taking things off your calendar happened in Outlook, and Undo can't take " & _
        "that back. Drag any you want back from Deleted Items."
    Say "Moved the " & n & IIf(n = 1, " thing", " things") & " you took from your calendar into Deleted Items. " & _
        "They stay there until Deleted Items gets emptied, so drag back any you want before then."
    Exit Sub
Oops:
    Say "Take them off my calendar stopped: " & Err.Description
End Sub


' ---- adding a line ------------------------------------------------------------

Private Function AddLine(ByVal text As String, ByVal from As String) As String
    ' Reads one line and adds it as a row. Returns what to say about it.
    Dim tm As Long, due As Long, what As String, who As String, notes As String, bad As String
    Dim k As Long, desc As String, rule As String, raw As Long
    mAdded = False
    If TrimAll(text) = "" Then
        AddLine = "Type something in the Add box first."
        Exit Function
    End If
    bad = ReadLine(text, what, due, tm, who, notes, rule, raw)
    If bad <> "" Then
        AddLine = bad
        Exit Function
    End If
    k = NewRow()
    mKeep.Add Array(CLng(mL(cId, k)), Empty)            ' for Undo: a row this change added
    mKeepName = "adding """ & what & """": mKeepVerb = "adding it"
    mL(cWhat, k) = what
    mL(cDue, k) = CDbl(due)
    If tm <> NO_TIME Then mL(cTime, k) = tm / 1440
    If who <> "" Then mL(cWait, k) = who
    mL(cFrom, k) = from
    If rule <> "" Then
        mL(cRule, k) = rule
        mL(cRaw, k) = CDbl(raw)
        mL(cRep, k) = RuleLabel(rule)
    End If
    desc = Describe(due, tm, who, rule, raw)
    LogLine CLng(mL(cId, k)), what, "Added. " & desc, "added"
    mAdded = True
    AddLine = "Added """ & what & """: " & desc & IIf(notes <> "", " " & notes, "")
End Function

Private Function Describe(ByVal due As Long, ByVal tm As Long, ByVal who As String, ByVal rule As String, _
                          ByVal raw As Long) As String
    ' What a row's date means, in words: due when, how it'll come up, and how it repeats
    Dim d As String, t As String
    If due = mToday Then d = "today" Else d = FmtDay(due)
    If who <> "" Then t = "waiting on " & who & ", check back " & d Else t = "due " & d
    If tm <> NO_TIME Then
        t = t & ", popup at " & FmtTime(tm)
    ElseIf due = mToday And mMin >= MORNING_AT Then
        t = t & ", on Today now"
    ElseIf due = mToday Then
        t = t & ", in the " & FmtTime(MORNING_AT) & " list"
    Else
        t = t & ", in the " & FmtTime(MORNING_AT) & " list that morning"
    End If
    If rule <> "" Then t = t & ", repeats " & RuleLabel(rule)
    t = t & "."
    If rule <> "" And raw <> due Then t = t & " " & MovedWhy(raw, due)
    Describe = t
End Function

Private Function MovedWhy(ByVal raw As Long, ByVal due As Long) As String
    ' 11/26 is Thanksgiving, so it shows Wed 11/25.
    MovedWhy = FmtShort(raw) & " is " & OffName(raw) & ", so it shows " & FmtDay(due) & "."
End Function

Private Function PastedLines(ByVal ws As Worksheet) As Collection
    ' Each row's cells, one thing per line, bullets and numbers taken off
    Dim c As New Collection, v As Variant, r As Long, j As Long, last As Long, lastC As Long
    Dim rowText As String, part As Variant, t As String
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If last >= 3 Then
        v = ws.Range(ws.Cells(3, 1), ws.Cells(last, IIf(lastC < 1, 1, lastC))).Value
        If Not IsArray(v) Then
            ReDim v(1 To 1, 1 To 1)
            v(1, 1) = ws.Cells(3, 1).Value
        End If
        For r = 1 To UBound(v, 1)
            rowText = ""
            For j = 1 To UBound(v, 2)
                If Not IsError(v(r, j)) Then
                    If Trim$(CStr(v(r, j))) <> "" Then rowText = rowText & IIf(rowText = "", "", " ") & CStr(v(r, j))
                End If
            Next
            For Each part In Split(Replace(rowText, vbCr, vbLf), vbLf)
                t = Rx("^\s*(?:[-*" & ChrW(8226) & "]|\d+[.)])\s+").Replace(CStr(part), "")
                t = TrimAll(t)
                If t <> "" Then c.Add t
            Next
        Next
    End If
    Set PastedLines = c
End Function

Private Sub DropPasteSheet()
    If Not SheetExists(PASTE_SHEET) Then Exit Sub
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(PASTE_SHEET).Delete
    Application.DisplayAlerts = True
End Sub


' ---- the line reader ------------------------------------------------------------

Private Function ReadLine(ByVal text As String, ByRef what As String, ByRef due As Long, ByRef tm As Long, _
                          ByRef who As String, ByRef notes As String, ByRef rule As String, _
                          ByRef rawDay As Long) As String
    ' Reads one typed line. Fills in what, due, time, who it's waiting on, and for a repeat its
    ' rule and the date the rule landed on, and returns why it can't take the line, or "" when
    ' it can. Time words first, then until, then waiting on, then the repeat and its days, then
    ' a date.
    Dim s As String, m As Object, mc As Object, n As Long, h As Long, t As Long, i As Long, after As Long
    Dim dt As Long, hasDate As Boolean, fixed As Boolean, endOn As Long, raw As String, ok As Boolean
    Dim unit As String, days As String, dows As String, semi As Boolean, every As Long, src As Long
    raw = TrimAll(text)
    what = "": who = "": notes = "": due = 0: tm = NO_TIME: rule = "": rawDay = 0: every = 1: mReadDate = False
    s = Replace(Replace(text, vbTab, " "), ChrW(160), " ")
    s = " " & Trim$(Rx("\s+", True).Replace(s, " ")) & " "

    ' "tomorrow 9" and "fri 2" mean o'clock
    s = Rx("\s(today|tomorrow|tmrw|tmw|" & DAY_DATE_ALT & ")\s(\d{1,2})" & WORD_END).Replace(s, " $1 at $2")

    ' in an hour, in 20 minutes
    If Find(s, "\sin\s(an?|\d{1,3})\s(hours?|hrs?|minutes?|mins?)" & WORD_END, m) Then
        If Rx("^an?$").Test(SM(m, 1)) Then n = 1 Else n = CLng(SM(m, 1))
        If LCase$(Left$(SM(m, 2), 1)) = "h" Then n = n * 60
        Cut s, m
        If n <> 0 Then
            t = mMin + n
            dt = mToday + t \ 1440
            tm = t Mod 1440
            hasDate = True
        End If
    End If

    If tm = NO_TIME Then
        If Find(s, "\s(?:at\s|@\s?)?(\d{1,2})(?::([0-5]\d))?\s?([ap])\.?m\.?" & WORD_END, m) Then
            h = CLng(SM(m, 1))
            If h >= 1 And h <= 12 Then
                tm = ((h Mod 12) + IIf(LCase$(SM(m, 3)) = "p", 12, 0)) * 60 + Val(SM(m, 2))
                Cut s, m
            End If
        End If
    End If
    If tm = NO_TIME Then
        If Find(s, "\s(?:at\s)?noon" & WORD_END, m) Then
            tm = 720
            Cut s, m
        End If
    End If
    If tm = NO_TIME Then
        If Find(s, "\s(?:at\s|@\s?)?(\d{1,2}):([0-5]\d)" & WORD_END, m) Then
            h = CLng(SM(m, 1))
            If h <= 23 Then
                If h >= 1 And h <= 6 Then h = h + 12
                tm = h * 60 + CLng(SM(m, 2))
                Cut s, m
            End If
        End If
    End If
    If tm = NO_TIME Then
        ' "at 3": 1 to 6 is afternoon, 7 to 11 is morning
        If Find(s, "\s(?:at|@)\s?(\d{1,2})" & WORD_END, m) Then
            h = CLng(SM(m, 1))
            If h >= 1 And h <= 12 Then
                If h <= 6 Then h = h + 12
                tm = h * 60
                Cut s, m
            End If
        End If
    End If

    If Find(s, "\s(?:until|till|thru|through|ending)\s(\d{1,2})/(\d{1,2})(?:/(\d{4}|\d{2}))?" & WORD_END, m) Then
        endOn = PickYear(CLng(SM(m, 1)), CLng(SM(m, 2)), SM(m, 3))
        If endOn > 0 Then Cut s, m
    End If
    If endOn = 0 Then
        If Find(s, "\s(?:until|till|thru|through|ending)\s" & MON1 & "\s(\d{1,2})" & WORD_END, m) Then
            endOn = PickYear(MonNum(SM(m, 1)), CLng(SM(m, 2)), "")
            If endOn > 0 Then Cut s, m
        End If
    End If

    If Find(s, "\s(?:waiting|wait)\son\s(.+?)(?=[,;]|\s(?:for|about|re|to|on|with|check|follow|by|until|from|regarding|they|he|she)\s|" & _
            "\s(?:" & REPEAT_WORDS & "|" & DAYS_FULL & ")" & WORD_END & "|\s\d|\s*$)", m) Then
        who = Trim$(SM(m, 1))
        after = m.FirstIndex + m.Length
        Set mc = Rx("\s(?:check(?:\sback)?|follow\sup)(?:\s(?:on|by))?(?=\s)").Execute(Mid$(s, after + 1))
        If mc.Count > 0 Then
            i = after + mc(0).FirstIndex
            s = Left$(s, i) & " " & Mid$(s, i + mc(0).Length + 1)
        End If
    End If

    ' the repeat
    If Find(s, "\s(?:semi-?monthly|twice\sa\smonth)" & WORD_END, m) Then
        unit = "month": semi = True: Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:every|each)\sother\s" & DAY1 & "s?" & WORD_END, m) Then
            unit = "week": every = 2: dows = CStr(DayNum(SM(m, 1))): Cut s, m
        End If
    End If
    If unit = "" Then
        If Find(s, "\s(?:every|each)\s(\d{1,2})\s(day|week|month)s?" & WORD_END, m) Then
            n = CLng(SM(m, 1))
            If n <> 0 Then
                Select Case LCase$(SM(m, 2))
                    Case "day": If n = 1 Then unit = "workday" Else unit = "day"
                    Case Else: unit = LCase$(SM(m, 2))
                End Select
                every = n
                Cut s, m
            End If
        End If
    End If
    If unit = "" Then
        If Find(s, "\s(?:bi-?weekly|every\s(?:two|other)\sweeks?)" & WORD_END, m) Then unit = "week": every = 2: Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:(?:every|each)\s(?:work\s?day|week\s?day|business\sday|day)|weekdays|daily)" & WORD_END, m) Then unit = "workday": Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:quarterly|(?:every|each)\squarter)" & WORD_END, m) Then unit = "month": every = 3: Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:monthly|(?:every|each)\smonth|once\sa\smonth|of\s(?:every|each)\smonth)" & WORD_END, m) Then unit = "month": Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:every|each)\s" & DayList("") & WORD_END, m) Then unit = "week": dows = DayNums(SM(m, 1)): Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(" & DAYS_FULL & "(?:\s?(?:,|and|&)\s" & DAYS_FULL & ")*)" & WORD_END, m) Then
            unit = "week": dows = DayNums(SM(m, 1)): Cut s, m
        End If
    End If
    If unit = "" Then
        If Find(s, "\s(?:weekly|(?:every|each)\sweek|once\sa\sweek)" & WORD_END, m) Then unit = "week": Cut s, m
    End If
    If unit = "" Then
        If Find(s, "\s(?:yearly|annually|annual|(?:every|each)\syear|once\sa\syear)" & WORD_END, m) Then unit = "year": Cut s, m
    End If

    If unit = "month" Then
        If Find(s, "\s(?:on\s|by\s)?" & SPEC_LIST & WORD_END, m) Then
            days = SpecList(SM(m, 1))
            If days <> "" Then Cut s, m
        End If
        If Find(s, "\s(?:of\s)?(?:the|every|each)\smonth" & WORD_END, m) Then Cut s, m
    End If
    If unit = "week" And dows = "" Then
        If Find(s, "\s(?:on\s)?" & DayList("s?") & WORD_END, m) Then dows = DayNums(SM(m, 1)): Cut s, m
    End If

    If Rx("\b(?:monthly|weekly|quarterly|yearly|annually|semi-?monthly|bi-?weekly|twice\sa\s(?:month|week)|" & _
          "of\s(?:every|each)\smonth|every\s(?:other|\d+)\b|every\s(?:sun|mon|tue|wed|thu|fri|sat))").Test(s) _
       Or (semi And ListCount(days) < 2) Or (unit = "" And Rx("\bevery\b").Test(s)) Then
        ' never a guess: "every" with nothing it can read after it is refused too
        ReadLine = "Couldn't read the repeat in """ & raw & """, so nothing was added. Try ""monthly on the 15th"", " & _
            """every friday"" or ""first monday of every month""."
        Exit Function
    End If

    ' the date. A written-out date counts only after on, by, due, starting or before, or at the
    ' start or end of the line, never mid-sentence. A yearly repeat needs its month and day, so
    ' there a date anywhere in the line counts.
    If Not hasDate Then
        If unit = "year" Then
            ok = TakeDate(s, "\s" & NUM_DATE & WORD_END, "md", dt)
            If Not ok Then ok = TakeDate(s, "\s" & MON_DAY & WORD_END, "mn", dt)
        End If
        If Not ok Then ok = TakeDate(s, "\s" & LEAD & "\s" & NUM_DATE & WORD_END, "md", dt)
        If Not ok Then ok = TakeDate(s, "^\s" & NUM_DATE & "(?=\s)", "md", dt)
        If Not ok Then ok = TakeDate(s, "\s" & NUM_DATE & "\s*$", "md", dt)
        If Not ok Then ok = TakeDate(s, "\s" & LEAD & "\s" & MON_DAY & WORD_END, "mn", dt)
        If Not ok Then ok = TakeDate(s, "^\s" & MON_DAY & "(?=\s)", "mn", dt)
        If Not ok Then ok = TakeDate(s, "\s" & MON_DAY & "\s*$", "mn", dt)
        If Not ok Then ok = TakeDate(s, "\s" & LEAD & "\s(?:the\s)?(\d{1,2})(?:st|nd|rd|th)" & WORD_END, "dom", dt)
        If Not ok Then ok = TakeDate(s, "\s(?:the\s)?(\d{1,2})(?:st|nd|rd|th)\s*$", "dom", dt)
        If Not ok Then ok = TakeDate(s, "\s(?:by\s)?(?:eom|(?:the\s)?end\sof\s(?:the\s)?month)" & WORD_END, "eom", dt)
        If ok Then
            fixed = True
        Else
            ok = TakeDate(s, "\s(?:today|tonight|eod|end\sof\s(?:the\s)?day)" & WORD_END, "today", dt)
            If Not ok Then ok = TakeDate(s, "\s(?:tomorrow|tmrw|tmw|tomorow)" & WORD_END, "tomorrow", dt)
            If Not ok Then ok = TakeDate(s, "\s(?:by\s)?(?:(?:the\s)?end\sof\s(?:the\s)?week|eow)" & WORD_END, "eow", dt)
            If Not ok Then ok = TakeDate(s, "\snext\sweek" & WORD_END, "nextweek", dt)
            If Not ok Then ok = TakeDate(s, "\snext\smonth" & WORD_END, "nextmonth", dt)
            If Not ok Then ok = TakeDate(s, "\s(?:in\s(an?|\d{1,3})\s(days?|weeks?|months?)|(\d{1,3})(d|w))" & WORD_END, "in", dt)
            ' "fri" is the coming Friday; "next fri" is the Friday of next week. Like a written date,
            ' a day name on a weekend or day off shows the workday before.
            If Not ok Then
                ok = TakeDate(s, "\s(?:on\s|by\s|this\s|due\s)?(next\s)?" & DAY_DATE & WORD_END, "dow", dt)
                fixed = ok
            End If
        End If
        hasDate = ok
    End If
    mReadDate = hasDate Or unit <> ""

    what = Trim$(Rx("\s+", True).Replace(s, " "))
    For i = 1 To 3
        what = Rx("^(?:on|by|at|due|and|starting)\s+").Replace(what, "")
        what = Rx("(?:^|\s)(?:on|by|at|due|and|every|each|starting|from)$").Replace(what, "")
        what = Trim$(Rx("[\s,;:-]+$").Replace(what, ""))
    Next
    If what = "" Then what = raw
    what = UCase$(Left$(what, 1)) & Mid$(what, 2)

    If unit <> "" Then
        ' a repeat: its rule, then the first date it lands on from today or the date given
        If hasDate Then src = dt Else src = mToday
        rule = unit
        Select Case unit
            Case "day"
                rule = rule & " every=" & every & " anchor=" & src
            Case "week"
                If dows = "" Then dows = CStr(Dow(src))
                rule = rule & " every=" & every & " dows=" & dows
                If every > 1 Then
                    If hasDate Then
                        rule = rule & " anchor=" & dt
                    Else
                        rule = rule & " anchor=" & NextDow(mToday, CLng(Split(dows, ",")(0)), False)
                    End If
                End If
            Case "month"
                If days = "" Then
                    days = CStr(Day(src))
                    If Not hasDate Then AddNote notes, "No day was given, so it uses today's date, the " & Ord(Day(mToday)) & "."
                End If
                rule = rule & " every=" & every & " days=" & days
                If every > 1 Then
                    If hasDate Then
                        rule = rule & " phase=" & Month(dt)
                    ElseIf every = 3 Then
                        rule = rule & " phase=1"                ' quarterly: Jan, Apr, Jul and Oct
                    Else
                        rule = rule & " phase=" & Month(mToday)
                    End If
                End If
            Case "year"
                rule = rule & " month=" & Month(src) & " dom=" & Day(src)
        End Select
        If endOn > 0 Then rule = rule & " until=" & endOn
        If hasDate And dt > mToday Then src = dt Else src = mToday
        rawDay = FirstRaw(rule, src, tm)
        If rawDay = 0 Then
            ReadLine = "That repeat ends before its first date, so nothing was added."
            rule = ""
            Exit Function
        End If
        due = WorkdayOnOrBefore(rawDay)
        Exit Function
    End If

    If hasDate Then
        due = dt
        If fixed And IsOff(dt) Then
            due = WorkdayOnOrBefore(dt)
            AddNote notes, FmtShort(dt) & " is " & OffName(dt) & ", so it shows " & FmtDay(due) & "."
        End If
    ElseIf who <> "" Then
        due = WorkdayOnOrAfter(mToday + WAIT_DAYS)
    Else
        due = mToday
        If tm = NO_TIME Then AddNote notes, "There's no date in that line, so it's due today."
    End If
    If tm <> NO_TIME And due = mToday And tm <= mMin Then
        If Not hasDate Then
            due = NextWorkday(mToday)
            AddNote notes, FmtTime(tm) & " has already passed today, so it's " & FmtDay(due) & "."
        Else
            AddNote notes, FmtTime(tm) & " has already passed, so there's no popup today."
        End If
    End If
End Function

Private Function TakeDate(ByRef s As String, ByVal pat As String, ByVal kind As String, ByRef dt As Long) As Boolean
    ' Looks for one kind of date. When it's there and makes a real day, takes it out of the line.
    Dim m As Object, n As Long, y As Long, mo As Long, i As Long, x As Long, u As String, d As Long
    If Not Find(s, pat, m) Then Exit Function
    y = Year(mToday): mo = Month(mToday)
    Select Case kind
        Case "md": d = PickYear(CLng(SM(m, 1)), CLng(SM(m, 2)), SM(m, 3))
        Case "mn": d = PickYear(MonNum(SM(m, 1)), CLng(SM(m, 2)), "")
        Case "dom"
            n = CLng(SM(m, 1))
            If n >= 1 And n <= 31 Then
                For i = 0 To 2
                    If n <= DaysIn(y, mo) Then
                        x = DateSerial(y, mo, n)
                        If x >= mToday Then d = x: Exit For
                    End If
                    mo = mo + 1
                    If mo > 12 Then mo = 1: y = y + 1
                Next
            End If
        Case "eom": d = DateSerial(y, mo, DaysIn(y, mo))
        Case "today": d = mToday
        Case "tomorrow": d = NextWorkday(mToday)
        Case "eow": d = WorkdayOnOrBefore(NextDow(mToday, 5, False))
        Case "nextweek": d = WorkdayOnOrAfter(NextDow(mToday, 1, True))
        Case "nextmonth": d = WorkdayOnOrAfter(AddMonthsKeepDay(mToday, 1))
        Case "in"
            If SM(m, 1) <> "" And Rx("^an?$").Test(SM(m, 1)) Then
                n = 1
            ElseIf SM(m, 1) <> "" Then
                n = CLng(SM(m, 1))
            Else
                n = CLng(SM(m, 3))
            End If
            u = LCase$(Left$(SM(m, 2) & SM(m, 4), 1))
            If n <> 0 Then
                If u = "d" Then
                    d = WorkdayOnOrAfter(mToday + n)
                ElseIf u = "w" Then
                    d = WorkdayOnOrAfter(mToday + 7 * n)
                Else
                    d = WorkdayOnOrAfter(AddMonthsKeepDay(mToday, n))
                End If
            End If
        Case "dow"
            If SM(m, 1) <> "" Then
                d = NextDow(NextDow(mToday, 1, True), DayNum(SM(m, 2)), False)
            Else
                d = NextDow(mToday, DayNum(SM(m, 2)), False)
            End If
    End Select
    If d = 0 Then Exit Function
    dt = d
    Cut s, m
    TakeDate = True
End Function

Private Function PickYear(ByVal mo As Long, ByVal d As Long, ByVal y As String) As Long
    ' A month and day as a day number, 0 when there's no such day. With no year, the first one
    ' from two months back onward, so a recent date stays in the past and 2/29 waits for a leap year.
    Dim yy As Long, ty As Long
    If mo < 1 Or mo > 12 Or d < 1 Or d > 31 Then Exit Function
    If y <> "" Then
        yy = CLng(y)
        If Len(y) = 2 Then yy = 2000 + yy
        If d <= DaysIn(yy, mo) Then PickYear = DateSerial(yy, mo, d)
        Exit Function
    End If
    ty = Year(mToday)
    For yy = ty - 1 To ty + 4
        If d <= DaysIn(yy, mo) Then
            If DateSerial(yy, mo, d) >= mToday - 60 Then
                PickYear = DateSerial(yy, mo, d)
                Exit Function
            End If
        End If
    Next
End Function

Private Function DayList(ByVal plural As String) As String
    DayList = "(" & DAYNC & plural & "(?:\s?(?:,|and|&)\s" & DAYNC & plural & ")*)"
End Function

Private Function SplitList(ByVal t As String) As Variant
    ' "mon, wed and fri" as its parts
    Dim parts As Variant, out() As String, n As Long, i As Long
    parts = Split(Rx("\s?(?:,|\band\b|&)\s(?:on\s)?", True).Replace(t, Chr$(1)), Chr$(1))
    ReDim out(0 To UBound(parts) + 1)
    n = -1
    For i = 0 To UBound(parts)
        If Trim$(parts(i)) <> "" Then n = n + 1: out(n) = Trim$(parts(i))
    Next
    If n < 0 Then
        SplitList = Array()
    Else
        ReDim Preserve out(0 To n)
        SplitList = out
    End If
End Function

Private Function DayNums(ByVal t As String) As String
    ' "tuesday and friday" as "2,5"
    Dim p As Variant, out As String
    For Each p In SplitList(t)
        out = out & IIf(out = "", "", ",") & DayNum(CStr(p))
    Next
    DayNums = out
End Function

Private Function SpecList(ByVal t As String) As String
    ' "the 15th and last day" as "15;last", or "" when any part isn't a day of the month
    Dim p As Variant, one As String, out As String
    For Each p In SplitList(t)
        one = ParseSpec(CStr(p))
        If one = "" Then Exit Function
        out = out & IIf(out = "", "", ";") & one
    Next
    SpecList = out
End Function

Private Function ParseSpec(ByVal t As String) As String
    ' One day of the month: 15, last, lastwork, or 1:1 for the first Monday
    Dim m As Object
    t = Rx("^the\s").Replace(t, "")
    If Find(t, "^(" & NTH_NC & ")\s" & DAY1 & "$", m) Then
        ParseSpec = NthNum(SM(m, 1)) & ":" & DayNum(SM(m, 2))
    ElseIf Rx("^last\s(?:business|work|working)\s?day$").Test(t) Then
        ParseSpec = "lastwork"
    ElseIf Rx("^(?:last\sday|month[\s-]?end|end\sof\s(?:the\s)?month)$").Test(t) Then
        ParseSpec = "last"
    ElseIf Find(t, "^(\d{1,2})(?:st|nd|rd|th)$", m) Then
        If CLng(SM(m, 1)) >= 1 And CLng(SM(m, 1)) <= 31 Then ParseSpec = CStr(CLng(SM(m, 1)))
    End If
End Function

Private Function ListCount(ByVal t As String) As Long
    If t = "" Then ListCount = 0 Else ListCount = UBound(Split(t, ";")) + 1
End Function

Private Function DayNum(ByVal w As String) As Long
    Select Case LCase$(Left$(Trim$(w), 2))
        Case "su": DayNum = 0
        Case "mo": DayNum = 1
        Case "tu": DayNum = 2
        Case "we": DayNum = 3
        Case "th": DayNum = 4
        Case "fr": DayNum = 5
        Case "sa": DayNum = 6
    End Select
End Function

Private Function MonNum(ByVal w As String) As Long
    MonNum = (InStr("janfebmaraprmayjunjulaugsepoctnovdec", LCase$(Left$(w, 3))) + 2) \ 3
End Function

Private Function NthNum(ByVal w As String) As Long
    Select Case LCase$(w)
        Case "first", "1st": NthNum = 1
        Case "second", "2nd": NthNum = 2
        Case "third", "3rd": NthNum = 3
        Case "fourth", "4th": NthNum = 4
        Case "last": NthNum = -1
    End Select
End Function

Private Sub AddNote(ByRef notes As String, ByVal t As String)
    notes = notes & IIf(notes = "", "", " ") & t
End Sub

Private Function Rx(ByVal pat As String, Optional ByVal everyOne As Boolean = False) As Object
    ' A pattern, made once and kept. Case doesn't matter.
    Dim key As String, r As Object
    If mRx Is Nothing Then Set mRx = CreateObject("Scripting.Dictionary")
    key = IIf(everyOne, "g", "1") & pat
    If Not mRx.Exists(key) Then
        Set r = CreateObject("VBScript.RegExp")
        r.Pattern = pat
        r.IgnoreCase = True
        r.Global = everyOne
        mRx.Add key, r
    End If
    Set Rx = mRx(key)
End Function

Private Function Find(ByVal s As String, ByVal pat As String, ByRef m As Object) As Boolean
    ' The first place the pattern matches, if it does
    Dim mc As Object
    Set mc = Rx(pat).Execute(s)
    If mc.Count > 0 Then
        Set m = mc(0)
        Find = True
    End If
End Function

Private Function SM(ByVal m As Object, ByVal i As Long) As String
    ' A match's i-th part, "" when that part didn't match
    SM = m.SubMatches(i - 1) & ""
End Function

Private Sub Cut(ByRef s As String, ByVal m As Object)
    ' Takes what matched out of the line, leaving a space
    s = Left$(s, m.FirstIndex) & " " & Mid$(s, m.FirstIndex + m.Length + 1)
End Sub

Private Function TrimAll(ByVal t As String) As String
    TrimAll = Trim$(Rx("^\s+|\s+$", True).Replace(Replace(t, ChrW(160), " "), ""))
End Function


' ---- days ------------------------------------------------------------------------

Private Function Dow(ByVal d As Long) As Long
    Dow = Weekday(d, vbSunday) - 1          ' 0 is Sunday
End Function

Private Function DaysIn(ByVal y As Long, ByVal mo As Long) As Long
    DaysIn = Day(DateSerial(y, mo + 1, 0))
End Function

Private Function IsOff(ByVal d As Long) As Boolean
    IsOff = Dow(d) = 0 Or Dow(d) = 6 Or mOff.Exists(d)
End Function

Private Function OffName(ByVal d As Long) As String
    If mOff.Exists(d) Then OffName = mOff(d) Else OffName = "a " & Split(DOW_LONG(), "|")(Dow(d))
End Function

Private Function WorkdayOnOrAfter(ByVal d As Long) As Long
    Do While IsOff(d): d = d + 1: Loop
    WorkdayOnOrAfter = d
End Function

Private Function WorkdayOnOrBefore(ByVal d As Long) As Long
    Do While IsOff(d): d = d - 1: Loop
    WorkdayOnOrBefore = d
End Function

Private Function NextWorkday(ByVal d As Long) As Long
    NextWorkday = WorkdayOnOrAfter(d + 1)
End Function

Private Function NextDow(ByVal d As Long, ByVal w As Long, ByVal after As Boolean) As Long
    If after Then d = d + 1
    Do While Dow(d) <> w: d = d + 1: Loop
    NextDow = d
End Function

Private Function AddMonthsKeepDay(ByVal d As Long, ByVal n As Long) As Long
    Dim y As Long, mo As Long, dd As Long
    y = Year(d): mo = Month(d) + n: dd = Day(d)
    Do While mo > 12: mo = mo - 12: y = y + 1: Loop
    If dd > DaysIn(y, mo) Then dd = DaysIn(y, mo)
    AddMonthsKeepDay = DateSerial(y, mo, dd)
End Function

Private Function DOW_LONG() As String
    DOW_LONG = "Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday"
End Function

Private Function FmtDay(ByVal d As Long) As String
    ' Tue 11/17
    FmtDay = Split("Sun|Mon|Tue|Wed|Thu|Fri|Sat", "|")(Dow(d)) & " " & Month(d) & "/" & Day(d)
End Function

Private Function FmtShort(ByVal d As Long) As String
    FmtShort = Month(d) & "/" & Day(d)
End Function

Private Function FmtTime(ByVal tm As Long) As String
    ' 2:30 PM
    Dim h As Long, ap As String
    If tm = NO_TIME Then Exit Function
    h = tm \ 60
    If h >= 12 Then ap = "PM" Else ap = "AM"
    h = h Mod 12
    If h = 0 Then h = 12
    FmtTime = h & ":" & Format$(tm Mod 60, "00") & " " & ap
End Function

Private Function Ord(ByVal n As Long) As String
    ' 1st, 2nd, 3rd, 11th
    Select Case n Mod 100
        Case 11, 12, 13: Ord = n & "th"
        Case Else
            Select Case n Mod 10
                Case 1: Ord = n & "st"
                Case 2: Ord = n & "nd"
                Case 3: Ord = n & "rd"
                Case Else: Ord = n & "th"
            End Select
    End Select
End Function


' ---- repeats ---------------------------------------------------------------------
' A repeat's rule sits in List's hidden Rule column as words and numbers, like
' "week every=2 dows=5 anchor=46325". Day numbers are Excel's. The raw date is where the rule
' lands; the due date is the workday on or before it, and the next one steps from the raw date.

Private Function RuleVal(ByVal rule As String, ByVal key As String) As String
    Dim p As Variant
    For Each p In Split(rule, " ")
        If Left$(p, Len(key) + 1) = key & "=" Then RuleVal = Mid$(p, Len(key) + 2): Exit Function
    Next
End Function

Private Function RuleNum(ByVal rule As String, ByVal key As String) As Long
    RuleNum = CLng(Val(RuleVal(rule, key)))
End Function

Private Function RuleUnit(ByVal rule As String) As String
    RuleUnit = Split(Trim$(rule) & " ", " ")(0)
End Function

Private Function MonthOK(ByVal rule As String, ByVal mo As Long) As Boolean
    ' Whether a month is one of the rule's months, for every 2 months, quarterly and the like
    Dim ev As Long
    ev = RuleNum(rule, "every")
    If ev <= 1 Then MonthOK = True: Exit Function
    MonthOK = (((mo - RuleNum(rule, "phase")) Mod ev) + ev) Mod ev = 0
End Function

Private Function SpecDate(ByVal sp As String, ByVal y As Long, ByVal mo As Long) As Long
    ' One day of the month in a given month: 15, last, lastwork, or 1:1 for the first Monday.
    ' 0 when that month has no such day.
    Dim last As Long, d As Long, nth As Long, wd As Long
    last = DaysIn(y, mo)
    If sp = "last" Then
        SpecDate = DateSerial(y, mo, last)
    ElseIf sp = "lastwork" Then
        SpecDate = WorkdayOnOrBefore(DateSerial(y, mo, last))
    ElseIf InStr(sp, ":") > 0 Then
        nth = CLng(Split(sp, ":")(0)): wd = CLng(Split(sp, ":")(1))
        If nth = -1 Then
            d = last
            Do While Dow(DateSerial(y, mo, d)) <> wd: d = d - 1: Loop
        Else
            d = 1
            Do While Dow(DateSerial(y, mo, d)) <> wd: d = d + 1: Loop
            d = d + 7 * (nth - 1)
        End If
        If d <= last Then SpecDate = DateSerial(y, mo, d)
    Else
        d = CLng(sp)
        If d > last Then d = last
        SpecDate = DateSerial(y, mo, d)
    End If
End Function

Private Function RuleFirst(ByVal rule As String, ByVal from As Long) As Long
    ' The first date on or after from that the rule lands on, 0 when there's none
    Dim x As Long, i As Long, y As Long, mo As Long, ev As Long, sp As Variant, d As Long, best As Long
    ev = RuleNum(rule, "every")
    Select Case RuleUnit(rule)
        Case "workday"
            RuleFirst = WorkdayOnOrAfter(from)
        Case "day"
            x = RuleNum(rule, "anchor")
            If x = 0 Then x = from
            Do While x < from: x = x + ev: Loop
            RuleFirst = x
        Case "week"
            If ev <= 1 Then
                x = from
                For i = 0 To 6
                    If InStr("," & RuleVal(rule, "dows") & ",", "," & Dow(x) & ",") > 0 Then RuleFirst = x: Exit Function
                    x = x + 1
                Next
            Else
                x = RuleNum(rule, "anchor")
                Do While x < from: x = x + 7 * ev: Loop
                RuleFirst = x
            End If
        Case "month"
            y = Year(from): mo = Month(from)
            For i = 0 To 47
                If MonthOK(rule, mo) Then
                    best = 0
                    For Each sp In Split(RuleVal(rule, "days"), ";")
                        d = SpecDate(CStr(sp), y, mo)
                        If d >= from And (best = 0 Or d < best) Then best = d
                    Next
                    If best > 0 Then RuleFirst = best: Exit Function
                End If
                mo = mo + 1
                If mo > 12 Then mo = 1: y = y + 1
            Next
        Case "year"
            mo = RuleNum(rule, "month")
            For y = Year(from) To Year(from) + 4
                d = RuleNum(rule, "dom")
                If d > DaysIn(y, mo) Then d = DaysIn(y, mo)
                If DateSerial(y, mo, d) >= from Then RuleFirst = DateSerial(y, mo, d): Exit Function
            Next
    End Select
End Function

Private Function RuleNext(ByVal rule As String, ByVal raw As Long) As Long
    ' The date after raw that the rule lands on
    Select Case RuleUnit(rule)
        Case "day": RuleNext = raw + RuleNum(rule, "every")
        Case "week"
            If RuleNum(rule, "every") > 1 Then RuleNext = raw + 7 * RuleNum(rule, "every") Else RuleNext = RuleFirst(rule, raw + 1)
        Case Else: RuleNext = RuleFirst(rule, raw + 1)
    End Select
End Function

Private Function FirstRaw(ByVal rule As String, ByVal from As Long, ByVal tm As Long) As Long
    ' A new repeat's first date: not one already past, not a time already gone today, not past until
    Dim raw As Long
    raw = RuleFirst(rule, from)
    If raw <> 0 And tm <> NO_TIME And raw = mToday And tm <= mMin Then raw = RuleNext(rule, raw)
    Do While raw <> 0
        If WorkdayOnOrBefore(raw) >= mToday Then Exit Do
        raw = RuleNext(rule, raw)
    Loop
    If raw <> 0 And RuleNum(rule, "until") > 0 And raw > RuleNum(rule, "until") Then raw = 0
    FirstRaw = raw
End Function

Private Function RuleLabel(ByVal rule As String) As String
    ' How a rule reads on List and Today: "weekly, Mondays", "quarterly, last day (Jan, Apr, Jul, Oct)"
    Dim t As String, ev As Long, p As Variant, part As String, mo As Long, mons As String
    If Trim$(rule) = "" Then Exit Function
    ev = RuleNum(rule, "every")
    Select Case RuleUnit(rule)
        Case "workday": t = "every workday"
        Case "day": t = "every " & ev & " days"
        Case "week"
            If ev = 1 Then t = "weekly, " Else t = "every " & ev & " weeks, "
            part = ""
            For Each p In Split(RuleVal(rule, "dows"), ",")
                part = part & IIf(part = "", "", " and ") & Split(DOW_LONG(), "|")(CLng(p)) & "s"
            Next
            t = t & part
        Case "month"
            If ev = 1 Then
                t = "monthly, "
            ElseIf ev = 3 Then
                t = "quarterly, "
            Else
                t = "every " & ev & " months, "
            End If
            part = ""
            For Each p In Split(RuleVal(rule, "days"), ";")
                part = part & IIf(part = "", "", " and ") & SpecLabel(CStr(p))
            Next
            t = t & part
            If ev > 1 Then
                For mo = 1 To 12
                    If MonthOK(rule, mo) Then mons = mons & IIf(mons = "", "", ", ") & _
                        Split("Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec", "|")(mo - 1)
                Next
                t = t & " (" & mons & ")"
            End If
        Case "year": t = "yearly, " & RuleNum(rule, "month") & "/" & RuleNum(rule, "dom")
    End Select
    If RuleNum(rule, "until") > 0 Then t = t & ", until " & FmtShort(RuleNum(rule, "until"))
    RuleLabel = t
End Function

Private Function SpecLabel(ByVal sp As String) As String
    If sp = "last" Then
        SpecLabel = "last day"
    ElseIf sp = "lastwork" Then
        SpecLabel = "last workday"
    ElseIf InStr(sp, ":") > 0 Then
        SpecLabel = Split("last|first|second|third|fourth", "|")(IIf(CLng(Split(sp, ":")(0)) = -1, 0, CLng(Split(sp, ":")(0)))) & _
            " " & Split(DOW_LONG(), "|")(CLng(Split(sp, ":")(1)))
    Else
        SpecLabel = "the " & Ord(CLng(sp))
    End If
End Function


Private Sub FillDaysOff(ByVal y As Long)
    ' The usual 8, for this year and next, each on the weekday it's taken
    Dim out As Variant, k As Long, yy As Long, xmas As Long, eve As Long, tg As Long
    ReDim out(1 To 16, 1 To 2)
    For yy = y To y + 1
        xmas = Observed(DateSerial(yy, 12, 25))
        eve = xmas - 1
        Do While Dow(eve) = 0 Or Dow(eve) = 6: eve = eve - 1: Loop
        tg = NthDow(yy, 11, 4, 4)
        PutOff out, k, Observed(DateSerial(yy, 1, 1)), "New Year's Day"
        PutOff out, k, LastDow(yy, 5, 1), "Memorial Day"
        PutOff out, k, Observed(DateSerial(yy, 7, 4)), "the July 4 holiday"
        PutOff out, k, NthDow(yy, 9, 1, 1), "Labor Day"
        PutOff out, k, tg, "Thanksgiving"
        PutOff out, k, tg + 1, "the day after Thanksgiving"
        PutOff out, k, eve, "Christmas Eve"
        PutOff out, k, xmas, "Christmas"
    Next
    mOffT.Resize mOffT.HeaderRowRange.Resize(17)
    mOffT.HeaderRowRange.Offset(1).Resize(16).Value = out
End Sub

Private Sub PutOff(ByRef out As Variant, ByRef k As Long, ByVal d As Long, ByVal nm As String)
    k = k + 1
    out(k, 1) = CDbl(d)
    out(k, 2) = nm
End Sub

Private Function Observed(ByVal d As Long) As Long
    ' A holiday on a Saturday is taken Friday; on a Sunday, Monday
    Select Case Dow(d)
        Case 6: Observed = d - 1
        Case 0: Observed = d + 1
        Case Else: Observed = d
    End Select
End Function

Private Function NthDow(ByVal y As Long, ByVal mo As Long, ByVal w As Long, ByVal n As Long) As Long
    NthDow = NextDow(DateSerial(y, mo, 1), w, False) + 7 * (n - 1)
End Function

Private Function LastDow(ByVal y As Long, ByVal mo As Long, ByVal w As Long) As Long
    Dim d As Long
    d = DateSerial(y, mo, DaysIn(y, mo))
    Do While Dow(d) <> w: d = d - 1: Loop
    LastDow = d
End Function

Private Function NowT() As Double
    ' The time the buttons go by. A test can set it.
    If mClock > 0 Then NowT = mClock Else NowT = CDbl(Now)
End Function


' ---- the rows ----------------------------------------------------------------------

Private Function Ready() As Boolean
    ' Finds the tables and reads List in. Says what to do when setup hasn't run.
    Dim v As Variant, i As Long, j As Long, n As Long, t As Double
    Set mList = FindTable(LIST_TABLE)
    Set mLogT = FindTable(LOG_TABLE)
    Set mOffT = FindTable(OFF_TABLE)
    If mList Is Nothing Or mLogT Is Nothing Or mOffT Is Nothing Or Not SheetExists(TODAY_SHEET) Then GoTo NotSet
    If Not HasColumns(mList, LIST_COLS) Or Not HasColumns(mLogT, LOG_COLS) Or Not HasColumns(mOffT, OFF_COLS) Then GoTo NotSet
    t = NowT()
    mToday = Int(t)
    mMin = Int((t - mToday) * 1440 + 0.000001)
    If mMin > 1439 Then mMin = 1439
    cId = ColOf(mList, "#"): cWhat = ColOf(mList, "What"): cDue = ColOf(mList, "Due")
    cTime = ColOf(mList, "Time"): cRep = ColOf(mList, "Repeats"): cWait = ColOf(mList, "Waiting on")
    cFrom = ColOf(mList, "Came from"): cOff = ColOf(mList, "Off the list"): cOffOn = ColOf(mList, "Date off")
    cRule = ColOf(mList, "Rule"): cRaw = ColOf(mList, "Raw date"): cTask = ColOf(mList, "Outlook task")
    cMail = ColOf(mList, "Email"): cCal = ColOf(mList, "Calendar")
    mNCols = mList.ListColumns.Count
    n = mList.ListRows.Count
    ReDim mL(1 To mNCols, 1 To n + 16)
    mN = 0
    mNextId = 1
    If n > 0 Then
        v = mList.DataBodyRange.Value2
        For i = 1 To n
            If Trim$(CellText(v(i, cWhat))) <> "" Or Trim$(CellText(v(i, cId))) <> "" Then
                mN = mN + 1
                For j = 1 To mNCols
                    mL(j, mN) = v(i, j)
                Next
                If IsNumeric(mL(cId, mN)) And Not IsEmpty(mL(cId, mN)) Then
                    If CLng(mL(cId, mN)) >= mNextId Then mNextId = CLng(mL(cId, mN)) + 1
                End If
            End If
        Next
    End If
    For i = 1 To mN                                     ' a row typed in by hand gets a number
        If Not IsNumeric(mL(cId, i)) Or IsEmpty(mL(cId, i)) Then
            mL(cId, i) = mNextId
            mNextId = mNextId + 1
        End If
    Next
    LoadDaysOff
    Set mPend = New Collection
    Set mKeep = New Collection
    Set mGone = New Collection
    mKeepName = "": mKeepVerb = ""
    Ready = True
    Exit Function
NotSet:
    Say "Run TrackerSetup first: press Alt F8, pick TrackerSetup and press Run. It keeps every row you have."
End Function

Private Sub LoadDaysOff()
    Dim v As Variant, i As Long
    Set mOff = CreateObject("Scripting.Dictionary")
    If mOffT.ListRows.Count = 0 Then Exit Sub
    v = mOffT.DataBodyRange.Value2
    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
            If Not mOff.Exists(CLng(v(i, 1))) Then mOff.Add CLng(v(i, 1)), Trim$(CellText(v(i, 2)))
        End If
    Next
End Sub

Private Function NewRow() As Long
    ' A blank row at the end of mL with the next number
    If mN = UBound(mL, 2) Then ReDim Preserve mL(1 To mNCols, 1 To mN * 2 + 16)
    mN = mN + 1
    mL(cId, mN) = mNextId
    mNextId = mNextId + 1
    NewRow = mN
End Function

Private Function RowOf(ByVal id As Long) As Long
    Dim i As Long
    If id <= 0 Then Exit Function
    For i = 1 To mN
        If IsNumeric(mL(cId, i)) And Not IsEmpty(mL(cId, i)) Then
            If CLng(mL(cId, i)) = id Then RowOf = i: Exit Function
        End If
    Next
End Function

Private Function PickedId() As Long
    ' The number of the row clicked on Today or List, 0 when it isn't one
    Dim ws As Worksheet, c As Range, v As Variant
    Set ws = ActiveSheet
    Set c = ActiveCell
    If c Is Nothing Then Exit Function
    If ws.Name = TODAY_SHEET And ws.Parent Is ThisWorkbook Then
        If c.Row < FIRST_ROW Then Exit Function
        v = ws.Cells(c.Row, 1).Value2
    ElseIf ws.Name = LIST_SHEET And ws.Parent Is ThisWorkbook Then
        If mList.ListRows.Count = 0 Then Exit Function
        If Intersect(c, mList.DataBodyRange) Is Nothing Then Exit Function
        v = ws.Cells(c.Row, mList.ListColumns(cId).Range.Column).Value2
    Else
        Exit Function
    End If
    If IsNumeric(v) And Not IsEmpty(v) Then PickedId = CLng(v)
End Function

Private Function OpenRow(ByVal verb As String, ByRef id As Long, ByRef k As Long) As Boolean
    ' The clicked row, when it's one that's still on the list. Says what to do when it isn't.
    If Not Ready() Then Exit Function
    id = PickedId()
    k = RowOf(id)
    If k = 0 Then
        Say "Click a row first, then press " & verb & "."
        Exit Function
    End If
    If Trim$(CellText(mL(cOff, k))) <> "" Then
        Say """" & CellText(mL(cWhat, k)) & """ is already off the list."
        Exit Function
    End If
    OpenRow = True
End Function

Private Sub Keep(ByVal k As Long, ByVal named As String, ByVal verb As String)
    ' Before a row changes: keep it as it is, so Undo can put it back. k = 0 keeps no row.
    Dim vals As Variant, j As Long, e As Variant
    If named <> "" Then mKeepName = named: mKeepVerb = verb
    If k = 0 Then Exit Sub
    For Each e In mKeep
        If e(0) = mL(cId, k) Then Exit Sub
    Next
    ReDim vals(1 To mNCols)
    For j = 1 To mNCols
        vals(j) = mL(j, k)
    Next
    mKeep.Add Array(CLng(mL(cId, k)), vals)
End Sub

Private Sub DropRowAt(ByVal k As Long)
    ' Takes a row out of mL; the last one fills its place, since List gets sorted anyway
    Dim j As Long
    For j = 1 To mNCols
        mL(j, k) = mL(j, mN)
        If mN <> k Then mL(j, mN) = Empty
    Next
    If mN = k Then
        For j = 1 To mNCols: mL(j, k) = Empty: Next
    End If
    mN = mN - 1
End Sub

Private Sub RemoveLogLines(ByVal lines As Collection)
    ' Takes the given lines off the Log. New lines are on top, so that's where they're found.
    Dim v As Variant, n As Long, i As Long, j As Long, a As Variant, hit() As Boolean, nc As Long
    Dim cW As Long, cI As Long, cH As Long
    If lines Is Nothing Then Exit Sub
    If lines.Count = 0 Or mLogT.ListRows.Count = 0 Then Exit Sub
    cW = ColOf(mLogT, "When"): cI = ColOf(mLogT, "#"): cH = ColOf(mLogT, "What happened")
    nc = mLogT.ListColumns.Count
    v = mLogT.DataBodyRange.Value2
    If Not IsArray(v) Then Exit Sub
    n = UBound(v, 1)
    ReDim hit(1 To n)
    For Each a In lines
        For i = 1 To n
            If Not hit(i) Then
                If Abs(NumOf(v(i, cW)) - a(0)) < 0.00001 And NumOf(v(i, cI)) = a(1) And CellText(v(i, cH)) = a(3) Then
                    hit(i) = True
                    Exit For
                End If
            End If
        Next
    Next
    i = n
    Do While i >= 1                                     ' each run of lines goes in one delete, bottom up
        If hit(i) Then
            j = i
            Do While j > 1
                If Not hit(j - 1) Then Exit Do
                j = j - 1
            Loop
            mLogT.DataBodyRange.Parent.Range(mLogT.DataBodyRange.Cells(j, 1), mLogT.DataBodyRange.Cells(i, nc)).Delete Shift:=xlShiftUp
            i = j - 1
        Else
            i = i - 1
        End If
    Loop
End Sub

Private Sub ForgetUndo(ByVal why As String)
    ' A change made in Outlook can't be taken back here, so Undo lets go of the last change
    ' and says why instead
    Set mUndo = Nothing
    Set mUndoLog = Nothing
    mUndoWhy = why
End Sub

Private Sub LogLine(ByVal id As Long, ByVal what As String, ByVal text As String, ByVal kind As String)
    mPend.Add Array(NowT(), id, what, text, kind)
End Sub

Private Sub Finish()
    ' Writes List and the new Log lines back, then redraws Today. The change is what Undo takes back.
    Dim a As Variant
    If Not mUndoing And (mKeep.Count > 0 Or mPend.Count > 0) Then
        Set mUndo = mKeep
        Set mUndoLog = New Collection
        For Each a In mPend
            mUndoLog.Add a
        Next
        mUndoName = mKeepName
        mUndoVerb = mKeepVerb
        mUndoWhy = ""
    End If
    Application.ScreenUpdating = False
    PushOutlook False
    SaveList
    SaveLog
    DrawToday
    Application.ScreenUpdating = True
End Sub

Private Sub SaveList()
    ' List goes back sorted: open rows first, then by due date, then by number
    Dim keys As Variant, idx As Variant, out As Variant, i As Long, j As Long, oldN As Long
    Dim ws As Worksheet, top As Long
    Set ws = mList.Parent
    top = mList.HeaderRowRange.Row
    oldN = mList.ListRows.Count
    If mN = 0 Then
        If oldN > 0 Then mList.DataBodyRange.Delete
        Exit Sub
    End If
    ReDim keys(1 To mN): ReDim idx(1 To mN)
    For i = 1 To mN
        idx(i) = i
        keys(i) = IIf(Trim$(CellText(mL(cOff, i))) <> "", 1, 0) * 10000000000000# + NumOf(mL(cDue, i)) * 1000000# + NumOf(mL(cId, i))
    Next
    SortBy keys, idx, 1, mN
    ReDim out(1 To mN, 1 To mNCols)
    For i = 1 To mN
        For j = 1 To mNCols
            out(i, j) = mL(j, idx(i))
        Next
    Next
    mList.Resize ws.Range(ws.Cells(top, mList.Range.Column), ws.Cells(top + mN, mList.Range.Column + mNCols - 1))
    mList.HeaderRowRange.Offset(1).Resize(mN).Value = out    ' a table with no rows has no body to write to yet
    If oldN > mN Then ws.Range(ws.Cells(top + mN + 1, mList.Range.Column), _
        ws.Cells(top + oldN, mList.Range.Column + mNCols - 1)).ClearContents
    ' keep mL in the new order, so row numbers on the sheet match it
    For i = 1 To mN
        For j = 1 To mNCols
            mL(j, i) = out(i, j)
        Next
    Next
End Sub

Private Sub SaveLog()
    ' New lines go in at the top, newest first
    Dim k As Long, out As Variant, i As Long, a As Variant, nc As Long
    Dim cW As Long, cI As Long, cT As Long, cH As Long, cK As Long
    k = mPend.Count
    If k = 0 Then Exit Sub
    nc = mLogT.ListColumns.Count
    cW = ColOf(mLogT, "When"): cI = ColOf(mLogT, "#"): cT = ColOf(mLogT, "What")
    cH = ColOf(mLogT, "What happened"): cK = ColOf(mLogT, "Kind")
    ReDim out(1 To k, 1 To nc)
    For i = 1 To k
        a = mPend(k - i + 1)
        out(i, cW) = a(0): out(i, cT) = a(2): out(i, cH) = a(3): out(i, cK) = a(4)
        If a(1) <> 0 Then out(i, cI) = a(1)            ' a row Undo took back has no number left
    Next
    If mLogT.ListRows.Count = 0 Then
        mLogT.Resize mLogT.HeaderRowRange.Resize(k + 1)
    Else
        mLogT.DataBodyRange.Resize(k).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    End If
    mLogT.HeaderRowRange.Offset(1).Resize(k).Value = out
    Set mPend = New Collection
End Sub

Private Sub SortBy(ByRef keys As Variant, ByRef idx As Variant, ByVal lo As Long, ByVal hi As Long)
    ' Sorts idx by keys, smallest first. Equal keys keep their order (keys end in the row number).
    Dim i As Long, j As Long, p As Double, tk As Variant, ti As Variant
    Do While lo < hi
        i = lo: j = hi
        p = keys((lo + hi) \ 2)
        Do While i <= j
            Do While keys(i) < p: i = i + 1: Loop
            Do While keys(j) > p: j = j - 1: Loop
            If i <= j Then
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
                ti = idx(i): idx(i) = idx(j): idx(j) = ti
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then SortBy keys, idx, lo, j
            lo = i
        Else
            If i < hi Then SortBy keys, idx, i, hi
            hi = j
        End If
    Loop
End Sub

Private Function NumOf(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) Then NumOf = CDbl(v)
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    CellText = v & ""
End Function

Private Function TimeOf(ByVal v As Variant) As Long
    ' A Time cell as minutes after midnight, NO_TIME when blank
    If IsEmpty(v) Or IsError(v) Or Not IsNumeric(v) Then TimeOf = NO_TIME: Exit Function
    If CellText(v) = "" Then TimeOf = NO_TIME: Exit Function
    TimeOf = CLng(Int((CDbl(v) - Int(CDbl(v))) * 1440 + 0.5)) Mod 1440
End Function


' ---- Outlook ----------------------------------------------------------------------------
' Popups are Outlook tasks with a reminder, in a Tracker folder under Tasks, each subject
' starting "Tracker: ". The tracker reads only a task's ID, subject, dates, reminder, whether it's
' complete, and two plain boxes it fills in itself (Mileage, so it knows its own, and Billing
' information on the morning list). Of an email it reads the subject, the flag and its date, its
' folder and its IDs; of a calendar item the subject, times, all day or not, how it
' repeats and whether it's a meeting. It never reads who an item is from or to, who's invited, or
' what it says. It only uses an Outlook that's already open, and never starts one.

Private Function OutlookNow() As Object
    On Error Resume Next
    Set OutlookNow = GetObject(, "Outlook.Application")
    Err.Clear
End Function

Private Function OutlookOpen(ByVal verb As String, ByRef ns As Object) As Boolean
    ' Outlook's folders in ns, when Outlook is open. When it isn't, says to open it and press verb again.
    Dim ol As Object
    Set ol = OutlookNow()
    If ol Is Nothing Then
        If verb = "" Then
            Say "Open Outlook first, then run this again."
        Else
            Say "Open Outlook first, then press " & verb & " again."
        End If
        Exit Function
    End If
    Set ns = ol.GetNamespace("MAPI")
    OutlookOpen = True
End Function

Private Function UseMainTasks() As Boolean
    Dim v As String
    On Error Resume Next
    v = ThisWorkbook.Names(FOLDER_SETTING).RefersTo
    UseMainTasks = InStr(1, v, "main", vbTextCompare) > 0
End Function

Private Function TaskFolder(ByVal ns As Object) As Object
    ' The Tracker folder under Tasks, made the first time; or Tasks itself, after TrackerUseMainTasks
    Dim f As Object, sf As Object
    Set f = ns.GetDefaultFolder(13)                     ' Tasks
    If UseMainTasks() Then
        Set TaskFolder = f
        Exit Function
    End If
    On Error Resume Next
    Set sf = f.Folders(TRACKER_FOLDER)
    On Error GoTo 0
    If sf Is Nothing Then Set sf = f.Folders.Add(TRACKER_FOLDER, 13)
    Set TaskFolder = sf
End Function

Private Function ItemById(ByVal ns As Object, ByVal id As String) As Object
    On Error Resume Next
    Set ItemById = ns.GetItemFromID(id)
End Function

Private Function LinkId(ByVal link As String) As String
    If InStr(link, "|") > 0 Then LinkId = Left$(link, InStr(link, "|") - 1) Else LinkId = link
End Function

Private Function LinkKey(ByVal link As String) As Long
    If InStr(link, "|") > 0 Then LinkKey = CLng(Val(Mid$(link, InStr(link, "|") + 1)))
End Function

Private Function LinkItem(ByVal link As String) As String
    ' An email link is "message ID|item ID"; this is the item ID
    If InStr(link, "|") > 0 Then LinkItem = Mid$(link, InStr(link, "|") + 1)
End Function

' ---- Outlook: the calendar and flagged emails

Private Function CalTime(ByVal c As Object) As Long
    ' A calendar item's start time as minutes, NO_TIME when it takes all day
    Dim s As Double
    If c.AllDayEvent Then CalTime = NO_TIME: Exit Function
    s = CDbl(c.Start)
    CalTime = CLng(Int((s - Int(s)) * 1440 + 0.5)) Mod 1440
End Function

Private Function CalRule(ByVal c As Object, ByRef rule As String) As Boolean
    ' A calendar series' repeat as the tracker's rule, from its pattern and its first date, so it
    ' lands on the same days the calendar has. False when it repeats in a way the tracker can't follow.
    Dim p As Object, ev As Long, mask As Long, dows As String, w As Long, first As Long, sp As String
    rule = ""
    Set p = c.GetRecurrencePattern
    first = CLng(Int(CDbl(c.Start)))
    ev = p.Interval
    If ev < 1 Then ev = 1
    mask = p.DayOfWeekMask
    For w = 0 To 6
        If (mask And 2 ^ w) <> 0 Then dows = dows & IIf(dows = "", "", ",") & w
    Next
    Select Case p.RecurrenceType
        Case 0                                          ' daily, or every weekday
            If ev = 1 Or mask = 62 Then rule = "workday" Else rule = "day every=" & ev & " anchor=" & first
        Case 1                                          ' weekly
            If dows = "" Then dows = CStr(Dow(first))
            If ev = 1 And mask = 62 Then
                rule = "workday"
            ElseIf ev = 1 Then
                rule = "week every=1 dows=" & dows
            ElseIf InStr(dows, ",") = 0 Then
                rule = "week every=" & ev & " dows=" & dows & " anchor=" & first
            Else
                Exit Function                           ' every other week on two days
            End If
        Case 2                                          ' monthly, on a day of the month
            rule = "month every=" & ev & " days=" & p.DayOfMonth
            If ev > 1 Then rule = rule & " phase=" & Month(first)
        Case 3                                          ' monthly, like the first Monday
            sp = NthSpec(p.Instance, mask)
            If sp = "" Then Exit Function
            rule = "month every=" & ev & " days=" & sp
            If ev > 1 Then rule = rule & " phase=" & Month(first)
        Case 5                                          ' yearly on a date; Outlook counts a year as 12
            If ev <> 1 And ev <> 12 Then Exit Function
            rule = "year month=" & p.MonthOfYear & " dom=" & p.DayOfMonth
        Case 6                                          ' yearly, like the first Monday of November
            If ev <> 1 And ev <> 12 Then Exit Function
            sp = NthSpec(p.Instance, mask)
            If sp = "" Then Exit Function
            rule = "month every=12 days=" & sp & " phase=" & p.MonthOfYear
        Case Else
            Exit Function
    End Select
    If Not p.NoEndDate Then rule = rule & " until=" & CLng(Int(CDbl(p.PatternEndDate)))
    CalRule = True
End Function

Private Function NthSpec(ByVal inst As Long, ByVal mask As Long) As String
    ' Outlook's first to fourth (5 is last) of a day as a day of the month: 1:1 is the first Monday,
    ' lastwork the last weekday, last the last day. "" when the tracker has no such day.
    Dim w As Long, n As Long
    If inst < 1 Or inst > 5 Then Exit Function
    If inst = 5 Then n = -1 Else n = inst
    Select Case mask
        Case 127: If n = -1 Then NthSpec = "last" Else NthSpec = CStr(n)
        Case 62: If n = -1 Then NthSpec = "lastwork"
        Case Else
            For w = 0 To 6
                If mask = 2 ^ w Then NthSpec = n & ":" & w
            Next
    End Select
End Function

Private Sub ListCalendar(ByVal ns As Object, ByVal found As Collection, ByRef n As Long)
    ' The calendar's series that haven't ended, then its one-time things from today on, leaving out
    ' what's already come in. Each is Array(Yes or No, what, when in Outlook, why, where it is).
    Dim cal As Object, its As Object, c As Object, have As Object, k As Long, rule As String, ok As Boolean
    Dim tm As Long, take As String, why As String, where As String, p As Object
    Set have = CreateObject("Scripting.Dictionary")
    For k = 1 To mN
        If CellText(mL(cCal, k)) <> "" Then have(CellText(mL(cCal, k))) = True
    Next
    Set cal = ns.GetDefaultFolder(9)
    Set its = cal.Items.Restrict("[IsRecurring] = True")
    For Each c In its
        If Not have.Exists(c.EntryID) Then
            Set p = c.GetRecurrencePattern
            If p.NoEndDate Or Int(CDbl(p.PatternEndDate)) >= mToday Then
                ok = CalRule(c, rule)
                tm = CalTime(c)
                If ok Then where = RuleLabel(rule) Else where = "repeats"
                where = where & ", " & IIf(tm = NO_TIME, "all day", FmtTime(tm))
                If c.MeetingStatus <> 0 Then
                    take = "No": why = "Meeting, stays on the calendar"
                ElseIf Not ok Then
                    take = "No": why = "Repeats in a way the tracker can't follow, so it stays on the calendar"
                Else
                    take = "Yes": why = "Repeats"
                End If
                found.Add Array(take, CStr(c.Subject), where, why, "cal|" & c.EntryID)
                n = n + 1
            End If
            Set p = Nothing
        End If
    Next
    Set its = cal.Items.Restrict("[IsRecurring] = False And [Start] >= '" & Format$(CDate(mToday), "ddddd h:nn AMPM") & "'")
    For Each c In its
        If Not have.Exists(c.EntryID) Then
            tm = CalTime(c)
            where = FmtDay(CLng(Int(CDbl(c.Start)))) & ", " & IIf(tm = NO_TIME, "all day", FmtTime(tm))
            If c.MeetingStatus <> 0 Then
                take = "No": why = "Meeting, stays on the calendar"
            ElseIf tm <> NO_TIME Then
                take = "No": why = "Set time, stays on the calendar"
            Else
                take = "Yes": why = "All day"
            End If
            found.Add Array(take, CStr(c.Subject), where, why, "cal|" & c.EntryID)
            n = n + 1
        End If
    Next
End Sub

Private Function TakeCalendar(ByVal ns As Object, ByVal id As String, ByRef notes As String) As String
    ' One calendar item comes in as a row: a series on its next date, a one-time thing on its day, or
    ' the workday before when that's a weekend or day off, which notes then says.
    ' Says took, ended, cant (a repeat the tracker can't follow) or gone.
    Dim c As Object, rule As String, tm As Long, first As Long, raw As Long, due As Long, k As Long, what As String
    Dim why As String
    Set c = ItemById(ns, id)
    If c Is Nothing Then TakeCalendar = "gone": Exit Function
    tm = CalTime(c)
    first = CLng(Int(CDbl(c.Start)))
    If c.IsRecurring Then
        If Not CalRule(c, rule) Then TakeCalendar = "cant": Exit Function
        If first < mToday Then first = mToday
        raw = FirstRaw(rule, first, tm)
        If raw = 0 Then TakeCalendar = "ended": Exit Function
        due = WorkdayOnOrBefore(raw)
    Else
        due = WorkdayOnOrBefore(first)
        If due <> first Then why = FmtShort(first) & " is " & OffName(first) & ", so it shows " & FmtDay(due) & "."
    End If
    what = Trim$(CStr(c.Subject))
    If what = "" Then what = "(no subject)"
    If why <> "" Then AddNote notes, "For """ & what & """, " & why
    k = NewRow()
    mKeep.Add Array(CLng(mL(cId, k)), Empty)
    mL(cWhat, k) = what
    mL(cDue, k) = CDbl(due)
    If tm <> NO_TIME Then mL(cTime, k) = tm / 1440
    mL(cFrom, k) = "calendar"
    mL(cCal, k) = CStr(c.EntryID)
    If rule <> "" Then
        mL(cRule, k) = rule
        mL(cRaw, k) = CDbl(raw)
        mL(cRep, k) = RuleLabel(rule)
    End If
    LogLine CLng(mL(cId, k)), what, "Came in from your calendar." & IIf(why <> "", " " & why, ""), "added"
    TakeCalendar = "took"
End Function

Private Function FlagRows(ByVal ns As Object) As Collection
    ' The flagged emails in Inbox and Sent Items, each Array(item ID, message ID, 6 for Inbox or 5 for Sent)
    Dim f As Variant, tb As Object, r As Object, c As New Collection
    For Each f In Array(6, 5)
        Set tb = ns.GetDefaultFolder(f).GetTable("@SQL=""" & FLAG_STATUS & """ > 1")
        tb.Columns.Add MSG_ID
        Do While Not tb.EndOfTable
            Set r = tb.GetNextRow
            c.Add Array(CStr(r("EntryID")), r(MSG_ID) & "", f)
        Loop
    Next
    Set FlagRows = c
End Function

Private Function OldFlags() As Object
    ' The emails that were flagged when the tracker first looked, by message ID and item ID, so a
    ' change to one later doesn't make it new. Nothing before that first look.
    Dim parts As Object, d As Object, w As Variant
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(OLD_FLAGS)
    If parts.Count = 0 Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    For Each w In Split(parts(1).DocumentElement.Text, " ")
        If w <> "" Then d(CStr(w)) = True
    Next
    Set OldFlags = d
End Function

Private Sub KeepOldFlags(ByVal rows As Collection)
    ' The first look: every email flagged now is an old one, kept in the workbook
    Dim t As String, r As Variant
    For Each r In rows
        t = t & " " & r(0) & " " & Replace(Replace(Replace(r(1), "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
    Next
    ThisWorkbook.CustomXMLParts.Add "<oldFlags xmlns=""" & OLD_FLAGS & """>" & t & "</oldFlags>"
End Sub

Private Function IsOldFlag(ByVal old As Object, ByVal r As Variant) As Boolean
    If old Is Nothing Then IsOldFlag = True: Exit Function       ' before the first look, every flag is old
    If old.Exists(CStr(r(0))) Then IsOldFlag = True: Exit Function
    If r(1) <> "" Then IsOldFlag = old.Exists(CStr(r(1)))
End Function

Private Function FlagDay(ByVal m As Object) As Long
    ' The day an email's flag is for, 0 when it has none
    If Year(m.TaskDueDate) < 4000 Then FlagDay = CLng(Int(CDbl(m.TaskDueDate)))
End Function

Private Function MailRow(ByVal link As String) As Long
    ' A row still on the list that came from this email: the same message ID, or the same item ID
    Dim k As Long, have As String
    For k = 1 To mN
        If Trim$(CellText(mL(cOff, k))) = "" Then
            have = CellText(mL(cMail, k))
            If have <> "" Then
                If LinkId(have) <> "" And LinkId(have) = LinkId(link) Then MailRow = k: Exit Function
                If LinkItem(have) <> "" And LinkItem(have) = LinkItem(link) Then MailRow = k: Exit Function
            End If
        End If
    Next
End Function

Private Function AddMailRow(ByVal subj As String, ByVal link As String, ByVal flagOn As Long, ByVal isSent As Boolean, _
                            ByVal how As String, ByRef why As String) As Long
    ' A row for a flagged email: due the workday on or after the flag's day (today when it has none
    ' or it's gone by); a sent one waits on a reply, checked back that day. why says so when a
    ' weekend or day off moved it.
    Dim k As Long, due As Long, what As String
    due = flagOn
    If due < mToday Then due = mToday
    due = WorkdayOnOrAfter(due)
    why = ""
    If flagOn >= mToday And due <> flagOn Then why = FmtShort(flagOn) & " is " & OffName(flagOn) & ", so it shows the next workday."
    If why <> "" Then how = how & " " & why
    what = Trim$(Rx("^\s*(?:(?:re|fw|fwd)\s*:\s*)+").Replace(subj, ""))
    If what = "" Then what = "(no subject)"
    what = UCase$(Left$(what, 1)) & Mid$(what, 2)
    k = NewRow()
    mKeep.Add Array(CLng(mL(cId, k)), Empty)
    mL(cWhat, k) = what
    mL(cDue, k) = CDbl(due)
    If isSent Then mL(cWait, k) = "a reply"
    mL(cFrom, k) = "flagged email"
    mL(cMail, k) = link
    LogLine CLng(mL(cId, k)), what, how, "added"
    AddMailRow = k
End Function

Private Sub ListOldFlags(ByVal ns As Object, ByVal found As Collection, ByRef n As Long)
    ' Emails in Inbox and Sent Items flagged before the tracker first looked and not in it yet
    Dim r As Variant, old As Object, m As Object, link As String, d As Long
    Set old = OldFlags()
    For Each r In FlagRows(ns)
        If IsOldFlag(old, r) Then
            link = r(1) & "|" & r(0)
            If MailRow(link) = 0 Then
                Set m = ItemById(ns, CStr(r(0)))
                If Not m Is Nothing Then
                    d = FlagDay(m)
                    found.Add Array("No", CStr(m.Subject), m.Parent.Name & ", flagged " & _
                        IIf(d = 0, "with no date", "for " & FmtDay(d)), _
                        "Flagged before the tracker. Say Yes to bring it in.", "mail|" & link)
                    n = n + 1
                End If
                Set m = Nothing
            End If
        End If
    Next
End Sub

Private Function TakeOldFlag(ByVal ns As Object, ByVal link As String, ByRef notes As String) As Boolean
    ' An email flagged before the tracker, said Yes to on From Outlook: a row, and its flag a check
    Dim m As Object, why As String, k As Long
    If MailRow(link) > 0 Then Exit Function
    Set m = FindEmail(ns, link)
    If m Is Nothing Then Exit Function
    k = AddMailRow(CStr(m.Subject), link, FlagDay(m), m.Parent.EntryID = ns.GetDefaultFolder(5).EntryID, _
        "Came in from an email you had flagged.", why)
    If why <> "" Then AddNote notes, "For """ & CellText(mL(cWhat, k)) & """, " & why
    m.FlagStatus = 1                                    ' the flag turns into a check
    m.Save
    TakeOldFlag = True
End Function

Private Function TakeFlags(ByVal ns As Object, ByRef said As String) As Long
    ' Emails flagged since the tracker first looked come in by themselves: an Inbox flag is due on
    ' the flag's day, a Sent Items flag waits on a reply. Each flag then turns into a check.
    Dim old As Object, rows As Collection, r As Variant, link As String, m As Object, k As Long, subj As String
    Dim due As Long, why As String
    On Error GoTo Trouble                               ' rows already made stay; their flags get checked next time
    Set rows = FlagRows(ns)
    Set old = OldFlags()
    If old Is Nothing Then                              ' the first look: every flag set so far is an old one
        KeepOldFlags rows
        Exit Function
    End If
    For Each r In rows
        If Not IsOldFlag(old, r) Then
            Set m = ItemById(ns, CStr(r(0)))
            If Not m Is Nothing Then
                link = r(1) & "|" & r(0)
                If MailRow(link) = 0 Then
                    subj = CStr(m.Subject)
                    k = AddMailRow(subj, link, FlagDay(m), r(2) = 5, "Came in from a flagged email.", why)
                    due = CLng(mL(cDue, k))
                    If why <> "" Then why = " " & why
                    If r(2) = 5 Then
                        said = said & IIf(said = "", "", " ") & "You flagged your sent """ & subj & _
                            """, so it's under Waiting on them, check back " & FmtDay(due) & "." & why
                    Else
                        said = said & IIf(said = "", "", " ") & "You flagged """ & subj & """, so it's due " & _
                            IIf(due = mToday, "today", FmtDay(due)) & "." & why & " The flag turns into a check so it " & _
                            "doesn't nag twice. Open email takes you back to it."
                    End If
                    TakeFlags = TakeFlags + 1
                End If
                m.FlagStatus = 1                        ' the flag turns into a check
                m.Save
            End If
            Set m = Nothing
        End If
    Next
    Exit Function
Trouble:
    mOlNote = "(Outlook didn't give the flagged emails: " & Err.Description & ")"
End Function

Private Function FindEmail(ByVal ns As Object, ByVal link As String) As Object
    ' An email from its link: by its message ID, which stays the same when it's filed, in Inbox,
    ' Sent Items, then every other mail folder; by its item ID when that finds nothing
    Dim msgId As String, flt As String
    msgId = LinkId(link)
    If msgId <> "" Then
        flt = "@SQL=""" & MSG_ID & """ = '" & Replace(msgId, "'", "''") & "'"
        Set FindEmail = FindIn(ns.GetDefaultFolder(6), flt, False)
        If FindEmail Is Nothing Then Set FindEmail = FindIn(ns.GetDefaultFolder(5), flt, False)
        If FindEmail Is Nothing Then Set FindEmail = FindIn(ns.GetDefaultFolder(6).Parent, flt, True)
    End If
    If FindEmail Is Nothing And LinkItem(link) <> "" Then Set FindEmail = ItemById(ns, LinkItem(link))
End Function

Private Function FindIn(ByVal fld As Object, ByVal flt As String, ByVal deep As Boolean) As Object
    ' The first email in a folder (and the folders under it, when deep) that the filter finds
    Dim its As Object, sf As Object, hit As Object
    On Error Resume Next
    If fld.DefaultItemType = 0 Then                     ' a mail folder
        Set its = fld.Items.Restrict(flt)
        If Err.Number = 0 Then
            If its.Count > 0 Then Set FindIn = its(1)
        End If
        Err.Clear
    End If
    On Error GoTo 0
    If Not FindIn Is Nothing Or Not deep Then Exit Function
    For Each sf In fld.Folders
        Set hit = FindIn(sf, flt, True)
        If Not hit Is Nothing Then
            Set FindIn = hit
            Exit Function
        End If
    Next
End Function

Private Sub DropPullSheet()
    If Not SheetExists(PULL_SHEET) Then Exit Sub
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(PULL_SHEET).Delete
    Application.DisplayAlerts = True
End Sub

Private Function WantKey(ByVal k As Long) As Long
    ' The day and minute row k should pop up, as minutes; 0 when it shouldn't
    Dim tm As Long, due As Long
    If Trim$(CellText(mL(cOff, k))) <> "" Or Trim$(CellText(mL(cWhat, k))) = "" Then Exit Function
    tm = TimeOf(mL(cTime, k))
    If tm = NO_TIME Then Exit Function
    due = CLng(Int(NumOf(mL(cDue, k))))
    If due > 0 Then WantKey = due * 1440 + tm
End Function

Private Function MakeTask(ByVal fld As Object, ByVal what As String, ByVal want As Long) As String
    ' A popup for one thing; returns its ID
    Dim t As Object
    Set t = fld.Items.Add
    t.Subject = TASK_PREFIX & what
    t.DueDate = CDate(want \ 1440)
    t.ReminderSet = True
    t.ReminderTime = CDate(want / 1440#)
    t.Mileage = TASK_MARK
    t.Save
    MakeTask = t.EntryID
End Function

Private Sub DropTask(ByVal ns As Object, ByVal id As String)
    ' Deletes one of the tracker's popups, if it's still there and still the tracker's
    Dim t As Object
    If id = "" Or id = GONE_MARK Then Exit Sub
    Set t = ItemById(ns, id)
    If t Is Nothing Then Exit Sub
    If t.Mileage <> TASK_MARK Then Exit Sub
    If t.Parent.EntryID = ns.GetDefaultFolder(3).EntryID Then Exit Sub      ' already in Deleted Items
    t.Delete
End Sub

Private Function PushOutlook(ByVal force As Boolean) As Boolean
    ' Makes the popups match the rows: one for each open row with a set time still to come, made
    ' again when its day or time changes and deleted when it comes off the list, plus the next
    ' morning list. Returns True when a row's Outlook task changed.
    Dim ol As Object, ns As Object, fld As Object, k As Long, link As String, want As Long
    Dim nowKey As Long, nl As String, g As Variant
    Set ol = OutlookNow()
    If ol Is Nothing Then Exit Function
    On Error GoTo Trouble
    Set ns = ol.GetNamespace("MAPI")
    Set fld = TaskFolder(ns)
    nowKey = CLng(Int(NowT() * 1440 + 0.000001))
    If Not mGone Is Nothing Then
        For Each g In mGone
            DropTask ns, LinkId(CStr(g))
        Next
        Set mGone = New Collection
    End If
    For k = 1 To mN
        link = CellText(mL(cTask, k))
        want = WantKey(k)
        If link = "" And want <= nowKey Then
            ' nothing to pop, or its time has gone
        ElseIf link <> "" And LinkKey(link) = want Then
            ' as it should be
        Else
            DropTask ns, LinkId(link)
            nl = ""
            If want > nowKey Then nl = MakeTask(fld, CellText(mL(cWhat, k)), want) & "|" & want
            If nl = "" Then mL(cTask, k) = Empty Else mL(cTask, k) = nl
            PushOutlook = True
        End If
    Next
    PushMorning fld, force
    Exit Function
Trouble:
    mOlNote = "(Outlook didn't take the popup: " & Err.Description & ")"
    PushOutlook = True
End Function

Private Function MorningDay() As Long
    ' The workday the next morning list is for: today until 8:30, then the next workday
    If Not IsOff(mToday) And mMin < MORNING_AT Then MorningDay = mToday Else MorningDay = NextWorkday(mToday)
End Function

Private Function MorningList(ByVal d As Long, ByRef txt As String) As String
    ' The morning list for day d: its subject, "Tracker: 4 for today, 1 late", and in txt the
    ' names, late ones first. "" when there's nothing for it.
    Dim i As Long, nL As Long, nT As Long, nX As Long, lateRows As Variant, todayRows As Variant, timedRows As Variant
    Dim t As String
    txt = ""
    SortDay d, True, lateRows, nL, todayRows, nT, timedRows, nX      ' timed ones pop up on their own
    If nL + nT = 0 Then Exit Function
    For i = 1 To nL
        txt = txt & IIf(txt = "", "", vbCrLf) & CellText(mL(cWhat, lateRows(i))) & " (late)"
    Next
    For i = 1 To nT
        txt = txt & IIf(txt = "", "", vbCrLf) & CellText(mL(cWhat, todayRows(i)))
    Next
    If nT > 0 Then t = nT & " for today"
    If nL > 0 Then t = t & IIf(t = "", "", ", ") & nL & " late"
    MorningList = TASK_PREFIX & t
End Function

Private Sub PushMorning(ByVal fld As Object, ByVal force As Boolean)
    ' Keeps one morning list at 8:30 AM on the next workday. Today's, once it has come up, stays
    ' until it's dismissed; older ones go.
    Dim d As Long, subj As String, txt As String, key As String, its As Object, t As Object
    Dim i As Long, td As Long, found As Boolean
    d = MorningDay()
    subj = MorningList(d, txt)
    key = d & "|" & Hash(subj & vbLf & txt)
    If Not force And key = mMornKey Then Exit Sub
    Set its = fld.Items.Restrict("[Mileage] = '" & MORNING_MARK & "'")
    For i = its.Count To 1 Step -1
        Set t = its(i)
        td = CLng(Int(CDbl(t.DueDate)))
        If t.Complete Or Not t.ReminderSet Or td < mToday Or (td <> d And td <> mToday) Or (td = d And (found Or subj = "")) Then
            t.Delete
        ElseIf td = d Then
            found = True
            If t.BillingInformation <> key Then FillMorning t, d, subj, txt, key
        End If
    Next
    If Not found And subj <> "" Then FillMorning fld.Items.Add, d, subj, txt, key
    mMornKey = key
End Sub

Private Sub FillMorning(ByVal t As Object, ByVal d As Long, ByVal subj As String, ByVal txt As String, ByVal key As String)
    t.Subject = subj
    t.Body = txt
    t.DueDate = CDate(d)
    t.ReminderSet = True
    t.ReminderTime = CDate(d + MORNING_AT / 1440#)
    t.Mileage = MORNING_MARK
    t.BillingInformation = key
    t.Save
End Sub

Private Function Hash(ByVal s As String) As String
    ' A short stand-in for the morning list's words, to tell when they change
    Dim h As Double, i As Long
    For i = 1 To Len(s)
        h = h * 31 + AscW(Mid$(s, i, 1))
        h = h - Int(h / 2147483647#) * 2147483647#
    Next
    Hash = Len(s) & "." & CStr(h)
End Function

Private Function ReadPopupsBack(ByVal ns As Object, ByVal fld As Object, ByRef said As String, ByRef lastDone As String) As Boolean
    ' Reads back the tracker's popups. One marked complete in Outlook is Done here (a repeat moves
    ' on); one dismissed, or deleted in Outlook, goes and isn't made again until its date or time
    ' changes. Returns True when a row changed.
    Dim k As Long, link As String, id As String, t As Object, fid As String
    fid = fld.EntryID
    For k = 1 To mN
        link = CellText(mL(cTask, k))
        id = LinkId(link)
        If id <> "" And id <> GONE_MARK And Trim$(CellText(mL(cOff, k))) = "" Then
            Set t = ItemById(ns, id)
            If t Is Nothing Then
                mL(cTask, k) = GONE_MARK & "|" & LinkKey(link)
                ReadPopupsBack = True
            ElseIf t.Complete Then
                lastDone = CellText(mL(cWhat, k))
                said = said & IIf(said = "", "", " ") & "Marked complete in Outlook, and the tracker took it from there. " & DoneRow(k)
                If t.Parent.EntryID = fid Then t.Delete
                mL(cTask, k) = Empty
                ReadPopupsBack = True
            ElseIf Not t.ReminderSet Or t.Parent.EntryID <> fid Then
                If t.Parent.EntryID = fid Then t.Delete
                mL(cTask, k) = GONE_MARK & "|" & LinkKey(link)
                ReadPopupsBack = True
            End If
        End If
    Next
End Function

Private Sub SweepTasks(ByVal fld As Object)
    ' A popup no row points to any more (a row deleted by hand, or a file closed without saving) goes
    Dim linked As Object, k As Long, its As Object, i As Long, id As String
    Set linked = CreateObject("Scripting.Dictionary")
    For k = 1 To mN
        id = LinkId(CellText(mL(cTask, k)))
        If id <> "" Then linked(id) = True
    Next
    Set its = fld.Items.Restrict("[Mileage] = '" & TASK_MARK & "'")
    For i = its.Count To 1 Step -1
        If Not linked.Exists(its(i).EntryID) Then its(i).Delete
    Next
End Sub

Private Sub RunTick()
    ' What the timer does each minute, and opening the file does at once: takes what was marked
    ' complete or dismissed in Outlook, puts the popups right, and keeps the next morning list armed
    Dim ol As Object, ns As Object, fld As Object, said As String, lastDone As String, changed As Boolean
    Dim linksMoved As Boolean, flagged As Long
    On Error GoTo Oops
    If FindTable(LIST_TABLE) Is Nothing Then Exit Sub
    If Not HasColumns(FindTable(LIST_TABLE), LIST_COLS) Then Exit Sub
    If Not Ready() Then Exit Sub
    Set ol = OutlookNow()
    If Not ol Is Nothing Then
        Set ns = ol.GetNamespace("MAPI")
        Set fld = TaskFolder(ns)
        changed = ReadPopupsBack(ns, fld, said, lastDone)
        flagged = TakeFlags(ns, said)
        If flagged > 0 Then changed = True
    End If
    linksMoved = PushOutlook(True)
    If changed Or mDrawn <> mToday Then
        Finish
    ElseIf linksMoved Then
        SaveList
    End If
    If Not fld Is Nothing Then SweepTasks fld
    If lastDone <> "" Then
        ForgetUndo "Nothing to undo here: """ & lastDone & """ was marked complete in Outlook, and Undo can't take " & _
            "back what's done in Outlook."
    End If
    If flagged > 0 Then ForgetUndo NO_UNDO_FLAG
    If lastDone <> "" Or flagged > 0 Then Say said
    Exit Sub
Oops:
    Application.ScreenUpdating = True
End Sub

Private Sub StartTimer()
    Dim secs As Long
    If mTesting And mTickSecs = 0 Then Exit Sub         ' a test run sets its own
    If mTickSecs > 0 Then secs = mTickSecs Else secs = TICK_SECS
    StopTimer
    mTickAt = Now + secs / 86400#
    Application.OnTime mTickAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!TrackerTick"
End Sub

Private Sub StopTimer()
    If mTickAt = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime mTickAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!TrackerTick", , False
    mTickAt = 0
End Sub

Private Sub SwitchTasks(ByVal toMain As Boolean)
    ' Moves every popup to the other folder: deleted from the one it's in, made again in the new one
    Dim ns As Object, fld As Object, k As Long, its As Object, i As Long
    On Error GoTo Oops
    If Not Ready() Then Exit Sub
    If Not OutlookOpen("", ns) Then Exit Sub
    Set fld = TaskFolder(ns)
    For k = 1 To mN
        DropTask ns, LinkId(CellText(mL(cTask, k)))
        mL(cTask, k) = Empty
    Next
    SweepTasks fld
    Set its = fld.Items.Restrict("[Mileage] = '" & MORNING_MARK & "'")
    For i = its.Count To 1 Step -1
        its(i).Delete
    Next
    ThisWorkbook.Names.Add Name:=FOLDER_SETTING, RefersTo:="=""" & IIf(toMain, "main", TRACKER_FOLDER) & """", Visible:=False
    mMornKey = ""
    Finish
    If toMain Then
        Say "Popups now come from the main Tasks folder in Outlook. TrackerUseTrackerFolder puts them back in the Tracker folder."
    Else
        Say "Popups now come from the Tracker folder under Tasks in Outlook."
    End If
    Exit Sub
Oops:
    Application.ScreenUpdating = True
    Say "Moving the popups stopped: " & Err.Description
End Sub


' ---- Today ------------------------------------------------------------------------------

Private Sub DrawToday()
    ' Late first and never folded, then today's untimed things with repeats on top, then today's
    ' timed ones, which pop up on their own. 5 at most under Now and 20 under Coming up; the rest
    ' fold away in a row group.
    Dim ws As Worksheet, i As Long, last As Long, due As Long, tm As Long
    Dim late As Variant, nLate As Long, untimed As Variant, nUn As Long, timed As Variant, nTi As Long
    Dim soon As Variant, nSoon As Long, waits As Variant, nWait As Long, nLater As Long
    Dim sk As Variant, wk As Variant
    Dim rest As Variant, nRest As Long, room As Long, out As Variant, n As Long
    Dim heads As Collection, lates As Collection, folds As Collection, notes As Object, f As Variant
    Dim nowShown As Long, soonShown As Long, band As Long
    Set ws = ThisWorkbook.Worksheets(TODAY_SHEET)
    mDrawn = mToday
    SortDay mToday, False, late, nLate, untimed, nUn, timed, nTi     ' the same counts the morning list has
    ReDim soon(1 To mN + 1): ReDim waits(1 To mN + 1)
    ReDim sk(1 To mN + 1): ReDim wk(1 To mN + 1)
    For i = 1 To mN
        If Trim$(CellText(mL(cOff, i))) = "" And Trim$(CellText(mL(cWhat, i))) <> "" Then
            due = CLng(Int(NumOf(mL(cDue, i))))
            tm = TimeOf(mL(cTime, i))
            If due <= mToday Then
                ' late or today's, sorted above
            ElseIf Trim$(CellText(mL(cWait, i))) <> "" Then
                nWait = nWait + 1: waits(nWait) = i: wk(nWait) = WhenKey(i, due, tm)
            ElseIf due <= mToday + SOON_DAYS Then
                nSoon = nSoon + 1: soon(nSoon) = i: sk(nSoon) = WhenKey(i, due, tm)
            Else
                nLater = nLater + 1
            End If
        End If
    Next
    If nSoon > 1 Then SortBy sk, soon, 1, nSoon
    If nWait > 1 Then SortBy wk, waits, 1, nWait
    ReDim rest(1 To nUn + nTi + 1)
    For i = 1 To nUn: nRest = nRest + 1: rest(nRest) = untimed(i): Next
    For i = 1 To nTi: nRest = nRest + 1: rest(nRest) = timed(i): Next
    room = NOW_CAP - nLate
    If room < 0 Then room = 0
    If room > nRest Then room = nRest
    nowShown = nLate + room
    soonShown = nSoon
    If soonShown > SOON_CAP Then soonShown = SOON_CAP

    ws.Cells(COUNT_ROW, 1).Value = Split(DOW_LONG(), "|")(Dow(mToday)) & " " & FmtShort(mToday) & ": " & _
        (nLate + nRest) & " now, " & nSoon & " coming up, " & nWait & " waiting, " & nLater & " later."

    Set notes = LatestNotes()
    Set heads = New Collection: Set lates = New Collection: Set folds = New Collection
    ReDim out(1 To nLate + nRest + nSoon + nWait + 12, 1 To 5)
    n = 0
    PutHead out, n, heads, "Now, late and today: " & IIf(nRest > room, "showing " & nowShown & " of " & (nLate + nRest), CStr(nLate + nRest))
    If nLate + nRest = 0 Then PutText out, n, "Nothing for today."
    For i = 1 To nLate
        PutItem out, n, late(i), notes
        lates.Add n
    Next
    For i = 1 To room: PutItem out, n, rest(i), notes: Next
    If nRest > room Then
        For i = room + 1 To nRest: PutItem out, n, rest(i), notes: Next
        folds.Add Array(n - (nRest - room) + 1, n)
        PutText out, n, (nRest - room) & " more for today"
    End If
    PutHead out, n, heads, "Coming up, next " & SOON_DAYS & " days: " & IIf(nSoon > soonShown, "showing " & soonShown & " of " & nSoon, CStr(nSoon))
    If nSoon = 0 Then PutText out, n, "Nothing in the next " & SOON_DAYS & " days."
    For i = 1 To nSoon: PutItem out, n, soon(i), notes: Next
    If nSoon > soonShown Then
        folds.Add Array(n - (nSoon - soonShown) + 1, n)
        PutText out, n, (nSoon - soonShown) & " more coming up"
    End If
    PutHead out, n, heads, "Waiting on them: " & nWait
    If nWait = 0 Then PutText out, n, "Nobody owes you anything right now."
    For i = 1 To nWait: PutItem out, n, waits(i), notes: Next
    PutHead out, n, heads, "Later: " & nLater & " more, all on the List sheet"

    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If last >= FIRST_ROW Then ws.Rows(FIRST_ROW & ":" & last).Delete
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + n - 1, 5)).Value = out
    ws.Rows(FIRST_ROW & ":" & (FIRST_ROW + n - 1)).RowHeight = TODAY_ROW_PT
    For Each f In heads
        With ws.Range(ws.Cells(FIRST_ROW + f - 1, 1), ws.Cells(FIRST_ROW + f - 1, 5))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
    Next
    ' every other thing under a heading in light gray, so a row is easy to follow across
    For i = 1 To n
        If VarType(out(i, 1)) = vbString Then
            band = 0
        ElseIf Not IsEmpty(out(i, 1)) Then
            band = band + 1
            If band Mod 2 = 0 Then ws.Range(ws.Cells(FIRST_ROW + i - 1, 1), ws.Cells(FIRST_ROW + i - 1, 5)).Interior.Color = TODAY_STRIPE
        End If
    Next
    For Each f In lates
        With ws.Cells(FIRST_ROW + f - 1, 2).Font
            .Color = RGB(192, 0, 0)
            .Bold = True
        End With
    Next
    ws.Outline.SummaryRow = xlSummaryBelow
    For Each f In folds
        ws.Rows((FIRST_ROW + f(0) - 1) & ":" & (FIRST_ROW + f(1) - 1)).Group
    Next
    If folds.Count > 0 Then ws.Outline.ShowLevels RowLevels:=1
End Sub

Private Sub SortDay(ByVal d As Long, ByVal asAdded As Boolean, ByRef late As Variant, ByRef nLate As Long, _
                    ByRef untimed As Variant, ByRef nUn As Long, ByRef timed As Variant, ByRef nTi As Long)
    ' The open rows for day d, the way Today and the morning list count them: late (due before d),
    ' due d with no time, and due d with a time. Sorted the way Today shows them (late by date,
    ' repeats first, timed by time), or with asAdded in the order they were added.
    Dim i As Long, due As Long, tm As Long, lk As Variant, uk As Variant, tk As Variant
    ReDim late(1 To mN + 1): ReDim untimed(1 To mN + 1): ReDim timed(1 To mN + 1)
    ReDim lk(1 To mN + 1): ReDim uk(1 To mN + 1): ReDim tk(1 To mN + 1)
    nLate = 0: nUn = 0: nTi = 0
    For i = 1 To mN
        If Trim$(CellText(mL(cOff, i))) = "" And Trim$(CellText(mL(cWhat, i))) <> "" Then
            due = CLng(Int(NumOf(mL(cDue, i))))
            tm = TimeOf(mL(cTime, i))
            If due < d Then
                nLate = nLate + 1: late(nLate) = i
                If asAdded Then lk(nLate) = NumOf(mL(cId, i)) Else lk(nLate) = WhenKey(i, due, tm)
            ElseIf due = d And tm = NO_TIME Then
                nUn = nUn + 1: untimed(nUn) = i
                uk(nUn) = IIf(asAdded Or Trim$(CellText(mL(cRule, i))) = "", 1, 0) * 1000000# + NumOf(mL(cId, i))
            ElseIf due = d Then
                nTi = nTi + 1: timed(nTi) = i: tk(nTi) = tm * 1000000# + NumOf(mL(cId, i))
            End If
        End If
    Next
    If nLate > 1 Then SortBy lk, late, 1, nLate
    If nUn > 1 Then SortBy uk, untimed, 1, nUn
    If nTi > 1 Then SortBy tk, timed, 1, nTi
End Sub

Private Function WhenKey(ByVal i As Long, ByVal due As Long, ByVal tm As Long) As Double
    ' Sorts by date, then time (none goes last), then number
    If tm = NO_TIME Then tm = 9999
    WhenKey = (CDbl(due) * 10000# + tm) * 100000# + NumOf(mL(cId, i))
End Function

Private Sub PutHead(ByRef out As Variant, ByRef n As Long, ByVal heads As Collection, ByVal t As String)
    n = n + 1
    out(n, 1) = t
    heads.Add n
End Sub

Private Sub PutText(ByRef out As Variant, ByRef n As Long, ByVal t As String)
    n = n + 1
    out(n, 2) = t
End Sub

Private Sub PutItem(ByRef out As Variant, ByRef n As Long, ByVal i As Long, ByVal notes As Object)
    ' One row on Today: number, when, what, repeats, latest note
    Dim due As Long, tm As Long, t As String, who As String, what As String
    n = n + 1
    due = CLng(Int(NumOf(mL(cDue, i))))
    tm = TimeOf(mL(cTime, i))
    who = Trim$(CellText(mL(cWait, i)))
    If tm <> NO_TIME Then t = " " & FmtTime(tm)
    If due < mToday Then
        out(n, 2) = "Late, " & FmtDay(due) & t
    ElseIf due = mToday Then
        If who <> "" Then out(n, 2) = "Check back today" Else out(n, 2) = "Today" & t
    Else
        out(n, 2) = IIf(who <> "", "Check ", "") & FmtDay(due) & t
    End If
    out(n, 1) = mL(cId, i)
    what = CellText(mL(cWhat, i))
    If who <> "" And InStr(1, what, "on " & who, vbTextCompare) = 0 Then what = what & " (waiting on " & who & ")"
    out(n, 3) = what
    out(n, 4) = CellText(mL(cRep, i))
    If notes.Exists(CLng(NumOf(mL(cId, i)))) Then out(n, 5) = notes(CLng(NumOf(mL(cId, i))))
End Sub

Private Function LatestNotes() As Object
    ' Each row's newest note from the Log
    Dim d As Object, v As Variant, i As Long, cI As Long, cH As Long, cK As Long, id As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set LatestNotes = d
    If mLogT.ListRows.Count = 0 Then Exit Function
    cI = ColOf(mLogT, "#"): cH = ColOf(mLogT, "What happened"): cK = ColOf(mLogT, "Kind")
    v = mLogT.DataBodyRange.Value2
    For i = 1 To UBound(v, 1)
        If CellText(v(i, cK)) = "note" Then
            id = CLng(NumOf(v(i, cI)))
            If Not d.Exists(id) Then d.Add id, CellText(v(i, cH))
        End If
    Next
End Function


' ---- setup ------------------------------------------------------------------------

Private Function EnsureSheet(ByVal nm As String, ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet
    If SheetExists(nm) Then
        Set EnsureSheet = ThisWorkbook.Worksheets(nm)
        Exit Function
    End If
    If after Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=after)
    End If
    ws.Name = nm
    Set EnsureSheet = ws
End Function

Private Sub DropBlankSheets()
    ' A new workbook's empty Sheet1 goes, so the tracker's sheets are all there is
    Dim ws As Worksheet, ours As String
    ours = "|" & TODAY_SHEET & "|" & LIST_SHEET & "|" & LOG_SHEET & "|" & OFF_SHEET & "|" & PASTE_SHEET & "|" & PULL_SHEET & "|"
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ours, "|" & ws.Name & "|") = 0 Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next
End Sub

Private Function EnsureTable(ByVal ws As Worksheet, ByVal nm As String, ByVal cols As String, ByVal fmts As String, _
                             ByVal widths As String, ByVal note As String) As ListObject
    ' The table, made when it's missing; a missing column gets added at the end. Rows stay.
    Dim lo As ListObject, c As Variant, names As Variant, f As Variant, w As Variant, k As Long
    names = Split(cols, "|"): f = Split(fmts, "|"): w = Split(widths, "|")
    Set lo = FindTable(nm)
    If lo Is Nothing Then
        For k = 0 To UBound(names)
            ws.Columns(k + 1).NumberFormat = f(k)
            ws.Columns(k + 1).ColumnWidth = CDbl(w(k))
            ws.Cells(TABLE_TOP, k + 1).Value = names(k)
        Next
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(TABLE_TOP, 1), ws.Cells(TABLE_TOP + 1, UBound(names) + 1)), , xlYes)
        lo.Name = nm
        lo.TableStyle = "TableStyleLight9"
        If lo.ListRows.Count > 0 Then lo.DataBodyRange.Delete
    Else
        For k = 0 To UBound(names)
            If ColOf(lo, CStr(names(k))) = 0 Then
                lo.ListColumns.Add.Name = names(k)
                With lo.ListColumns(lo.ListColumns.Count).Range.EntireColumn
                    .NumberFormat = f(k)
                    .ColumnWidth = CDbl(w(k))
                End With
            End If
        Next
    End If
    ws.Cells(1, 1).Value = note
    ws.Cells(1, 1).Font.Color = RGB(89, 89, 89)
    ws.Cells(1, 1).Font.Italic = True
    Set EnsureTable = lo
End Function

Private Sub HideColumns(ByVal lo As ListObject, ByVal cols As String)
    Dim c As Variant
    For Each c In Split(cols, "|")
        lo.ListColumns(CStr(c)).Range.EntireColumn.Hidden = True
    Next
End Sub

Private Sub CrossOutDone(ByVal ws As Worksheet)
    ' A List row that's off the list shows crossed out and gray
    Dim rng As Range, fc As FormatCondition, col As String
    col = Split(mList.ListColumns("Off the list").Range.Cells(1, 1).Address(True, False), "$")(0)
    Set rng = ws.Range(ws.Cells(TABLE_TOP + 1, mList.Range.Column), ws.Cells(CF_ROWS, mList.Range.Column + mList.ListColumns.Count - 1))
    ws.Cells.FormatConditions.Delete
    ws.Activate                                     ' Excel reads the rule from the selected cell
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=LEN($" & col & (TABLE_TOP + 1) & ")>0")
    fc.Font.Strikethrough = True
    fc.Font.Color = RGB(128, 128, 128)
End Sub

Private Sub DressToday(ByVal ws As Worksheet)
    ' Today's buttons, how-to, count line and column headers
    Dim names As Variant, f As Variant, w As Variant, k As Long, b As Object, tb As Object
    names = Split(TODAY_COLS, "|"): f = Split(TODAY_FORMATS, "|"): w = Split(TODAY_WIDTHS, "|")
    For k = 0 To UBound(names)
        ws.Columns(k + 1).NumberFormat = f(k)
        ws.Columns(k + 1).ColumnWidth = CDbl(w(k))
        ws.Cells(HEAD_ROW, k + 1).Value = names(k)
    Next
    With ws.Range(ws.Cells(HEAD_ROW, 1), ws.Cells(HEAD_ROW, UBound(names) + 1))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
    End With
    ws.Range(ws.Columns(1), ws.Columns(UBound(names) + 1)).Font.Size = TODAY_PT
    ws.Columns(UBound(names) + 2).ColumnWidth = HOWTO_WIDTH
    ws.Rows(COUNT_ROW).RowHeight = 27
    ws.Rows(2).RowHeight = 39
    ws.Rows(3).RowHeight = 30
    ws.Rows(SAID_ROW).RowHeight = 48
    ws.Rows(HEAD_ROW).RowHeight = TODAY_ROW_PT
    With ws.Cells(COUNT_ROW, 1)
        .Font.Bold = True
        .Font.Size = TODAY_PT + 1
        .VerticalAlignment = xlCenter
    End With
    ' an earlier copy kept the how-to at the top right; what a button says now runs the width of the list
    ws.Range(ws.Cells(1, 1), ws.Cells(SAID_ROW, 6)).UnMerge
    ws.Range(ws.Cells(1, 5), ws.Cells(SAID_ROW, 6)).Clear
    ws.Range(ws.Cells(SAID_ROW, 1), ws.Cells(SAID_ROW, 5)).Merge
    With ws.Cells(SAID_ROW, 1)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Italic = True
        .Font.Color = RGB(31, 78, 121)
    End With
    ' the how-to: a text box right of the list, under the rows that stay put, so a fold never cuts it
    On Error Resume Next
    ws.Shapes("trkHowTo").Delete
    On Error GoTo 0
    Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, ws.Columns(6).Left + 6, ws.Rows(FIRST_ROW).Top + 4, _
        ws.Columns(6).Width - 12, 300)
    tb.Name = "trkHowTo"
    tb.Placement = xlFreeFloating
    tb.Line.Visible = msoFalse
    tb.Fill.Visible = msoFalse
    With tb.TextFrame2
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeShapeToFitText
        .TextRange.Text = "Add: press Add, or Ctrl+Shift+A, and type one line, like call Huntsville back 11/17 9am. " & _
            "No date means today. What it understood shows above the list." & vbLf & _
            "Repeats: every monday, monthly on the 5th, quarterly, renew bond 1/15 yearly, until 12/31." & vbLf & _
            "Click a row, then: Done takes it off Today (a repeat moves to its next date). Move to tomorrow puts it " & _
            "on the next workday. Change date takes one line, like fri 2pm or 11/20. Note adds a dated line to the Log. " & _
            "Drop it takes it off the list, or stops a repeat. Open email opens the email it came from. " & _
            "List keeps them all, crossed out." & vbLf & _
            "Undo takes back your last change." & vbLf & _
            "Popups come from Outlook: a thing with a time pops up then, everything else is in the 8:30 AM list. " & _
            "Mark Complete on a popup marks it done here." & vbLf & _
            "Flag an email in Outlook and it lands here within a minute, due on the flag's day. A flagged sent " & _
            "email lands under Waiting on them." & vbLf & _
            "Pull from Outlook lists your calendar and the emails you'd flagged before, Yes or No each. Change any, " & _
            "then press Take the Yes ones. Take them off my calendar then moves those to Deleted Items." & vbLf & _
            "Paste a list: paste one thing per line, then press Take these." & vbLf & _
            "A line like ""8 more for today"" folds the rest away. Press the + at the far left to open it."
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(64, 64, 64)
        .TextRange.ParagraphFormat.SpaceAfter = 6
    End With
    For Each b In ws.Buttons
        If Left$(b.Name, 3) = "trk" Then b.Delete
    Next
    ' the top row is what you press all day, big; the rest go under it
    AddButton ws, "trkAdd", "Add", "AddThing", 0, 0, 80, True
    AddButton ws, "trkDone", "Done", "DoneThing", 0, 86, 80, True
    AddButton ws, "trkMove", "Move to tomorrow", "MoveTomorrow", 0, 172, 140, True
    AddButton ws, "trkChange", "Change date", "ChangeDate", 0, 318, 116, True
    AddButton ws, "trkNote", "Note", "NoteThing", 1, 0, 50
    AddButton ws, "trkDrop", "Drop it", "DropIt", 1, 54, 60
    AddButton ws, "trkOpen", "Open email", "OpenEmail", 1, 118, 80
    AddButton ws, "trkUndo", "Undo", "UndoLast", 1, 202, 50
    AddButton ws, "trkPaste", "Paste a list", "PasteList", 1, 272, 80
    AddButton ws, "trkPull", "Pull from Outlook", "PullFromOutlook", 1, 356, 110
    AddButton ws, "trkOffCal", "Take them off my calendar", "TakeOffCalendar", 1, 470, 150
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 0
        .SplitRow = HEAD_ROW
        .FreezePanes = True
        .DisplayGridlines = False
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal key As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal rowSlot As Long, ByVal x As Double, ByVal w As Double, _
                      Optional ByVal big As Boolean = False)
    ' Buttons sit in rows 2 and 3 on Today, row 1 elsewhere, x points in from the left and w wide.
    ' A big one is taller, in bold.
    Dim b As Object, r As Range, h As Double
    On Error Resume Next
    ws.Buttons(key).Delete
    On Error GoTo 0
    If ws.Name = TODAY_SHEET Then Set r = ws.Rows(2 + rowSlot) Else Set r = ws.Rows(1)
    If big Then h = 30 Else h = 24
    Set b = ws.Buttons.Add(ws.Range("A1").Left + 4 + x, r.Top + (r.RowHeight - h) / 2, w, h)
    b.Name = key
    b.Caption = caption
    b.OnAction = macroName
    If big Then
        b.Font.Bold = True
        b.Font.Size = TODAY_PT
    End If
End Sub


' ---- finding things ------------------------------------------------------------------

Private Function FindTable(ByVal nm As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nm Then Set FindTable = lo: Exit Function
        Next
    Next
End Function

Private Function ColOf(ByVal lo As ListObject, ByVal nm As String) As Long
    Dim k As Long
    For k = 1 To lo.ListColumns.Count
        If lo.ListColumns(k).Name = nm Then ColOf = k: Exit Function
    Next
End Function

Private Function HasColumns(ByVal lo As ListObject, ByVal cols As String) As Boolean
    Dim c As Variant
    For Each c In Split(cols, "|")
        If ColOf(lo, CStr(c)) = 0 Then Exit Function
    Next
    HasColumns = True
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next
End Function


' ---- the boxes -------------------------------------------------------------

Private Function Ask(ByVal prompt As String, ByVal title As String) As Variant
    ' What was typed, or False if the box was cancelled
    If Not mFeed Is Nothing Then
        If mFeed.Count = 0 Then
            Ask = False
        Else
            Ask = mFeed(1)
            mFeed.Remove 1
            If Ask = "<cancel>" Then Ask = False
        End If
        Exit Function
    End If
    Ask = Application.InputBox(prompt, title, "", Type:=2)
End Function

Private Sub Say(ByVal msg As String)
    ' What a button says shows at the top of Today, with anything Outlook turned away
    If mOlNote <> "" Then
        msg = msg & " " & mOlNote
        mOlNote = ""
    End If
    If Not mFeed Is Nothing Then mSaid = mSaid & msg & vbLf
    If SheetExists(TODAY_SHEET) Then
        ThisWorkbook.Worksheets(TODAY_SHEET).Cells(SAID_ROW, 1).Value = msg
    ElseIf mFeed Is Nothing Then
        MsgBox msg, vbInformation, "Tracker"
    End If
End Sub

Public Function TrackerTest(ByVal macroName As String, ByVal answers As String) As String
    ' For test runs: feeds answers to the boxes (split on |) and returns what the buttons
    ' said. Not needed day to day.
    Dim a As Variant
    mTesting = True
    If macroName = "Said" Then                          ' what was said since Arm
        TrackerTest = mSaid
        mSaid = ""
        Set mFeed = Nothing
        Exit Function
    End If
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
        Case "AddThing": AddThing
        Case "DoneThing": DoneThing
        Case "MoveTomorrow": MoveTomorrow
        Case "ChangeDate": ChangeDate
        Case "NoteThing": NoteThing
        Case "DropIt": DropIt
        Case "UndoLast": UndoLast
        Case "ForgetUndo"                                ' as a change in Outlook does
            ForgetUndo answers
            Set mFeed = New Collection
        Case "PasteList": PasteList
        Case "TakeThese": TakeThese
        Case "CancelPaste": CancelPaste
        Case "Auto_Open": Auto_Open
        Case "Auto_Close": Auto_Close
        Case "Tick": RunTick                            ' what the timer does, once
        Case "TickEvery"                                 ' the timer's seconds; -1 is its own minute, 0 is off
            mTickSecs = CLng(answers)
            If mTickSecs = 0 Then StopTimer
            Set mFeed = New Collection
        Case "Timer"                                     ' on and the seconds to the next one, or off
            If mTickAt > 0 Then
                mSaid = "on " & CLng((mTickAt - Now) * 86400#) & " ticks=" & mTicks
            Else
                mSaid = "off ticks=" & mTicks
            End If
            Set mFeed = New Collection
        Case "OpenEmail": OpenEmail
        Case "PullFromOutlook": PullFromOutlook
        Case "TakeYes": TakeYes
        Case "CancelPull": CancelPull
        Case "TakeOffCalendar": TakeOffCalendar
        Case "TrackerUseMainTasks": TrackerUseMainTasks
        Case "TrackerUseTrackerFolder": TrackerUseTrackerFolder
        Case "Clock"
            mClock = CDbl(answers)
            Set mFeed = New Collection
        Case "Arm"                                       ' the answers wait for a key press
            Exit Function
        Case Else: mSaid = "no macro called " & macroName & vbLf
    End Select
    If mFeed.Count > 0 Then mSaid = mSaid & "(unused answers: " & mFeed.Count & ")" & vbLf
    TrackerTest = mSaid
    Set mFeed = Nothing
End Function
