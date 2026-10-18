Option Explicit

'==================================================================
'  SUPPLIER LICENSE CHECK  (Texas)
'
'  Moves every line whose seller is not licensed in Texas as a
'  supplier of that fuel (licensed only as a distributor, or not
'  on the state's list at all) below the lines to keep, so they
'  can be checked and deleted.
'
'  First get the state's list: on comptroller.texas.gov go to
'  Taxes, Motor Fuels IFTA and Interstate Truckers, then Diesel
'  Fuel (or Gasoline). Under Additional Resources click Licensed
'  Fuels Taxpayer Search, then the Download button. It saves
'  permitted-fuels-taxpayers.csv. The state updates it nightly.
'
'  Run CheckSupplierLicenses with the sheet active. It will ask for:
'     1. any cell in the column with the seller's taxpayer number
'     2. the first row of real data
'     3. D for diesel or G for gasoline
'     4. the state's list you downloaded
'
'  On a file that also goes through the credit / rebill matcher,
'  run this one first and delete what it moves down.
'==================================================================

'--- Settings -----------------------------------------------------
' Leave True and the macro copies the sheet and works on the copy,
' so the original is untouched (Ctrl+Z can't undo a macro, so the
' copy is your undo). Set False to work on the sheet itself.
Private Const WORK_ON_A_COPY As Boolean = True

' The licenses that keep a line, split by |. A permissive supplier
' comes out like a distributor.
Private Const DIESEL_KEEPS As String = "DIESEL FUEL SUPPLIER"
Private Const GAS_KEEPS As String = "GASOLINE SUPPLIER"

' Say so when the state's list was saved more than this many days ago.
Private Const STALE_DAYS As Long = 7
'------------------------------------------------------------------

Private Const BOX_TITLE As String = "Supplier license check"
Private mIcon As Long


Public Sub CheckSupplierLicenses()

    Dim src As Worksheet
    Dim rng As Range
    Dim v As Variant, listPath As Variant
    Dim idCol As Long, startRow As Long
    Dim fuel As String, msg As String

    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set src = ActiveSheet

    '--- prompts --------------------------------------------------
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Click ANY cell inside the column with the seller's Texas taxpayer number:", _
        Title:="Select taxpayer number column", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    idCol = rng.Column

    v = Application.InputBox( _
        Prompt:="Enter the first row number containing actual data:", _
        Title:="Select start row", Default:=2, Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub      ' Cancel
    startRow = CLng(v)
    If startRow < 1 Then Exit Sub

    Do
        v = Application.InputBox( _
            Prompt:="Is this the diesel file or the gasoline file?" & vbCrLf & vbCrLf & _
                    "Type D for diesel or G for gasoline:", _
            Title:="Select fuel", Type:=2)
        If VarType(v) = vbBoolean Then Exit Sub  ' Cancel
        fuel = UCase$(Left$(Trim$(CStr(v)), 1))
    Loop Until fuel = "D" Or fuel = "G"

    listPath = PickList(src.Parent.Path)
    If VarType(listPath) = vbBoolean Then Exit Sub

    msg = RunCheck(src, idCol, startRow, fuel, CStr(listPath))
    MsgBox msg, mIcon, BOX_TITLE
End Sub


'==================================================================
'  The check itself. Returns the message to show.
'==================================================================
Public Function RunCheck(ByVal src As Worksheet, ByVal idCol As Long, ByVal startRow As Long, _
                         ByVal fuel As String, ByVal listPath As String) As String

    '--- all declarations up front -------------------------------
    Dim ws As Worksheet
    Dim listWb As Workbook, wb As Workbook
    Dim lic As Object, stat As Object
    Dim raw As Variant, ids As Variant, keeps As Variant
    Dim notes() As Variant, keys() As Variant
    Dim cnt(0 To 3) As Long
    Dim priorCalc As XlCalculation
    Dim madeCopy As Boolean, wasOpen As Boolean, suspended As Boolean
    Dim colId As Long, colType As Long, colSusp As Long
    Dim lastRow As Long, lastCol As Long, noteCol As Long
    Dim h1 As Long, h2 As Long, listRows As Long
    Dim i As Long, j As Long, n As Long, r As Long, sec As Long
    Dim rowSec1 As Long, rowSec2 As Long, rowSec3 As Long
    Dim errNum As Long
    Dim saved As Date
    Dim k As String, ptype As String, fuelWord As String
    Dim msg As String, eDesc As String, noteLetter As String
    '--------------------------------------------------------------

    mIcon = vbInformation
    If fuel = "D" Then
        keeps = Split(DIESEL_KEEPS, "|")
        fuelWord = "diesel"
    Else
        keeps = Split(GAS_KEEPS, "|")
        fuelWord = "gasoline"
    End If

    '--- save Excel state, then arm the handler -------------------
    priorCalc = Application.Calculation
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- read the state's list ------------------------------------
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, listPath, vbTextCompare) = 0 Then
            Set listWb = wb
            wasOpen = True
        End If
    Next wb
    If listWb Is Nothing Then Set listWb = Workbooks.Open(Filename:=listPath, ReadOnly:=True)
    raw = listWb.Worksheets(1).UsedRange.Value
    If Not wasOpen Then listWb.Close SaveChanges:=False
    Set listWb = Nothing
    saved = FileDateTime(listPath)

    If Not IsArray(raw) Then GoTo NotTheList
    For j = 1 To UBound(raw, 2)
        Select Case UCase$(Trim$(CellText(raw(1, j))))
            Case "TAXPAYER ID", "TP_ID": colId = j
            Case "PERMIT TYPE", "PERMIT_TYPE", "LICENSE TYPE": colType = j
            Case "SUSPENDED": colSusp = j
        End Select
    Next j
    If colId = 0 Or colType = 0 Then GoTo NotTheList

    ' One entry per taxpayer number: every license it holds, and
    ' 2 = supplier for this fuel, 1 = that license is suspended, 0 = neither.
    Set lic = CreateObject("Scripting.Dictionary")
    Set stat = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(raw, 1)
        k = TaxpayerKey(raw(i, colId))
        ptype = UCase$(Trim$(CellText(raw(i, colType))))
        If Len(k) > 0 And Len(ptype) > 0 Then
            listRows = listRows + 1
            If lic.Exists(k) Then
                lic(k) = lic(k) & ", " & ptype
            Else
                lic.Add k, ptype
                stat.Add k, 0
            End If
            If IsInList(ptype, keeps) Then
                suspended = False
                If colSusp > 0 Then suspended = IsYes(raw(i, colSusp))
                If Not suspended Then
                    stat(k) = 2
                ElseIf stat(k) = 0 Then
                    stat(k) = 1
                End If
            End If
        End If
    Next i
    raw = Empty
    If listRows = 0 Then GoTo NotTheList

    '--- work on a copy so the original survives ------------------
    If WORK_ON_A_COPY Then
        src.Copy After:=src
        Set ws = src.Parent.Sheets(src.Index + 1)
        ws.Name = SafeSheetName(src.Parent, src.Name & " SUPPLIERS")
        madeCopy = True
    Else
        Set ws = src
    End If

    '--- trim trailing Grand Total / blank rows -------------------
    lastRow = LastUsedRow(ws)
    Do While lastRow >= startRow
        If IsTotalOrBlankRow(ws, lastRow, idCol) Then
            ws.Rows(lastRow).Delete
            lastRow = lastRow - 1
        Else
            Exit Do
        End If
    Loop

    If lastRow < startRow Then
        If madeCopy Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            src.Activate
        End If
        Application.Calculation = priorCalc
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        mIcon = vbExclamation
        RunCheck = "No data rows found at or below row " & startRow & "."
        Exit Function
    End If

    lastCol = LastUsedColumn(ws)
    If lastCol < idCol Then lastCol = idCol

    '--- sort each line into a section ----------------------------
    '   0 = keep, supplier          2 = number not on the state's list
    '   1 = licensed, not supplier  3 = no number, or supplier suspended
    n = lastRow - startRow + 1
    If n = 1 Then
        ReDim ids(1 To 1, 1 To 1)
        ids(1, 1) = ws.Cells(startRow, idCol).Value
    Else
        ids = ws.Range(ws.Cells(startRow, idCol), ws.Cells(lastRow, idCol)).Value
    End If

    ReDim notes(1 To n, 1 To 1)
    ReDim keys(1 To n, 1 To 2)

    For i = 1 To n
        k = TaxpayerKey(ids(i, 1))
        If Len(k) = 0 Then
            sec = 3: notes(i, 1) = "No taxpayer number"
        ElseIf Not lic.Exists(k) Then
            sec = 2: notes(i, 1) = "Not on the state's list"
        ElseIf stat(k) = 2 Then
            sec = 0: notes(i, 1) = Empty
        ElseIf stat(k) = 1 Then
            sec = 3: notes(i, 1) = "Supplier license suspended. Holds: " & lic(k)
        Else
            sec = 1: notes(i, 1) = "Holds: " & lic(k)
        End If
        cnt(sec) = cnt(sec) + 1
        keys(i, 1) = sec
        keys(i, 2) = i                                ' keeps original order stable
    Next i

    noteCol = lastCol + 1
    noteLetter = Split(ws.Cells(1, noteCol).Address(True, False), "$")(0)

    '--- reorder the sheet with ONE sort --------------------------
    ' Only when something has to move. The reason goes in the column
    ' just past the data; the lines kept leave it blank.
    If cnt(0) < n Then
        h1 = lastCol + 3: h2 = h1 + 1
        ws.Range(ws.Cells(startRow, noteCol), ws.Cells(lastRow, noteCol)).Value = notes
        ws.Range(ws.Cells(startRow, h1), ws.Cells(lastRow, h2)).Value = keys

        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(startRow, h1), ws.Cells(lastRow, h1)), Order:=xlAscending
            .SortFields.Add Key:=ws.Range(ws.Cells(startRow, h2), ws.Cells(lastRow, h2)), Order:=xlAscending
            .SetRange ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, h2))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
            .SortFields.Clear
        End With

        ws.Range(ws.Cells(1, h1), ws.Cells(1, h2)).EntireColumn.Delete

        '--- colour each section --------------------------------
        r = startRow + cnt(0)
        For sec = 1 To 3
            If cnt(sec) > 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r + cnt(sec) - 1, noteCol)).Interior.Color = SectionColour(sec)
                r = r + cnt(sec)
            End If
        Next sec

        '--- banners, inserted bottom-up so row numbers hold ------
        rowSec1 = startRow + cnt(0)
        rowSec2 = rowSec1 + cnt(1)
        rowSec3 = rowSec2 + cnt(2)

        If cnt(3) > 0 Then
            InsertBanner ws, rowSec3, "--- NO TAXPAYER NUMBER, OR THE SUPPLIER LICENSE IS SUSPENDED - CHECK THESE BY HAND ---"
        End If
        If cnt(2) > 0 Then
            InsertBanner ws, rowSec2, "--- NOT ON THE STATE'S LIST - CHECK THE NUMBER, THEN DELETE ---"
        End If
        If cnt(1) > 0 Then
            InsertBanner ws, rowSec1, "--- NOT LICENSED AS A " & UCase$(fuelWord) & " SUPPLIER - DELETE THESE ---"
        End If
    End If

    '--- the message ----------------------------------------------
    msg = n & " lines checked against the state's list (" & Format$(listRows, "#,##0") & _
          " licenses, saved " & Format$(saved, "mmm d, yyyy") & ")." & vbCrLf & vbCrLf & _
          cnt(0) & " lines kept at the top: licensed as a " & fuelWord & " supplier" & vbCrLf & _
          cnt(1) & " lines - RED: licensed, but not as a " & fuelWord & " supplier. Delete these." & vbCrLf & _
          cnt(2) & " lines - ORANGE: number not on the state's list. Check it, then delete."
    If cnt(3) > 0 Then
        msg = msg & vbCrLf & cnt(3) & " lines - GREY: no taxpayer number, or the supplier license is suspended. Check by hand."
    End If
    If cnt(0) < n Then
        msg = msg & vbCrLf & vbCrLf & "Why each line moved is in column " & noteLetter & "."
    End If
    If cnt(2) > 0 And cnt(0) + cnt(1) = 0 Then
        mIcon = vbExclamation
        msg = msg & vbCrLf & vbCrLf & "WARNING: not one number was on the state's list. " & _
              "Check you clicked the column with the seller's Texas taxpayer number."
    End If
    If Date - Int(saved) > STALE_DAYS Then
        msg = msg & vbCrLf & vbCrLf & "NOTE: the state's list was saved " & CLng(Date - Int(saved)) & _
              " days ago. Download a fresh one if this is for a new month."
    End If
    If madeCopy Then
        msg = msg & vbCrLf & vbCrLf & "Original sheet untouched; this is """ & ws.Name & """."
    End If

    ws.Activate
    ws.Cells(startRow, 1).Select
    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    RunCheck = msg
    Exit Function

NotTheList:
    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    mIcon = vbExclamation
    RunCheck = "That file doesn't look like the state's license list. It needs a Taxpayer ID " & _
               "column and a Permit Type column with licenses in it." & vbCrLf & vbCrLf & _
               "Download it again from the Licensed Fuels Taxpayer Search page (the Download button)."
    Exit Function

CleanFail:
    errNum = Err.Number
    eDesc = Err.Description
    Resume AfterFail
AfterFail:
    On Error Resume Next
    If Not listWb Is Nothing Then
        If Not wasOpen Then listWb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    Application.Calculation = priorCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    mIcon = vbCritical
    msg = "Stopped on error " & errNum & ": " & eDesc & vbCrLf & vbCrLf & _
          "Excel's calculation and screen settings have been put back. "
    If WORK_ON_A_COPY Then
        msg = msg & "Your original sheet was not touched."
    Else
        msg = msg & "The sheet was being edited in place - check it before continuing."
    End If
    RunCheck = msg
End Function


'==================================================================
'  Helpers
'==================================================================

' The list picked, or False if the box was cancelled. Starts in this
' workbook's folder, or in Downloads if the workbook isn't saved.
Private Function PickList(ByVal startFolder As String) As Variant
    Dim dl As String
    dl = Environ$("USERPROFILE") & "\Downloads"
    On Error Resume Next
    If Mid$(startFolder, 2, 1) = ":" Then
        ChDrive Left$(startFolder, 1)
        ChDir startFolder
    ElseIf Len(Dir$(dl, vbDirectory)) > 0 Then
        ChDrive Left$(dl, 1)
        ChDir dl
    End If
    On Error GoTo 0
    PickList = Application.GetOpenFilename("Texas license list (*.csv),*.csv", 1, _
        "Pick the license list you downloaded (permitted-fuels-taxpayers.csv)")
End Function


' Digits only, leading zeros dropped, so a taxpayer number stored as
' a number, as text, or with dashes or spaces in it still matches.
Private Function TaxpayerKey(ByVal v As Variant) As String
    Dim s As String, t As String, ch As String
    Dim j As Long

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbString Then
        s = v
    ElseIf IsNumeric(v) Then
        s = Format$(v, "0")
    Else
        s = CStr(v)
    End If

    For j = 1 To Len(s)
        ch = Mid$(s, j, 1)
        If ch >= "0" And ch <= "9" Then t = t & ch
    Next j
    Do While Len(t) > 1 And Left$(t, 1) = "0"
        t = Mid$(t, 2)
    Loop
    If t = "0" Then t = ""

    TaxpayerKey = t
End Function


Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CellText = CStr(v)
End Function


' The state leaves the suspended column blank for a license in good standing.
Private Function IsYes(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CellText(v)))
    IsYes = Len(s) > 0 And s <> "N" And s <> "NO" And s <> "FALSE" And s <> "0"
End Function


Private Function IsInList(ByVal s As String, ByVal list As Variant) As Boolean
    Dim j As Long
    For j = LBound(list) To UBound(list)
        If s = list(j) Then
            IsInList = True
            Exit Function
        End If
    Next j
End Function


Private Function SectionColour(ByVal sec As Long) As Long
    Select Case sec
        Case 1: SectionColour = RGB(255, 199, 206)    ' licensed, not supplier -- red
        Case 2: SectionColour = RGB(252, 229, 205)    ' not on the list -- orange
        Case Else: SectionColour = RGB(217, 217, 217) ' check by hand -- grey
    End Select
End Function


' A seller's name can start with TOTAL, so only a whole-word total in
' column A counts, or the word anywhere in the taxpayer number column.
Private Function IsTotalOrBlankRow(ws As Worksheet, ByVal r As Long, ByVal idCol As Long) As Boolean
    Dim colA As String, colId As String

    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
        IsTotalOrBlankRow = True
        Exit Function
    End If

    colA = UCase$(Trim$(ws.Cells(r, 1).Text))
    colId = UCase$(Trim$(ws.Cells(r, idCol).Text))

    IsTotalOrBlankRow = (InStr(colId, "TOTAL") > 0) _
                     Or (colA = "TOTAL") _
                     Or (Right$(colA, 6) = " TOTAL") _
                     Or (Left$(colA, 11) = "GRAND TOTAL")
End Function


Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        LastUsedRow = 0
    Else
        LastUsedRow = f.Row
    End If
End Function


Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = f.Column
    End If
End Function


Private Sub InsertBanner(ws As Worksheet, ByVal atRow As Long, ByVal txt As String)
    ws.Range(ws.Rows(atRow), ws.Rows(atRow + 1)).Insert Shift:=xlDown
    ws.Range(ws.Rows(atRow), ws.Rows(atRow + 1)).Interior.Pattern = xlNone
    With ws.Cells(atRow + 1, 1)
        .Value = txt
        .Font.Bold = True
    End With
End Sub


Private Function SafeSheetName(wb As Workbook, ByVal nm As String) As String
    Dim stem As String, attempt As String
    Dim c As Variant, i As Long

    stem = nm
    For Each c In Array("\", "/", "*", "?", ":", "[", "]")
        stem = Replace$(stem, CStr(c), "-")
    Next c
    stem = Left$(Trim$(stem), 27)

    attempt = stem
    i = 1
    Do While SheetExists(wb, attempt)
        i = i + 1
        attempt = stem & " " & i
    Loop
    SafeSheetName = attempt
End Function


Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not (s Is Nothing)
End Function
