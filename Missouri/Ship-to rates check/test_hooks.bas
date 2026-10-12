Option Compare Binary
Option Explicit

' Test-only entry points, loaded beside MORatesCode by test_access.py.
' Never shipped.

' SetUp, then the data folder this test uses.
Public Function TestSetUp(Optional ByVal folder As String = "") As String
    gQuiet = True
    SetUp
    TestSetUp = gLastMessage
    If folder <> "" Then CurrentDb.Execute "UPDATE Settings SET SettingValue='" & folder & _
        "' WHERE SettingName='DataFolder'", dbFailOnError
End Function

Public Function TestLoad(ByVal bPath As String, ByVal rPath As String) As String
    Dim t As Double, ok As Boolean
    gQuiet = True
    t = Timer
    ok = LoadQuarterFiles(bPath, rPath)
    TestLoad = IIf(ok, "OK ", "FAILED ") & Format$(Timer - t, "0.0") & "s: " & gLastMessage
End Function

' Street pieces for each "pre|name|suf|post" line, one result line each.
Public Function TestKeys(ByVal inPath As String, ByVal outPath As String) As String
    Dim fi As Integer, fo As Integer, s As String, p() As String, n As Long
    fi = FreeFile
    Open inPath For Input As #fi
    fo = FreeFile
    Open outPath For Output As #fo
    Do While Not EOF(fi)
        Line Input #fi, s
        p = Split(s, "|")
        Print #fo, StreetPieces(p(0), p(1), p(2), p(3))
        n = n + 1
    Loop
    Close #fo
    Close #fi
    TestKeys = CStr(n)
End Function

Public Function TestCityKeys(ByVal inPath As String, ByVal outPath As String) As String
    Dim fi As Integer, fo As Integer, s As String, n As Long
    fi = FreeFile
    Open inPath For Input As #fi
    fo = FreeFile
    Open outPath For Output As #fo
    Do While Not EOF(fi)
        Line Input #fi, s
        Print #fo, CityKeyOf(s)
        n = n + 1
    Loop
    Close #fo
    Close #fi
    TestCityKeys = CStr(n)
End Function

' The step one results, in the same columns as mo_changes.py writes. They
' live in this quarter's file in the data folder.
Public Function TestExport(ByVal folder As String) As String
    Dim db As DAO.Database, rs As DAO.Recordset, f As Integer, n1 As Long, n2 As Long
    Set db = DBEngine.OpenDatabase(QuarterFile(NowQuarter()), False, True)
    f = FreeFile
    Open folder & "\actual_changes.csv" For Output As #f
    Print #f, "code,old,new,cities,rows"
    Set rs = db.OpenRecordset("SELECT * FROM RateChanges", dbOpenSnapshot)
    Do While Not rs.EOF
        Print #f, Q(rs!JCode) & "," & Q(Num6(rs!OldUse)) & "," & Q(Num6(rs!NewUse)) & "," & _
            Q(Nz(rs!CitiesAll, "")) & "," & Q(Nz(rs!AddrRows, 0))
        n1 = n1 + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #f
    f = FreeFile
    Open folder & "\actual_moves.csv" For Output As #f
    Print #f, "zip,street,city,side,lo,hi,old_codes,new_codes,old,new"
    Set rs = db.OpenRecordset("SELECT * FROM Moves", dbOpenSnapshot)
    Do While Not rs.EOF
        Print #f, Q(rs!Zip) & "," & Q(rs!Street) & "," & Q(rs!City) & "," & Q(rs!Side) & "," & _
            Q(rs!LoTxt) & "," & Q(rs!HiTxt) & "," & Q(rs!OldCodes) & "," & Q(rs!NewCodes) & "," & _
            Q(Num6(rs!OldUse)) & "," & Q(Num6(rs!NewUse))
        n2 = n2 + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #f
    db.Close
    TestExport = n1 & " changes, " & n2 & " moves"
End Function

Public Function TestReportPdf(ByVal path As String, Optional ByVal rpt As String = "Use tax changes") As String
    DoCmd.OutputTo acOutputReport, rpt, acFormatPDF, path
    TestReportPdf = "ok"
End Function

Public Function TestFormPdf(ByVal path As String, ByVal frm As String) As String
    DoCmd.OutputTo acOutputForm, frm, acFormatPDF, path
    TestFormPdf = "ok"
End Function

Public Function TestLoadList(ByVal path As String) As String
    Dim t As Double, ok As Boolean
    gQuiet = True
    t = Timer
    ok = LoadShipToFile(path)
    If ok Then CheckShipTos
    TestLoadList = IIf(ok, "OK ", "FAILED ") & Format$(Timer - t, "0.0") & "s: " & gLastMessage & " | " & ShipToLine()
End Function

Public Function TestCheck() As String
    gQuiet = True
    CheckShipTos
    TestCheck = ShipToLine()
End Function

Public Function TestPin(ByVal acct As String, ByVal shipTo As String, ByVal code As String) As String
    CurrentDb.Execute "UPDATE ShipTos SET PinnedCode='" & code & "' WHERE AcctNo='" & acct & "' AND ShipToNo='" & _
        shipTo & "'", dbFailOnError
    TestPin = "ok"
End Function

Public Function TestExportShipTos(ByVal path As String) As String
    Dim rs As DAO.Recordset, f As Integer, n As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, "acct,ship_to,result,why,code_now,how_now,where_now,line_now,use_now,code_last,use_last,exempt,pinned,tax"
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM ShipTos", dbOpenSnapshot)
    Do While Not rs.EOF
        Print #f, Q(rs!AcctNo) & "," & Q(rs!ShipToNo) & "," & Q(rs!Result) & "," & Q(rs!Why) & "," & _
            Q(rs!CodeNow) & "," & Q(rs!HowNow) & "," & Q(rs!WhereNow) & "," & Q(rs!LineNow) & "," & _
            Q(Num6(rs!UseNow)) & "," & Q(rs!CodeLast) & "," & Q(Num6(rs!UseLast)) & "," & Q(rs!Exempt) & "," & _
            Q(rs!PinnedCode) & "," & Q(Num6(rs!Tax))
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #f
    TestExportShipTos = CStr(n)
End Function

' What's where: each quarter's file and its rows, this file's size, and the
' main form's lines.
Public Function TestInfo() As String
    Dim db As DAO.Database, s As String, q As Variant, f As String
    For Each q In Array(LastQuarter(), NowQuarter())
        If q <> "" Then
            f = QuarterFile(CStr(q))
            Set db = DBEngine.OpenDatabase(f, False, True)
            s = s & Mid$(f, InStrRev(f, "\") + 1) & ": " & Format$(FileLen(f) / 1048576, "0") & "MB, " & _
                db.OpenRecordset("SELECT Count(*) FROM Boundary").Fields(0) & " rows, " & _
                db.OpenRecordset("SELECT Count(*) FROM Rates").Fields(0) & " rates; "
            db.Close
        End If
    Next
    s = s & "this file: " & Format$(FileLen(CurrentDb.Name) / 1048576, "0.0") & "MB | " & LoadedLine() & " | " & _
        ChangesLine() & " | " & ListLine() & " | " & BigNumber("Fix") & " " & BigLabel("Fix") & ", " & _
        BigNumber("Map") & " " & BigLabel("Map") & ", " & BigNumber("OK") & " " & BigLabel("OK") & ", " & _
        BigNumber("Skipped") & " " & BigLabel("Skipped") & " | " & NextLine()
    TestInfo = s
End Function

' What one click on a 100879/4 box does.
Public Function TestCopy(ByVal acct As String, ByVal shipTo As String) As String
    TestCopy = IIf(CopyShipTo(acct, shipTo), "copied", "not copied")
End Function

Public Function TestOption(ByVal nm As String) As String
    TestOption = CStr(Application.GetOption(nm))
End Function

Private Function Q(ByVal v As Variant) As String
    Q = """" & Replace(Nz(v, ""), """", """""") & """"
End Function

Private Function Num6(ByVal v As Variant) As String
    If Not IsNull(v) Then Num6 = Format$(v, "0.000000")
End Function
