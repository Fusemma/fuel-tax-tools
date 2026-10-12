' Maps a shared drive: writes the folder tree to a text file on the desktop.
' Use: blank workbook, Alt+F11, Insert > Module, paste this, press F5.
' Set the path below first. Reads names only, never opens a file.

Sub ListShareFolders()

    ' Settings
    Const SHARE_PATH As String = "S:\Tax"   ' the drive to map, copied from Explorer's address bar
    Const LIST_FILES As Boolean = False     ' True = list every file too, not just folders

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim deskPath As String
    deskPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    Dim outPath As String
    If LIST_FILES Then
        outPath = deskPath & "\share-files.txt"
    Else
        outPath = deskPath & "\share-map.txt"
    End If

    Dim outFile As Object
    Set outFile = fso.CreateTextFile(outPath, True)

    WalkFolder fso.GetFolder(SHARE_PATH), outFile, LIST_FILES

    outFile.Close
    MsgBox "Done. Saved to: " & outPath

End Sub

Private Sub WalkFolder(fld As Object, outFile As Object, listFiles As Boolean)

    On Error Resume Next   ' folders without access get skipped

    outFile.WriteLine fld.Path

    If listFiles Then
        Dim f As Object
        For Each f In fld.Files
            outFile.WriteLine f.Path
        Next f
    End If

    Dim subFld As Object
    For Each subFld In fld.SubFolders
        WalkFolder subFld, outFile, listFiles
    Next subFld

End Sub
