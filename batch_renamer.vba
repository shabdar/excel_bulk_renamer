Sub batch_rename()
    On Error GoTo errHndl
   
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Dim sourcePath As String, destPath As String
    Dim sourceFile As String, destFile As String, sourceExtension As String
    Dim rng As Range, cell As Range, row As Range
   
    sourcePath = "\path to old files\"
    destPath = "\path to new files\"
    sourceFile = ""
    destFile = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rng = ActiveSheet.Range("A2", "B10")
 
    For Each row In rng.Rows
        sourceExtension = Split(Trim(row.Cells(, 2)), ".")(1)
        sourceFile = sourcePath + Trim(row.Cells(, 2))
        destFile = destPath + Trim(row.Cells(, 1)) + "." + sourceExtension
        fso.CopyFile sourceFile, destFile, False
    Next row
   
    MsgBox "Yay! Operation was successful.", vbOKOnly + vbInformation, "Done"
    Exit Sub
 
errHndl:
    MsgBox "Error happened while working on: " + vbCrLf + _
        sourceFile + vbCrLf + vbCrLf + "Error " + _
        Str(Err.Number) + ": " + Err.Description, vbCritical + vbOKOnly, "Error"
 
End Sub