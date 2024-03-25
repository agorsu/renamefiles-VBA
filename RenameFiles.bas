Sub SelectFolder()

Dim FldrPicker As FileDialog
Dim myFolder As String

Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

With FldrPicker
  .Title = "Select A Target Folder"
  .AllowMultiSelect = False
  If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
  myFolder = .SelectedItems(1) & "\"
End With
  
ActiveWorkbook.Sheets("Rename").Range("B1").Value = myFolder

With Sheet1.ListObjects("Files")
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Delete
    End If
End With

listfiles (myFolder)

End Sub
Function listfiles(ByVal sPath As String)

    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Function

    i = 1
    For Each oFile In oFiles
        ActiveWorkbook.Sheets("Rename").Range("A" & i + 5).Value = oFile.Name
        i = i + 1
    Next

End Function

Public Sub RenameFiles()
Dim i As Long
cFolder = Sheets("Rename").Range("B1").Value
    For Each r In Sheet1.ListObjects("Files").ListRows
        Name cFolder & r.Range(1).Value As cFolder & r.Range(2).Value
        Debug.Print "Rename: " & r.Range(1).Value & " -> " & r.Range(2).Value
        i = i + 1
    Next r
MsgBox (i & " Files renamed")
End Sub
