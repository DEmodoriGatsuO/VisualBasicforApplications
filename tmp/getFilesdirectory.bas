Attribute VB_Name = "Module1"
Option Explicit

Sub getFilesdirectory()
    Const filesDirectory = "\*.*"
    Dim strDirectory As String, strPathName As String
    Dim myCol As New Collection
    Dim myArray
    Dim i As Long
    With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then
        strDirectory = .SelectedItems(1)
      End If
    End With
    If strDirectory = "" Then Exit Sub
    strPathName = Dir(strDirectory & filesDirectory, vbNormal)
    myCol.Add "Directorys"
    Do While strPathName <> "" ' ファイルが見つからなくなるまでLoop
        myCol.Add strPathName
        strPathName = Dir()
    Loop
    ReDim myArray(1 To myCol.Count + 1, 1 To 1)
    For i = 1 To myCol.Count
        myArray(i, 1) = myCol(i)
    Next i
    With ActiveWorkbook.ActiveSheet
        .Range(.Cells(ActiveCell.Row, ActiveCell.Column), .Cells(ActiveCell.Row + myCol.Count, ActiveCell.Column)) = myArray
    End With
End Sub
