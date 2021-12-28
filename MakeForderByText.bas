Attribute VB_Name = "makeForderByText"
Option Explicit
'==================================
 'mainプロシージャ
'==================================
Sub MakeFoldersofDirectory()
    Dim strDirectory As String
    Dim r As Range
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            strDirectory = .SelectedItems(1)
        End If
    End With
    If strDirectory = "" Then Exit Sub
    Select Case IsArray(Selection)
        Case False
            If Selection = "" Then Exit Sub
            Call MakeFolder(strDirectory & "" & Selection.Text)
        Case True
            For Each r In Selection
                If r.Text = "" Then Exit For
                Call MakeFolder(strDirectory & "\" & r.Text)
            Next r
    End Select
    MsgBox "Complete", vbInformation, "info"

End Sub
'==================================
'実際にフォルダを作るプロシージャ
'==================================
Private Sub MakeFolder(ByVal FolderPath As String)
    If Dir(FolderPath) <> "" Then Exit Sub
    On Error Resume Next
    Rem 無理な表現は全部resume next ではじく
    MkDir FolderPath
End Sub
