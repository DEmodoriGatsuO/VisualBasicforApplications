Attribute VB_Name = "GetWorkbookAllSheet"
Option Explicit
Sub getBookAllSheets()

    Dim myCol   As New Collection
    Dim c       As Variant
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim wbName  As String
    
    Const colKey As String = "status"
    Set myCol = getPathByPicker_xls
    If myCol(colKey) = "False" Then
        Set myCol = Nothing
        Exit Sub
    End If

    'エラーは無視
    On Error Resume Next
    Application.DisplayAlerts = False
    
    Set wb = Workbooks.Add
    
    For Each c In myCol
        Workbooks.Open Filename:=c
        wbName = Dir(c)
        For Each ws In Worksheets
            ws.Copy after:=wb.ActiveSheet
        Next ws
        Workbooks(wbName).Close
    Next c

    Application.DisplayAlerts = True
    Set myCol = Nothing
    Set wb = Nothing

End Sub
'==================================
'Folderの一覧をコレクションに格納する関数
'==================================
Private Function getPathByPicker_xls() As Collection

    Dim myCol   As New Collection
    Dim strPathName As String
    Dim strFileName As String
    Const cnsDIR = "\*.xls*"
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            strPathName = .SelectedItems(1)
            myCol.Add "Items", "status"
        Else
            myCol.Add "False", "status"
            Set getPathByPicker_xls = myCol
            Set myCol = Nothing
            Exit Function
        End If
    End With

    strFileName = Dir(strPathName & cnsDIR)
    myCol.Add (strPathName & "\" & strFileName)
    
    strFileName = Dir()
    Do While strFileName <> ""
        myCol.Add (strPathName & "\" & strFileName)
        strFileName = Dir()
    Loop
    
    Set getPathByPicker_xls = myCol
    Set myCol = Nothing

End Function
