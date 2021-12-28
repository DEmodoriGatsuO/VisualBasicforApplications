Attribute VB_Name = "SummarizeOneSheetAllData"
Option Explicit
Sub SummarizeOneSheetAllData()
    Dim wb As Workbook
    Dim myCol As New Collection
    Dim c As Variant
    Dim data As Variant
    Dim IRow As Integer
    Dim tmp  As Variant
    Dim extension As String

    Const colKey As String = "status"
    Set myCol = getPathByPicker
    If myCol(colKey) = "False" Then
        Set myCol = Nothing
        Exit Sub
    End If
    
    Set wb = Workbooks.Add

    On Error Resume Next
    For Each c In myCol
        tmp = Split(c, ".")
        extension = tmp(UBound(tmp))
        If extension = "csv" Or InStr(extension, "xls") > 0 Then
            Workbooks.Open Filename:=c
            data = ActiveSheet.Range("A1").CurrentRegion
            With wb.ActiveSheet
                IRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
                .Range(.Cells(IRow, 1), .Cells(IRow + (UBound(data, 1) - 1), UBound(data, 2))) = data
            End With
            
            Application.DisplayAlerts = False
            Workbooks(Dir(c)).Close
            Application.DisplayAlerts = True
        End If
    Next c
End Sub

'==================================
'FolderÇÃàÍóóÇÉRÉåÉNÉVÉáÉìÇ…äiî[Ç∑ÇÈä÷êî
'==================================
Private Function getPathByPicker() As Collection

    Dim myCol   As New Collection
    Dim strPathName As String
    Dim strFileName As String
    Const cnsDIR = "\*.*"
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            strPathName = .SelectedItems(1)
            myCol.Add "Items", "status"
        Else
            myCol.Add "False", "status"
            Set getPathByPicker = myCol
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
    
    Set getPathByPicker = myCol
    Set myCol = Nothing

End Function
