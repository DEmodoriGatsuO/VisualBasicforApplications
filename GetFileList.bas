Attribute VB_Name = "GetFileList"
Option Explicit
'==================================
'main
'==================================
Sub getDirectory()
    Dim myCol   As New Collection
    Dim c       As Variant
    Dim dirCol  As New Collection
    Dim colLine As Variant
    Dim cnt     As Integer
    Dim myArray As Variant
    Dim i       As Integer
    Dim j       As Integer
    
    Const colKey As String = "status"
    Set myCol = getPathByPicker
    If myCol(colKey) = "False" Then
        Set myCol = Nothing
        Exit Sub
    End If
    
    cnt = 0
    For Each c In myCol
        If c <> "Items" Then
            colLine = mergeArray(Array(c), Split(c, "\"))
            If cnt < UBound(colLine) Then cnt = UBound(colLine)
            dirCol.Add colLine
        End If
    Next c
    
    Set myCol = Nothing
    ReDim myArray(dirCol.Count - 1, cnt)
    
    i = 0
    For Each c In dirCol
        For j = LBound(c) To UBound(c)
            myArray(i, j) = c(j)
        Next j
        i = i + 1
    Next c
    Set dirCol = Nothing
    
    With ActiveSheet
        .Range(ActiveCell, _
            .Cells(ActiveCell.Row + UBound(myArray, 1), ActiveCell.Column + UBound(myArray, 2))) = myArray
    End With

End Sub
'==================================
'Folderの一覧をコレクションに格納する関数
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
'==================================
'二つの一次元配列を一つの配列に集約する関数
'==================================
Private Function mergeArray(argArray1 As Variant, argArray2 As Variant)
    Dim arr: arr = Array(argArray1, argArray2)
    Dim var
    Dim cnt   As Integer: cnt = (UBound(argArray1) + 1) + (UBound(argArray2) + 1)
    Dim i     As Integer
    Dim j     As Integer
    ReDim merge(cnt)
    
    j = 0
    For Each var In arr
        For i = LBound(var) To UBound(var)
            merge(j) = var(i)
            j = j + 1
        Next i
    Next var
    
    mergeArray = merge

End Function
