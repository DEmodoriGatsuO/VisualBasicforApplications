Attribute VB_Name = "SheetsControl"
Option Explicit
'==================================
'アクティブワークブックの全シートの列幅調整
'==================================
Sub columnsSizeAutoFit()

    Dim ws  As Worksheet
    Dim Ans As String: _
            Ans = MsgBox(ActiveWorkbook.Name & " の全てのシートの列幅調整を行います。", vbInformation + vbYesNo, "Question")

    If Ans = vbNo Then Exit Sub

    With ActiveWorkbook
        For Each ws In Worksheets
            ws.Activate
            Range(Cells(ActiveCell.Row, ActiveCell.Column), _
                    Cells(ActiveCell.Row, ActiveCell.Column)).CurrentRegion.Select
            Selection.Columns.AutoFit
            Range("A1").Select
        Next ws
    End With
    
End Sub
'==================================
'値が0もしくは空白行の一括削除
'==================================
Sub ZeroValue_Blank_CellsEntireRowDelete()

    Dim r   As Range
    Dim Ans As String: _
            Ans = MsgBox("Is it really okay?", vbCritical + vbYesNo, "Infomation")
    
    If Ans = vbNo Then Exit Sub
    If Selection.Columns.Count > 1 Then
        MsgBox "Impossible operation", vbCritical
        Exit Sub
    End If
    For Each r In Selection
        If r.Value = 0 Then r.ClearContents
    Next r
    
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
End Sub
'==================================
'アクティブワークブックの全シートのカーソルをA1に設定
'==================================
Sub ActiveteA1()
    Dim ws As Worksheet
    Dim tmp As String
    Dim Ans As String: _
            Ans = MsgBox("Are You Ready?", vbInformation + vbYesNo, "Infomation")
            
    If Ans = vbNo Then Exit Sub
    
    With ActiveWorkbook
        tmp = ActiveSheet.Name
        For Each ws In Worksheets
            ws.Activate
            Range("A1").Select
        Next ws
        .Sheets(tmp).Select
    End With

End Sub
'==================================
'アクティブワークブックの全シートのパスワードを適当に設定する
'==================================
Sub AllSheetProtect()
    Dim inp    As String: inp = InputBox("Password", "Password", "text")
    Dim ws     As Worksheet
    If inp = "" Then Exit Sub
    For Each ws In Worksheets
        ws.Activate
        ActiveSheet.Protect _
            Password:=inp, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Next ws
End Sub
'==================================
'選択した範囲の列のアルファベットを返す *:*
'==================================
Sub getColumnsAddress()

    If Selection.Rows.Count > 1 Then Exit Sub
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

    Dim rangeArray  As Variant
    Dim selectRange As Range
    Dim rowNum      As Long
    Dim i As Integer: i = 0
    ReDim rangeArray(Selection.Columns.Count)
    
    For Each selectRange In Selection
        rowNum = selectRange.Row
        rangeArray(i) = Replace(selectRange.Address(RowAbsolute:=False, ColumnAbsolute:=False), rowNum, "") _
                        & ":" & Replace(selectRange.Address(RowAbsolute:=False, ColumnAbsolute:=False), rowNum, "")
        i = i + 1
    Next selectRange
    Selection = rangeArray
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

