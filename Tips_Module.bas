Attribute VB_Name = "Tips_Module"
Option Explicit
'Project Name    : Excel VBA Tips
'File Name       : Tips_Module.bas
'Feature         : 大好きな小技をアップしていきます!(^^)!
'Creation Date   : 2022.05.04 - Updated from time to time
'Programming language used:
'' Visual Basic for Application
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.
'Copyright (c) 2022, Tech Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

'==================================
'Tips　テキストファイルに文字列出力 引数1はパス、引数2は文字列
'==================================
Private Sub outputTextFile(targetPath, txt)
    '参照設定対策のためCreateObject採用
    'シェルをたたく準備をする
    Dim wsh
    Set wsh = CreateObject("Wscript.Shell")
    
    '書き込みモード（既存のパスのファイルは上書き、パスが無い場合は新規作成でテキストファイル書き出し
    Open targetPath For Output As #1
        Print #1, txt
    Close #1
    
    'ウインドウの最前面にテキストファイルを表示
    wsh.Run targetPath, 3
    Set wsh = Nothing
    
End Sub
'==================================
'Tips　ファイルサーバーのカレントディレクトリを設定する
'==================================
Private Sub command_cd(argv)
    '参照設定対策のためCreateObject採用
    'command cd
    With CreateObject("WScript.Shell")
        .CurrentDirectory = argv
    End With
End Sub
'==================================
'Tips　アクティブワークブックの全シートの列幅調整
'==================================
Sub activeworkbook_allSheets_columnsSizeAutofit()
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
'Tips　値が0もしくは空白行の一括削除
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
'Tips　アクティブワークブックの全シートのカーソルをA1に設定
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
'Tips　アクティブワークブックの全シートのパスワードを適当に設定する
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
'Tips　選択した範囲の列のアルファベットを返す *:*
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
