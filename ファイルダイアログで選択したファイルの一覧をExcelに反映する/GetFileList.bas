Attribute VB_Name = "GetFileList"
Option Explicit
'Project Name    : ファイルダイアログで選択したファイルの一覧をExcelに反映する
'File Name       : GetFileList.bas
'Feature         : アクティブセルにファイルの一覧を出力できます(^▽^)
'Creation Date   : 2022.05.02
'Programming language used:
'' Visual Basic for Application
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.
'Copyright (c) 2022, Tech Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

'==================================
'GetFileList Main Module
'==================================
Sub getFileList()
    
    '----定数・変数一覧
    Const cnsDir     As String = "\*.*" '拡張子
    Dim dirCol       As New Collection    'ファイルのフルパスを一度全て格納するコレクション
    Dim strPathName  As String  'FileのPath
    Dim strFileName  As String  'FileのName
    Dim colLine      As Variant: colLine = Array("FullPath", "Filename") 'ヘッダー
    Dim c            As Variant 'コレクション用のイテレータ
    Dim i            As Long    '配列用 一次元インデックス
    Dim j            As Long    '配列用 二次元インデックス
    
    '1. ファイルダイアログを開いて一覧化するフォルダを選択する
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With
    
    '2. ファイルが存在しないフォルダでは作動しません
    strFileName = Dir(strPathName & cnsDir)
    If strFileName = "" Then
        MsgBox "ファイルが存在しません", vbCritical, "Error"
        Exit Sub
    End If
    
    '3. ファイルが存在するフォルダであればヘッダーをセットします
    dirCol.Add colLine
    dirCol.Add Array((strPathName & "\" & strFileName), strFileName)
    
    '4. Dir()関数にヒットした数だけコレクションに追加していきます。
    Do While strFileName <> ""
        strFileName = Dir()
        If strFileName <> "" Then dirCol.Add Array((strPathName & "\" & strFileName), strFileName)
    Loop
    
    '5. コレクションを二次元配列に変えます
    ReDim dirarr(dirCol.Count, LBound(dirCol(1)) To UBound(dirCol(1)))
    
    '6. Rangeに入力する上で一次元配列も二次元配列もLboundは0にしています。
    ''各列の要素は全て2つになります。ヘッダーの数でバインドされます
    i = 0
    For Each c In dirCol
        For j = LBound(c) To UBound(c)
            dirarr(i, j) = c(j)
        Next j
        i = i + 1
    Next c
    
    '7. シートオブジェクトの指定を外しているのでアクティブセルが起算になります（Attention：注意！！）
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row + UBound(dirarr, 1), ActiveCell.Column + UBound(dirarr, 2))) = dirarr
    
    '8. 作業完了Msgbox
    MsgBox strPathName & vbNewLine & "ファイル一覧を出力しました。", vbInformation, "Success"
    
End Sub
