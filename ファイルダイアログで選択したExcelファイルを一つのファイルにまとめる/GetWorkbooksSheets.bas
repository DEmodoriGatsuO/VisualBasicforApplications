Attribute VB_Name = "GetWorkbooksSheets"
Option Explicit
'Project Name    : ファイルダイアログで選択したフォルダの直下のエクセルファイルを一つのワークブックにまとめます
'File Name       : GetWorkbooksSheets.bas
'Feature         : Errorは一切無視の力技スタイルです！！新規で作るWorkbooksの後ろに加えていくので元ファイルを壊す心配もなし！(^▽^)
'Creation Date   : 2022.05.03
'Programming language used:
'' Visual Basic for Application
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.
'Copyright (c) 2022, Tech Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

'==================================
'getWorkbooksSheets Main Module
'==================================
Sub getWorkbooksSheets()

    '----定数・変数一覧
    Const cnsDir     As String = "\*.xls*" 'Excelの拡張子
    Dim dirCol       As New Collection    'ファイルのフルパスを一度全て格納するコレクション
    Dim strPathName  As String  'FileのPath
    Dim strFileName  As String  'FileのName
    Const sheetName  As String = "contents" '作成した新規ファイルには表紙を付けてデータのパスとファイル名、シート名を一覧化する
    Dim wb           As Workbook 'オブジェクト変数 - 新規で作成するワークブック
    Dim ws           As Worksheet 'オブジェクト変数 - ワークシートのイテレータ
    Dim wc           As Long 'ワークシートの数、毎回計算するのではなく、インクリメント形式で + 1をしていく
    Dim getCol       As New Collection '表紙用のコレクション arrという変数で二次元配列にまとめる
    Dim c            As Variant
    Dim colLine      As Variant: colLine = Array("FilePath", "FileName", "SheetName")   'コレクション用のイテレータ
    Dim i            As Long    '配列用 一次元インデックス
    Dim j            As Long    '配列用 二次元インデックス
    
    '1. ファイルダイアログを開いて一覧化するフォルダを選択する
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With

    '2. 拡張子に該当するファイルが存在しないフォルダでは作動しません
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

    ' * !!!エラーは一切無視の力技スタイル!!!
    On Error Resume Next
   
    ' - Workbookに付帯するエラーメッセージは全て無視する
    ' - 同時にマクロ高速化のため、ScreenUpdatingとEnableEventsを起動時の間OFFにする
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' - 新規Workbooksの一番最後のシートに追加する　Ctrl+Nの技と一緒です
    ' - Sheet1は追加していったファイルの詳細を書く目次にします
    Set wb = Workbooks.Add
    wb.Sheets(1).Name = sheetName
    
    ' * getColはファイルフルパス、ファイル名、シート名を蓄積するコレクションです
    ' - 拾い上げたブックの詳細を書くシートのためにコレクションでデータを蓄積する
    ' - ヘッダーのセット---colLine
    ' - wcはExcelのバージョン対策です
    
    getCol.Add colLine
    wc = Worksheets.Count
    
    '5. Copyの反復
    ' - Item1はヘッダーなので最初のインデックスは2
    For i = 2 To dirCol.Count
        
        ' - 配列 [ FullPath FileName SheetName ]
        colLine = Array(dirCol(i)(0), Dir(dirCol(i)(0)), "SheetName")
        
        ' - Excelブックを開く
        Workbooks.Open Filename:=colLine(0)
        
        ' - Excelブックのワークシートの数だけCopyを繰り返す
        For Each ws In Workbooks(colLine(1)).Worksheets
            ' - colLineにSheetNameを代入
            colLine(2) = ws.Name
            
            ' - 新規ワークブックの最後にコピー
            ws.Copy after:=wb.Sheets(wc)
            ' - インクリメント
            wc = wc + 1
            
            ' - コレクションに配列を追加
            getCol.Add colLine
        Next ws
        
        ' - 有無を言わさず消す！！
        Workbooks(colLine(1)).Close
        
    Next i
    
    ' エラー無視はここで解除する
    Application.DisplayAlerts = True

    '5. コレクションを二次元配列に変えます
    ReDim arr(getCol.Count, LBound(getCol(1)) To UBound(getCol(1)))
    
    '6. Rangeに入力する上で一次元配列も二次元配列もLboundは0にしています。
    ''各列の要素は全て2つになります。ヘッダーの数でバインドされます
    i = 0
    For Each c In getCol
        For j = LBound(c) To UBound(c)
            arr(i, j) = c(j)
        Next j
        i = i + 1
    Next c
    
    '7. シートオブジェクトの指定を外しているのでアクティブセルが起算になります（Attention：注意！！）
    With wb.Sheets(sheetName)
        .Range(.Cells(1, 1), .Cells(UBound(arr, 1) + 1, UBound(arr, 2) + 1)) = arr
        .Range(.Cells(1, 1), .Cells(UBound(arr, 1) + 1, UBound(arr, 2) + 1)).Columns.AutoFit
        .Select
    End With
    
    '8. オフにしていた機能をTrueに
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    '9. 作業完了Msgbox
    MsgBox wb.Name & vbNewLine & "こちらに選択したフォルダのExcelブックのシートをまとめました！", vbInformation, "Success"
    
    ' オブジェクト変数を解放
    Set wb = Nothing

End Sub
