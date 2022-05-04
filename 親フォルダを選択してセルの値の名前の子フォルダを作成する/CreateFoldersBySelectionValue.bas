Attribute VB_Name = "CreateFoldersBySelectionValue"
Option Explicit
'Project Name    : ファイルダイアログで選択したフォルダの直下に選択したセルの値に応じてフォルダを作成します
'File Name       : CreateFoldersBySelectionValue.bas
'Feature         : 気軽にフォルダが作れる機能です！エラー対策を今回はしました！!(^^)!
'Creation Date   : 2022.05.04
'Programming language used:
'' Visual Basic for Application
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.
'Copyright (c) 2022, Tech Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

'==================================
'createFoldersBySelectionValue Main Module
'==================================
Sub createFoldersBySelectionValue()

    '----定数・変数一覧
    Const successPrompt As String = "Everything was completed successfully" '正常に終了した際のプロンプト
    Const successTitle  As String = "Success" '正常に終了した際のタイトル
    Dim strPathName     As String  'FileのPath
    Dim dat             As Variant '選択範囲の値を格納する二次元配列
    Dim col             As New Collection '作成するフォルダをまとめるコレクション
    Dim c               As Variant 'コレクション用のイテレータ
    Dim errCol          As New Collection 'エラーのログを残すためのコレクション
    Dim errLine         As Variant 'コレクションのテキストファイルの行
    Dim i               As Long '整数型イテレータ
    Dim msgStr          As String 'MsgBoxの文字列
    Dim logPath         As String 'ログを出力するパス
    
    '1. ファイルダイアログを開いて一覧化するフォルダを選択する
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With
    
    '2. 選択範囲の値を dat に格納する
    dat = Selection.Value
    
    '3. 二次元配列の値をコレクションにまとめる
    ''作成対象ではない空白セルは除外する
    ''' Infomation
    '''選択範囲が一つの場合、datは配列にならない
    '''仮に選択セルが空白の場合、作業は意味をなさないので終了する
    Select Case IsArray(dat)
        Case False
            If dat = "" Then Exit Sub
            '選択セルの値をコレクションに格納する
            col.Add dat
        Case True
            '選択セルの値を全てコレクションに格納する
            For Each c In dat
                If c <> "" Then col.Add c
            Next c
    End Select
    
    '4. コレクションのアイテムを反復してフォルダを作成する
    '' エラー値は無視
    '' ex
    '' 無効な文字または重複するフォルダ名
    '''　エラーはerrColというコレクションに格納する
    For Each c In col
        On Error Resume Next
        MkDir strPathName & "\" & c
        If Err.Number <> 0 Then
            errLine = Array(c, Err.Number, Err.Description)
            errCol.Add errLine
            Err.Clear
        End If
    Next c
    
    '5. エラーが無い場合はメッセージボックスを出して終了する
    '' フォルダは問題なく作成されている
    If errCol.Count = 0 Then
        MsgBox successPrompt, vbInformation, successTitle
        Exit Sub
    End If
    
    '!!エラーが存在する場合!!
    If errCol.Count <> 0 Then
        'errColのItemは要素3つの配列になるので文字列化してmsgStrという変数に重ねる
        For i = 1 To errCol.Count
            Select Case i
                Case 1 '最初のインデックスはタイトルをセット
                    msgStr = "Error Log" & vbCrLf & Join(errCol(i), " ") & vbCrLf
                Case errCol.Count  '最後のインデックスは合計をセット
                    msgStr = msgStr & Join(errCol(i), " ") & vbCrLf & errCol.Count & "件"
                Case Else
                    msgStr = msgStr & Join(errCol(i), " ") & vbCrLf
            End Select
        Next i
    End If
    
    ''このVBAファイルがバインドされているフォルダを出力フォルダにする。logファイル名はnow関数でフォーマットするので重複は起こらない前提
    logPath = ThisWorkbook.path & "\" & Format(Now(), "yyyymmddhhmmss") & "error_log.txt"
    
    '' ☆Private Subに飛ぶ　Tipsに挙げますがテキストファイルに書き出し、shellで表示するテクニックです！
    Call outputTextFile(logPath, msgStr)
    
End Sub
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
