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
