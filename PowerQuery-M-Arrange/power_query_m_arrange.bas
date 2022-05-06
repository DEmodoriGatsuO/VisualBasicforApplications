Attribute VB_Name = "power_query_m_arrange"
Option Explicit
'Project Name    : Power Query M-Editor Code Arrange
'File Name       : PowerQuery-M-Arrange.xlsm
'Creation Date   : 2022.04.30
'Update          : 2022.05.06
'                  - Deal with blanks
'                  - Rename Module
'Visual Basic for Applications
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.

'Copyright (c) 2022, Excel Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

'このモジュールで利用する定数一覧
Private Const SHEETS_NAME_SOURCE   As String = "original" 'Sheet1の名前
Private Const COLUMN_DECLARE       As Integer = 1 'Table(Original_Data)のdeclareの列番号
Private Const COLUMN_RETURN_VALUE  As Integer = 2 'Table(Original_Data)のreturn valueの列番号
Private Const COLUMN_CALL_FUNCTION As Integer = 3 'Table(Original_Data)のcall functionの列番号
Private Const SHEETS_NAME_REPLACE  As String = "replace"  'Sheet2の名前
Private Const COLUMN_INDEX         As Integer = 1 'Table(Replacement)のindexの列番号
Private Const COLUMN_PATTERN       As Integer = 2 'Table(Replacement)のpatternの列番号
Private Const COLUMN_REPLACE       As Integer = 3 'Table(Replacement)のreplaceの列番号
Private Const TABLE_DATA_ADDRESS   As String = "$A$1" 'Tableが存在するアドレス
Private Const STR_INDENT_M         As String = "    " 'Power Query-M式用のインデント
Private Const STR_EQUAL_M          As String = " = " 'Power Query-M式 equal
Private Const NEXT_STEP_INCREMENT  As Integer = 1 'Power Query-M式 次のステップの行番号の増分
Private Const LAST_STEP_INCREMENT  As Integer = 2 'Power Query-M式 最後のステップの行番号の増分
'メイン
Sub main()
    '--このプロシージャで利用する変数一覧
    Dim msg             As String: msg = MsgBox("Are you sure you want to run?", vbYesNo + vbInformation, "Confirmation") '作業開始の確認
    Dim output_txt_Path As String: output_txt_Path = ThisWorkbook.Path & "\editor_text.txt" '相対パスでこのWorkbookがあるフォルダにテキストファイル(.txt)を作成
    Dim write_TXT       As String: write_TXT = replacePowerQuery_M '文字列の作成はPrivate Function replacePowerQuery_Mにて作成
    
    '1. 作業開始の確認
    If msg = vbNo Then Exit Sub
    
    '2. テキストファイルに詳細エディター用のM言語を書き出し（完全上書き、データが無い場合は作成)
    Open output_txt_Path For Output As #1
        Print #1, write_TXT
    Close #1
    
    '3. 終了メッセージ
    MsgBox "Work is complete!!", vbInformation, "Success"

End Sub
'詳細エディター用変換
Private Function replacePowerQuery_M() As String
    '--このプロシージャで利用する変数一覧
    Dim source_value              'Table(Original_Data)の値
    Dim replace_value             'Table(Replacement)の値
    Dim max_index      As Long    'インデックスの最大値の取得
    Dim i              As Long    'loop用のイテレータ
    Dim str_pattern    As String  '変換前の文字列
    Dim str_replace    As String  '変換後の文字列
    Dim str_expression As String  '変換後の文字列を含んだ文字列全体
    Dim return_string  As String  '書き出し用の文字列
    Dim str_line       As String  '書き出し用文字列を作成するための行
    
    '--WORKING SECTION IN THIS WORKBOOK--
    With ThisWorkbook
        '1. Sheet1(original)のTable(Original_Data)のヘッダーを除く全ての値を取得
        With .Sheets(SHEETS_NAME_SOURCE)
            source_value = .Range(TABLE_DATA_ADDRESS).ListObject.DataBodyRange.Value
        End With
        
        '2. Sheet1(replace)のTable(Replacement)のヘッダーを除く全ての値を取得
        With .Sheets(SHEETS_NAME_REPLACE)
            replace_value = .Range(TABLE_DATA_ADDRESS).ListObject.DataBodyRange.Value
        End With
        
    End With
    
    '3. 置き換えをするにあたりインデックスの最大値を取得する
    max_index = replace_value(UBound(replace_value, 1), COLUMN_INDEX)
    
    'Table(Replacement)で二次元配列(source_value)を上書き
    For i = LBound(replace_value, COLUMN_INDEX) To UBound(replace_value, COLUMN_INDEX)
        
        str_pattern = replace_value(i, COLUMN_PATTERN)  '変換前の文字列
        str_replace = replace_value(i, COLUMN_REPLACE)  '変換後の文字列
        
        'Blank Cell https://docs.microsoft.com/ja-JP/power-query/power-query-ui#applied-steps
        '空白を含む場合、Valueを加工
        If InStr(str_pattern, " ") > 0 Then str_pattern = "#" & """" & str_pattern & """"
        If InStr(str_replace, " ") > 0 Then str_replace = "#" & """" & str_replace & """"
        
        '4. 元データ 左辺（戻り値）の設定
        source_value(replace_value(i, COLUMN_INDEX), COLUMN_RETURN_VALUE) = str_replace
        
        '5. 元データ 右辺（戻り値）の設定
        Select Case replace_value(i, COLUMN_INDEX)
            Case max_index
                Rem 最終行の場合は in の後の左辺（戻り値）も設定する
                Rem 最大値+2がinの後と定義 LAST_STEP_INCREMENT = 2
                source_value(replace_value(i, COLUMN_INDEX) + LAST_STEP_INCREMENT, COLUMN_RETURN_VALUE) = str_replace
                
            Case Else
                Rem 右辺の設定はindex(replace_value(i, 1) + 1)が対象
                Rem 関数の中身の入れ替え NEXT_STEP_INCREMENT = 1
                str_expression = source_value(replace_value(i, COLUMN_INDEX) + NEXT_STEP_INCREMENT, COLUMN_CALL_FUNCTION)
                str_expression = _
                    Replace(str_expression, str_pattern, str_replace)
                
        End Select
        
        ''右辺用に置き換え後の文字列を代入
        source_value(replace_value(i, COLUMN_INDEX) + NEXT_STEP_INCREMENT, COLUMN_CALL_FUNCTION) = str_expression
        
    Next i
    
    '6. テキストを作成
    For i = LBound(source_value, 1) To UBound(source_value, 1)
        Select Case source_value(i, COLUMN_DECLARE)
            Case "let", "in"
                return_string = return_string & source_value(i, COLUMN_DECLARE) & vbCrLf
            Case Else
                Select Case source_value(i, COLUMN_CALL_FUNCTION)
                    Case ""
                        str_line = STR_INDENT_M & source_value(i, COLUMN_RETURN_VALUE)
                    Case Is <> ""
                        str_line = STR_INDENT_M & source_value(i, COLUMN_RETURN_VALUE) & _
                            STR_EQUAL_M & source_value(i, COLUMN_CALL_FUNCTION) & vbCrLf
                End Select
                return_string = return_string & str_line
        End Select
    Next i
    
    '7. 文字列全体を返す
    replacePowerQuery_M = return_string

End Function
