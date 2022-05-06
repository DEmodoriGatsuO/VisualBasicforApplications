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

'���̃��W���[���ŗ��p����萔�ꗗ
Private Const SHEETS_NAME_SOURCE   As String = "original" 'Sheet1�̖��O
Private Const COLUMN_DECLARE       As Integer = 1 'Table(Original_Data)��declare�̗�ԍ�
Private Const COLUMN_RETURN_VALUE  As Integer = 2 'Table(Original_Data)��return value�̗�ԍ�
Private Const COLUMN_CALL_FUNCTION As Integer = 3 'Table(Original_Data)��call function�̗�ԍ�
Private Const SHEETS_NAME_REPLACE  As String = "replace"  'Sheet2�̖��O
Private Const COLUMN_INDEX         As Integer = 1 'Table(Replacement)��index�̗�ԍ�
Private Const COLUMN_PATTERN       As Integer = 2 'Table(Replacement)��pattern�̗�ԍ�
Private Const COLUMN_REPLACE       As Integer = 3 'Table(Replacement)��replace�̗�ԍ�
Private Const TABLE_DATA_ADDRESS   As String = "$A$1" 'Table�����݂���A�h���X
Private Const STR_INDENT_M         As String = "    " 'Power Query-M���p�̃C���f���g
Private Const STR_EQUAL_M          As String = " = " 'Power Query-M�� equal
Private Const NEXT_STEP_INCREMENT  As Integer = 1 'Power Query-M�� ���̃X�e�b�v�̍s�ԍ��̑���
Private Const LAST_STEP_INCREMENT  As Integer = 2 'Power Query-M�� �Ō�̃X�e�b�v�̍s�ԍ��̑���
'���C��
Sub main()
    '--���̃v���V�[�W���ŗ��p����ϐ��ꗗ
    Dim msg             As String: msg = MsgBox("Are you sure you want to run?", vbYesNo + vbInformation, "Confirmation") '��ƊJ�n�̊m�F
    Dim output_txt_Path As String: output_txt_Path = ThisWorkbook.Path & "\editor_text.txt" '���΃p�X�ł���Workbook������t�H���_�Ƀe�L�X�g�t�@�C��(.txt)���쐬
    Dim write_TXT       As String: write_TXT = replacePowerQuery_M '������̍쐬��Private Function replacePowerQuery_M�ɂč쐬
    
    '1. ��ƊJ�n�̊m�F
    If msg = vbNo Then Exit Sub
    
    '2. �e�L�X�g�t�@�C���ɏڍ׃G�f�B�^�[�p��M����������o���i���S�㏑���A�f�[�^�������ꍇ�͍쐬)
    Open output_txt_Path For Output As #1
        Print #1, write_TXT
    Close #1
    
    '3. �I�����b�Z�[�W
    MsgBox "Work is complete!!", vbInformation, "Success"

End Sub
'�ڍ׃G�f�B�^�[�p�ϊ�
Private Function replacePowerQuery_M() As String
    '--���̃v���V�[�W���ŗ��p����ϐ��ꗗ
    Dim source_value              'Table(Original_Data)�̒l
    Dim replace_value             'Table(Replacement)�̒l
    Dim max_index      As Long    '�C���f�b�N�X�̍ő�l�̎擾
    Dim i              As Long    'loop�p�̃C�e���[�^
    Dim str_pattern    As String  '�ϊ��O�̕�����
    Dim str_replace    As String  '�ϊ���̕�����
    Dim str_expression As String  '�ϊ���̕�������܂񂾕�����S��
    Dim return_string  As String  '�����o���p�̕�����
    Dim str_line       As String  '�����o���p��������쐬���邽�߂̍s
    
    '--WORKING SECTION IN THIS WORKBOOK--
    With ThisWorkbook
        '1. Sheet1(original)��Table(Original_Data)�̃w�b�_�[�������S�Ă̒l���擾
        With .Sheets(SHEETS_NAME_SOURCE)
            source_value = .Range(TABLE_DATA_ADDRESS).ListObject.DataBodyRange.Value
        End With
        
        '2. Sheet1(replace)��Table(Replacement)�̃w�b�_�[�������S�Ă̒l���擾
        With .Sheets(SHEETS_NAME_REPLACE)
            replace_value = .Range(TABLE_DATA_ADDRESS).ListObject.DataBodyRange.Value
        End With
        
    End With
    
    '3. �u������������ɂ�����C���f�b�N�X�̍ő�l���擾����
    max_index = replace_value(UBound(replace_value, 1), COLUMN_INDEX)
    
    'Table(Replacement)�œ񎟌��z��(source_value)���㏑��
    For i = LBound(replace_value, COLUMN_INDEX) To UBound(replace_value, COLUMN_INDEX)
        
        str_pattern = replace_value(i, COLUMN_PATTERN)  '�ϊ��O�̕�����
        str_replace = replace_value(i, COLUMN_REPLACE)  '�ϊ���̕�����
        
        'Blank Cell https://docs.microsoft.com/ja-JP/power-query/power-query-ui#applied-steps
        '�󔒂��܂ޏꍇ�AValue�����H
        If InStr(str_pattern, " ") > 0 Then str_pattern = "#" & """" & str_pattern & """"
        If InStr(str_replace, " ") > 0 Then str_replace = "#" & """" & str_replace & """"
        
        '4. ���f�[�^ ���Ӂi�߂�l�j�̐ݒ�
        source_value(replace_value(i, COLUMN_INDEX), COLUMN_RETURN_VALUE) = str_replace
        
        '5. ���f�[�^ �E�Ӂi�߂�l�j�̐ݒ�
        Select Case replace_value(i, COLUMN_INDEX)
            Case max_index
                Rem �ŏI�s�̏ꍇ�� in �̌�̍��Ӂi�߂�l�j���ݒ肷��
                Rem �ő�l+2��in�̌�ƒ�` LAST_STEP_INCREMENT = 2
                source_value(replace_value(i, COLUMN_INDEX) + LAST_STEP_INCREMENT, COLUMN_RETURN_VALUE) = str_replace
                
            Case Else
                Rem �E�ӂ̐ݒ��index(replace_value(i, 1) + 1)���Ώ�
                Rem �֐��̒��g�̓���ւ� NEXT_STEP_INCREMENT = 1
                str_expression = source_value(replace_value(i, COLUMN_INDEX) + NEXT_STEP_INCREMENT, COLUMN_CALL_FUNCTION)
                str_expression = _
                    Replace(str_expression, str_pattern, str_replace)
                
        End Select
        
        ''�E�ӗp�ɒu��������̕��������
        source_value(replace_value(i, COLUMN_INDEX) + NEXT_STEP_INCREMENT, COLUMN_CALL_FUNCTION) = str_expression
        
    Next i
    
    '6. �e�L�X�g���쐬
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
    
    '7. ������S�̂�Ԃ�
    replacePowerQuery_M = return_string

End Function
