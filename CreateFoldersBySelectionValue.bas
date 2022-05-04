Attribute VB_Name = "CreateFoldersBySelectionValue"
Option Explicit
'Project Name    : �t�@�C���_�C�A���O�őI�������t�H���_�̒����ɑI�������Z���̒l�ɉ����ăt�H���_���쐬���܂�
'File Name       : CreateFoldersBySelectionValue.bas
'Feature         : �C�y�Ƀt�H���_������@�\�ł��I�G���[�΍������͂��܂����I!(^^)!
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

    '----�萔�E�ϐ��ꗗ
    Const successPrompt As String = "Everything was completed successfully" '����ɏI�������ۂ̃v�����v�g
    Const successTitle  As String = "Success" '����ɏI�������ۂ̃^�C�g��
    Dim strPathName     As String  'File��Path
    Dim dat             As Variant '�I��͈͂̒l���i�[����񎟌��z��
    Dim col             As New Collection '�쐬����t�H���_���܂Ƃ߂�R���N�V����
    Dim c               As Variant '�R���N�V�����p�̃C�e���[�^
    Dim errCol          As New Collection '�G���[�̃��O���c�����߂̃R���N�V����
    Dim errLine         As Variant '�R���N�V�����̃e�L�X�g�t�@�C���̍s
    Dim i               As Long '�����^�C�e���[�^
    Dim msgStr          As String 'MsgBox�̕�����
    Dim logPath         As String '���O���o�͂���p�X
    
    '1. �t�@�C���_�C�A���O���J���Ĉꗗ������t�H���_��I������
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With
    
    '2. �I��͈͂̒l�� dat �Ɋi�[����
    dat = Selection.Value
    
    '3. �񎟌��z��̒l���R���N�V�����ɂ܂Ƃ߂�
    ''�쐬�Ώۂł͂Ȃ��󔒃Z���͏��O����
    ''' Infomation
    '''�I��͈͂���̏ꍇ�Adat�͔z��ɂȂ�Ȃ�
    '''���ɑI���Z�����󔒂̏ꍇ�A��Ƃ͈Ӗ����Ȃ��Ȃ��̂ŏI������
    Select Case IsArray(dat)
        Case False
            If dat = "" Then Exit Sub
            '�I���Z���̒l���R���N�V�����Ɋi�[����
            col.Add dat
        Case True
            '�I���Z���̒l��S�ăR���N�V�����Ɋi�[����
            For Each c In dat
                If c <> "" Then col.Add c
            Next c
    End Select
    
    '4. �R���N�V�����̃A�C�e���𔽕����ăt�H���_���쐬����
    '' �G���[�l�͖���
    '' ex
    '' �����ȕ����܂��͏d������t�H���_��
    '''�@�G���[��errCol�Ƃ����R���N�V�����Ɋi�[����
    For Each c In col
        On Error Resume Next
        MkDir strPathName & "\" & c
        If Err.Number <> 0 Then
            errLine = Array(c, Err.Number, Err.Description)
            errCol.Add errLine
            Err.Clear
        End If
    Next c
    
    '5. �G���[�������ꍇ�̓��b�Z�[�W�{�b�N�X���o���ďI������
    '' �t�H���_�͖��Ȃ��쐬����Ă���
    If errCol.Count = 0 Then
        MsgBox successPrompt, vbInformation, successTitle
        Exit Sub
    End If
    
    '!!�G���[�����݂���ꍇ!!
    If errCol.Count <> 0 Then
        'errCol��Item�͗v�f3�̔z��ɂȂ�̂ŕ����񉻂���msgStr�Ƃ����ϐ��ɏd�˂�
        For i = 1 To errCol.Count
            Select Case i
                Case 1 '�ŏ��̃C���f�b�N�X�̓^�C�g�����Z�b�g
                    msgStr = "Error Log" & vbCrLf & Join(errCol(i), " ") & vbCrLf
                Case errCol.Count  '�Ō�̃C���f�b�N�X�͍��v���Z�b�g
                    msgStr = msgStr & Join(errCol(i), " ") & vbCrLf & errCol.Count & "��"
                Case Else
                    msgStr = msgStr & Join(errCol(i), " ") & vbCrLf
            End Select
        Next i
    End If
    
    ''����VBA�t�@�C�����o�C���h����Ă���t�H���_���o�̓t�H���_�ɂ���Blog�t�@�C������now�֐��Ńt�H�[�}�b�g����̂ŏd���͋N����Ȃ��O��
    logPath = ThisWorkbook.path & "\" & Format(Now(), "yyyymmddhhmmss") & "error_log.txt"
    
    '' ��Private Sub�ɔ�ԁ@Tips�ɋ����܂����e�L�X�g�t�@�C���ɏ����o���Ashell�ŕ\������e�N�j�b�N�ł��I
    Call outputTextFile(logPath, msgStr)
    
End Sub
'==================================
'Tips�@�e�L�X�g�t�@�C���ɕ�����o�� ����1�̓p�X�A����2�͕�����
'==================================
Private Sub outputTextFile(targetPath, txt)
    '�Q�Ɛݒ�΍�̂���CreateObject�̗p
    '�V�F��������������������
    Dim wsh
    Set wsh = CreateObject("Wscript.Shell")
    
    '�������݃��[�h�i�����̃p�X�̃t�@�C���͏㏑���A�p�X�������ꍇ�͐V�K�쐬�Ńe�L�X�g�t�@�C�������o��
    Open targetPath For Output As #1
        Print #1, txt
    Close #1
    
    '�E�C���h�E�̍őO�ʂɃe�L�X�g�t�@�C����\��
    wsh.Run targetPath, 3
    Set wsh = Nothing
    
End Sub
