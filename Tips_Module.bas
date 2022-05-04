Attribute VB_Name = "Tips_Module"
Option Explicit
'Project Name    : Excel VBA Tips
'File Name       : Tips_Module.bas
'Feature         : ��D���ȏ��Z���A�b�v���Ă����܂�!(^^)!
'Creation Date   : 2022.05.04 - Updated from time to time
'Programming language used:
'' Visual Basic for Application
'Author          : DEmodoriGatsuO https://github.com/DEmodoriGatsuO
'Twitter         : https://twitter.com/DemodoriGatsuo Follow Me!
'Sorry           : I like English. But I can't use English.
'Copyright (c) 2022, Tech Lovers. All rights reserved
'I can't use English, so I'll write in Japanese from now on.

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
'==================================
'Tips�@�t�@�C���T�[�o�[�̃J�����g�f�B���N�g����ݒ肷��
'==================================
Private Sub command_cd(argv)
    '�Q�Ɛݒ�΍�̂���CreateObject�̗p
    'command cd
    With CreateObject("WScript.Shell")
        .CurrentDirectory = argv
    End With
End Sub
