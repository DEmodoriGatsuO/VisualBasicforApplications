Attribute VB_Name = "GetFileList"
Option Explicit
'Project Name    : �t�@�C���_�C�A���O�őI�������t�@�C���̈ꗗ��Excel�ɔ��f����
'File Name       : GetFileList.bas
'Feature         : �A�N�e�B�u�Z���Ƀt�@�C���̈ꗗ���o�͂ł��܂�(^��^)
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
    Const cnsDir     As String = "\*.*" '�g���q
    Const colKey     As String = "status" 'getPathByPicker�̖߂�l�𔻒f���邽�߂ɐ݂���key
    Dim dirCol       As New Collection    '�t�@�C���̃t���p�X����x�S�Ċi�[����R���N�V����
    Dim strPathName  As String  'File��Path
    Dim strFileName  As String  'File��Name
    Dim colLine      As Variant: colLine = Array("FullPath", "Filename") '�w�b�_�[
    Dim c            As Variant '�R���N�V�����p�̃C�e���[�^
    Dim i            As Long    '�z��p �ꎟ���C���f�b�N�X
    Dim j            As Long    '�z��p �񎟌��C���f�b�N�X
    
    '1. �t�@�C���_�C�A���O���J���Ĉꗗ������t�H���_���Z���^k����
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With
    
    '2. �t�@�C�������݂��Ȃ��t�H���_�ł͍쓮���܂���
    strFileName = Dir(strPathName & cnsDir)
    If strFileName = "" Then
        MsgBox "�t�@�C�������݂��܂���", vbCritical, "Error"
        Exit Sub
    End If
    
    '3. �t�@�C�������݂���t�H���_�ł���΃w�b�_�[���Z�b�g���܂�
    dirCol.Add colLine
    dirCol.Add Array((strPathName & "\" & strFileName), strFileName)
    
    '4. Dir()�֐��Ƀq�b�g�����������R���N�V�����ɒǉ����Ă����܂��B
    Do While strFileName <> ""
        strFileName = Dir()
        If strFileName <> "" Then dirCol.Add Array((strPathName & "\" & strFileName), strFileName)
    Loop
    
    '5. �R���N�V������񎟌��z��ɕς��܂�
    ReDim dirarr(dirCol.Count, LBound(dirCol(1)) To UBound(dirCol(1)))
    
    '6. Range�ɓ��͂����ňꎟ���z����񎟌��z���Lbound��0�ɂ��Ă��܂��B
    ''�e��̗v�f�͑S��2�ɂȂ�܂��B�w�b�_�[�̐��Ńo�C���h����܂�
    i = 0
    For Each c In dirCol
        For j = LBound(c) To UBound(c)
            dirarr(i, j) = c(j)
        Next j
        i = i + 1
    Next c
    
    '7. �V�[�g�I�u�W�F�N�g�̎w����O���Ă���̂ŃA�N�e�B�u�Z�����N�Z�ɂȂ�܂��iAttention�F���ӁI�I�j
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row + UBound(dirarr, 1), ActiveCell.Column + UBound(dirarr, 2))) = dirarr
    
    '8. ��Ɗ���Msgbox
    MsgBox strPathName & vbNewLine & "�t�@�C���ꗗ���o�͂��܂����B", vbInformation, "Success"
    
End Sub
