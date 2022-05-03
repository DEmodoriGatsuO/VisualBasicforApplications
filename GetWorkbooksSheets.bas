Attribute VB_Name = "GetWorkbooksSheets"
Option Explicit
'Project Name    : �t�@�C���_�C�A���O�őI�������t�H���_�̒����̃G�N�Z���t�@�C������̃��[�N�u�b�N�ɂ܂Ƃ߂܂�
'File Name       : GetWorkbooksSheets.bas
'Feature         : Error�͈�ؖ����̗͋Z�X�^�C���ł��I�I�V�K�ō��Workbooks�̌��ɉ����Ă����̂Ō��t�@�C�����󂷐S�z���Ȃ��I(^��^)
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

    '----�萔�E�ϐ��ꗗ
    Const cnsDir     As String = "\*.xls*" 'Excel�̊g���q
    Dim dirCol       As New Collection    '�t�@�C���̃t���p�X����x�S�Ċi�[����R���N�V����
    Dim strPathName  As String  'File��Path
    Dim strFileName  As String  'File��Name
    Const sheetName  As String = "contents" '�쐬�����V�K�t�@�C���ɂ͕\����t���ăf�[�^�̃p�X�ƃt�@�C�����A�V�[�g�����ꗗ������
    Dim wb           As Workbook '�I�u�W�F�N�g�ϐ� - �V�K�ō쐬���郏�[�N�u�b�N
    Dim ws           As Worksheet '�I�u�W�F�N�g�ϐ� - ���[�N�V�[�g�̃C�e���[�^
    Dim wc           As Long '���[�N�V�[�g�̐��A����v�Z����̂ł͂Ȃ��A�C���N�������g�`���� + 1�����Ă���
    Dim getCol       As New Collection '�\���p�̃R���N�V���� arr�Ƃ����ϐ��œ񎟌��z��ɂ܂Ƃ߂�
    Dim c            As Variant
    Dim colLine      As Variant: colLine = Array("FilePath", "FileName", "SheetName")   '�R���N�V�����p�̃C�e���[�^
    Dim i            As Long    '�z��p �ꎟ���C���f�b�N�X
    Dim j            As Long    '�z��p �񎟌��C���f�b�N�X
    
    '1. �t�@�C���_�C�A���O���J���Ĉꗗ������t�H���_��I������
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case .Show
            Case True: strPathName = .SelectedItems(1)
            Case False: Exit Sub
        End Select
    End With

    '2. �g���q�ɊY������t�@�C�������݂��Ȃ��t�H���_�ł͍쓮���܂���
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

    ' * !!!�G���[�͈�ؖ����̗͋Z�X�^�C��!!!
    On Error Resume Next
   
    ' - Workbook�ɕt�т���G���[���b�Z�[�W�͑S�Ė�������
    ' - �����Ƀ}�N���������̂��߁AScreenUpdating��EnableEvents���N�����̊�OFF�ɂ���
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' - �V�KWorkbooks�̈�ԍŌ�̃V�[�g�ɒǉ�����@Ctrl+N�̋Z�ƈꏏ�ł�
    ' - Sheet1�͒ǉ����Ă������t�@�C���̏ڍׂ������ڎ��ɂ��܂�
    Set wb = Workbooks.Add
    wb.Sheets(1).Name = sheetName
    
    ' * getCol�̓t�@�C���t���p�X�A�t�@�C�����A�V�[�g����~�ς���R���N�V�����ł�
    ' - �E���グ���u�b�N�̏ڍׂ������V�[�g�̂��߂ɃR���N�V�����Ńf�[�^��~�ς���
    ' - �w�b�_�[�̃Z�b�g---colLine
    ' - wc��Excel�̃o�[�W�����΍�ł�
    
    getCol.Add colLine
    wc = Worksheets.Count
    
    '5. Copy�̔���
    ' - Item1�̓w�b�_�[�Ȃ̂ōŏ��̃C���f�b�N�X��2
    For i = 2 To dirCol.Count
        
        ' - �z�� [ FullPath FileName SheetName ]
        colLine = Array(dirCol(i)(0), Dir(dirCol(i)(0)), "SheetName")
        
        ' - Excel�u�b�N���J��
        Workbooks.Open Filename:=colLine(0)
        
        ' - Excel�u�b�N�̃��[�N�V�[�g�̐�����Copy���J��Ԃ�
        For Each ws In Workbooks(colLine(1)).Worksheets
            ' - colLine��SheetName����
            colLine(2) = ws.Name
            
            ' - �V�K���[�N�u�b�N�̍Ō�ɃR�s�[
            ws.Copy after:=wb.Sheets(wc)
            ' - �C���N�������g
            wc = wc + 1
            
            ' - �R���N�V�����ɔz���ǉ�
            getCol.Add colLine
        Next ws
        
        ' - �L�������킳�������I�I
        Workbooks(colLine(1)).Close
        
    Next i
    
    ' �G���[�����͂����ŉ�������
    Application.DisplayAlerts = True

    '5. �R���N�V������񎟌��z��ɕς��܂�
    ReDim arr(getCol.Count, LBound(getCol(1)) To UBound(getCol(1)))
    
    '6. Range�ɓ��͂����ňꎟ���z����񎟌��z���Lbound��0�ɂ��Ă��܂��B
    ''�e��̗v�f�͑S��2�ɂȂ�܂��B�w�b�_�[�̐��Ńo�C���h����܂�
    i = 0
    For Each c In getCol
        For j = LBound(c) To UBound(c)
            arr(i, j) = c(j)
        Next j
        i = i + 1
    Next c
    
    '7. �V�[�g�I�u�W�F�N�g�̎w����O���Ă���̂ŃA�N�e�B�u�Z�����N�Z�ɂȂ�܂��iAttention�F���ӁI�I�j
    With wb.Sheets(sheetName)
        .Range(.Cells(1, 1), .Cells(UBound(arr, 1) + 1, UBound(arr, 2) + 1)) = arr
        .Range(.Cells(1, 1), .Cells(UBound(arr, 1) + 1, UBound(arr, 2) + 1)).Columns.AutoFit
        .Select
    End With
    
    '8. �I�t�ɂ��Ă����@�\��True��
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    '9. ��Ɗ���Msgbox
    MsgBox wb.Name & vbNewLine & "������ɑI�������t�H���_��Excel�u�b�N�̃V�[�g���܂Ƃ߂܂����I", vbInformation, "Success"
    
    ' �I�u�W�F�N�g�ϐ������
    Set wb = Nothing

End Sub
