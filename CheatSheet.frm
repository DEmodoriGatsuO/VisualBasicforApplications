VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheatSheet 
   Caption         =   "CheatSheet"
   ClientHeight    =   11535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   OleObjectBlob   =   "CheatSheet.frx":0000
   StartUpPosition =   3  'Windows の既定値
End
Attribute VB_Name = "CheatSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cmd1_Click()
    Unload Me
End Sub
Private Sub Mails_Click()
    Dim buf As String, CB As New DataObject
    buf = Mails.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L1_Click()
    Dim buf As String, CB As New DataObject
    buf = L1.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L2_Click()
    Dim buf As String, CB As New DataObject
    buf = L2.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L3_Click()
    Dim buf As String, CB As New DataObject
    buf = L3.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L4_Click()
    Dim buf As String, CB As New DataObject
    buf = L4.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L5_Click()
    Dim buf As String, CB As New DataObject
    buf = L5.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L6_Click()
    Dim buf As String, CB As New DataObject
    buf = L6.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L7_Click()
    Dim buf As String, CB As New DataObject
    buf = L7.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L8_Click()
    Dim buf As String, CB As New DataObject
    buf = L8.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L9_Click()
    Dim buf As String, CB As New DataObject
    buf = L9.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L10_Click()
    Dim buf As String, CB As New DataObject
    buf = L10.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L11_Click()
    Dim buf As String, CB As New DataObject
    buf = L11.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L12_Click()
    Dim buf As String, CB As New DataObject
    buf = L12.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L13_Click()
    Dim buf As String, CB As New DataObject
    buf = L13.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L14_Click()
    Dim buf As String, CB As New DataObject
    buf = L14.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L15_Click()
    Dim buf As String, CB As New DataObject
    buf = L15.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L16_Click()
    Dim buf As String, CB As New DataObject
    buf = L16.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L17_Click()
    Dim buf As String, CB As New DataObject
    buf = L17.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L18_Click()
    Dim buf As String, CB As New DataObject
    buf = L18.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L19_Click()
    Dim buf As String, CB As New DataObject
    buf = L19.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L20_Click()
    Dim buf As String, CB As New DataObject
    buf = L20.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L21_Click()
    Dim buf As String, CB As New DataObject
    buf = L21.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L22_Click()
    Dim buf As String, CB As New DataObject
    buf = L22.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L23_Click()
    Dim buf As String, CB As New DataObject
    buf = L23.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L24_Click()
    Dim buf As String, CB As New DataObject
    buf = L24.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L25_Click()
    Dim buf As String, CB As New DataObject
    buf = L25.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L26_Click()
    Dim buf As String, CB As New DataObject
    buf = L26.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L27_Click()
    Dim buf As String, CB As New DataObject
    buf = L27.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L28_Click()
    Dim buf As String, CB As New DataObject
    buf = L28.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub
Private Sub L29_Click()
    Dim buf As String, CB As New DataObject
    buf = L29.Caption
    With CB
        .SetText buf        ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
    End With
    
    MsgBox "選択ラベルのデータをクリップボードに保存しました。" & vbNewLine & buf, vbInformation, "メッセージ"
End Sub


