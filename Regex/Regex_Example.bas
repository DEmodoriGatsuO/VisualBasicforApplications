Attribute VB_Name = "Regex_Example"
Option Explicit
'Microsoft VBScript Regular Expressions 5.5 = True
'===============================================
'正規表現を用いた成功事例 英数字単語の抜き出し
'===============================================
Sub Regex_Example()
    Dim i      As Integer
    Dim IRow   As Integer
    Dim iCount As Integer
    Dim myLine As Variant
    Dim mc     As MatchCollection
    Dim m      As Match
    Dim str    As String
    Dim re     As New RegExp

    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "[a-zA-Z0-9]+"
    End With

    With ActiveSheet
        IRow = .Cells(Rows.Count, 1).End(xlUp).Row
        Set mc = re.Execute(str)
        iCount = mc.Count
        
        For Each m In mc
            MsgBox m.Value
        Next
        Set mc = Nothing
    End With

End Sub
