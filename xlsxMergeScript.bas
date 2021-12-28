Sub xlsxMergeScript()
    Dim OpenFiles As Variant
    '複数選択可能のダイアログボックスを開く
    OpenFiles = Application.GetOpenFilename("Microsoft Excelブック,*.xlsx", MultiSelect:=True)
    If IsArray(OpenFiles) = False Then Exit Sub
    If UBound(OpenFiles) = 1 Then
        MsgBox "選択されたファイルが単一のファイルです。" & vbNewLine & OpenFiles(1) & _
            vbNewLine & "単一のファイルではスクリプトを実行しません。", vbCritical
        Exit Sub
    End If
    'Application～は決まり文句
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    Dim Mergebook As Workbook
    Dim wc As Integer
    '新しいブックを開く、何のシートを持ってきたのか表すのがList、データをひとまとめにしたのがMerge
    Workbooks.Add
    With ActiveSheet
        .Name = "List"
        .Copy After:=Sheets("List")
    End With
    ActiveSheet.Name = "Merge"
    Set Mergebook = ActiveWorkbook
    wc = Mergebook.Worksheets.Count
    '
    Dim targetBookName As String, ws As Worksheet, i As Integer
    Dim regionValues As Variant, MergeEndRow As Integer: MergeEndRow = 1
    Dim ListValues As Variant, ListRows As Integer: ListRows = 1
    ReDim ListValues(3)
    On Error GoTo PasswordError
    For i = LBound(OpenFiles) To UBound(OpenFiles)
        Workbooks.Open Filename:=OpenFiles(i), Password:=vbNullString
        targetBookName = Dir(OpenFiles(i))
        With Workbooks(targetBookName)
            ListValues(0) = OpenFiles(i)
            ListValues(1) = Dir(OpenFiles(i))
            For Each ws In .Worksheets
                ListValues(2) = ws.Name
                regionValues = ActiveSheet.Range("A1").CurrentRegion
                With Mergebook
                    With .Sheets("List")
                        .Range(.Cells(ListRows, 1), .Cells(ListRows, 3)) = ListValues
                        ListRows = ListRows + 1
                    End With
                    With .Sheets("Merge")
                        .Range(.Cells(MergeEndRow, 1), .Cells(MergeEndRow + UBound(regionValues, 1) - 1, UBound(regionValues, 2))) = regionValues
                        MergeEndRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
                    End With
                End With
                ws.Move After:=Mergebook.Sheets(wc)
                wc = wc + 1
            Next ws
        End With
    Next i
    Mergebook.Sheets("List").Select
    Set Mergebook = Nothing
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub
PasswordError:
    MsgBox Dir(OpenFiles(i)) & vbNewLine & "パスワード付ファイルのため、作業を継続できませんでした。"
    Application.DisplayAlerts = False
    Mergebook.Close
    Set Mergebook = Nothing
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub