Sub xlsxSeparationScript()
    Dim OpenFiles As Variant
    '複数選択可能のダイアログボックスを開く
    OpenFiles = Application.GetOpenFilename("Microsoft Excelブック,*.xlsx", MultiSelect:=True)
    If IsArray(OpenFiles) = False Then Exit Sub
    'ファイルの出力先を選ぶ
    Dim strDirectory As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            strDirectory = .SelectedItems(1)
        End If
    End With
    If strDirectory = "" Then Exit Sub
    '=============================
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    '=============================
    Dim targetBookName As String, ws As Worksheet, i As Integer, j As Integer
    Const MAX_Column As Integer = 4
    Dim Listlinks As Variant
    'カツオのコメント_力技でださい==
    Dim FileName As Variant: ReDim FileName(1 To 1) '二次元配列 1列目 ファイル名
    Dim SheetName As Variant: ReDim SheetName(1 To 1) '二次元配列 1列目 シート名
    Dim LinkCheck As Variant: ReDim LinkCheck(1 To 1) '二次元配列 1列目 外部リンク
    Dim LinkDetail As Variant: ReDim LinkDetail(1 To 1) '二次元配列 1列目 外部リンク詳細
    '===============================
    Dim objFileSys As Object
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    'ヘッダーの設定
    FileName(1) = "ファイル名"
    SheetName(1) = "シート名"
    LinkCheck(1) = "外部リンク"
    LinkDetail(1) = "外部リンク詳細"
    '配列の要素数追加_カツオのコメント_ここダサイので修正したい==
    Dim ListNum As Integer: ListNum = 2
    ReDim Preserve FileName(1 To ListNum)
    ReDim Preserve SheetName(1 To ListNum)
    ReDim Preserve LinkCheck(1 To ListNum)
    ReDim Preserve LinkDetail(1 To ListNum)
    '============================================================
    Dim SheetIndex As Integer
    'Passwordファイルは不可
    On Error GoTo PasswordError
    '1.実際にファイル分け分け
    Application.DisplayAlerts = False
    For i = LBound(OpenFiles) To UBound(OpenFiles)
        Workbooks.Open FileName:=OpenFiles(i), Password:=vbNullString, UpdateLinks:=False
        targetBookName = Dir(OpenFiles(i))
        With Workbooks(targetBookName)
            For Each ws In .Worksheets
                SheetIndex = ws.Index
                ws.Copy
                Listlinks = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks) '開いたブックの中に外部リンクがあるかチェック
                If IsArray(Listlinks) Then
                    For j = 1 To UBound(Listlinks)
                        ActiveWorkbook.BreakLink Listlinks(j), xlLinkTypeExcelLinks 'リンクの解除
                    Next j
                    FileName(ListNum) = targetBookName
                    SheetName(ListNum) = ws.Name
                    LinkCheck(ListNum) = True
                    LinkDetail(ListNum) = Join(Listlinks, ",")
                Else
                    FileName(ListNum) = targetBookName
                    SheetName(ListNum) = ws.Name
                    LinkCheck(ListNum) = False
                    LinkDetail(ListNum) = Null
                End If
                '配列の要素数追加_カツオのコメント_ここダサイので修正したい==
                ListNum = ListNum + 1
                ReDim Preserve FileName(1 To ListNum)
                ReDim Preserve SheetName(1 To ListNum)
                ReDim Preserve LinkCheck(1 To ListNum)
                ReDim Preserve LinkDetail(1 To ListNum)
                '============================================================
                ActiveWorkbook.SaveAs FileName:=strDirectory & "\" & objFileSys.GetBaseName(OpenFiles(i)) & "_Sheet" & SheetIndex & "_" & ws.Name & ".xlsx"
                ActiveWorkbook.Close
            Next ws
            .Close
        End With
    Next i
    Application.DisplayAlerts = True
    ' end 1
    '2.新しいブック出して履歴を出力
    Dim Two_ListArray As Variant
    ReDim Two_ListArray(1 To UBound(FileName), 1 To MAX_Column)
    For i = 1 To UBound(FileName)
        Two_ListArray(i, 1) = FileName(i)
        Two_ListArray(i, 2) = SheetName(i)
        Two_ListArray(i, 3) = LinkCheck(i)
        Two_ListArray(i, 4) = LinkDetail(i)
    Next i
    Workbooks.Add
    With ActiveWorkbook
        With .Sheets(1)
            .Range(.Cells(1, 1), .Cells(UBound(Two_ListArray, 1), MAX_Column)) = Two_ListArray
        End With
        .SaveAs FileName:=strDirectory & "\" & Format(Now, "yyyymmddhhmmss") & "separationDetail.xlsx"
    End With
    Set objFileSys = Nothing
    ' end 2
    '=============================
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    '=============================
    MsgBox "Complete", vbInformation, "Split"
    Exit Sub
PasswordError:
    MsgBox Dir(OpenFiles(i)) & vbNewLine & "パスワード付ファイルのため、作業を継続できませんでした。"
    Set objFileSys = Nothing
    '=============================
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    '=============================
End Sub
