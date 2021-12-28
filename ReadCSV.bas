Attribute VB_Name = "ReadCSV"
Option Explicit
Sub ReadCSV()
    Dim filePath As String: filePath = "****"
    
    Dim csvData As Object
    With Sheet1
        Set csvData = .QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=.Range("A1"))
    End With
    
    With csvData
        .TextFileCommaDelimiter = True
        .TextFileParseType = xlDelimited
        .TextFileStartRow = 1
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 932
        .Refresh
        .Delete
    End With
End Sub
