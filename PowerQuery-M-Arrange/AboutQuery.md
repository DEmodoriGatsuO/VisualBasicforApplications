# Book Query
## Power Query M formula language

 -  There is no folder.

#### Original Data

```html:Original Data
let
    FilePath = Excel.CurrentWorkbook(){[Name="FilePath"]}[Content]{0}[Column1] & "source.txt",
    Source = Table.FromColumns({Lines.FromBinary(File.Contents(FilePath), null, null, 65001)}),
    #"Add declare Column" = Table.AddColumn(Source, "declare", each if [Column1] = "let" then "let" else if [Column1] = "in" then "in" else null),
    #"Add For Split Column" = Table.AddColumn(#"Add declare Column", "function column", each if [declare] = null then [Column1] else null),
    #"Split Columns return value and call function" = Table.SplitColumn(#"Add For Split Column", "function column", Splitter.SplitTextByEachDelimiter({"="}, QuoteStyle.None, false), {"return value", "call function"}),
    #"Set Type return value and call function" = Table.TransformColumnTypes(#"Split Columns return value and call function",{{"return value", type text}, {"call function", type text}}),
    #"Select Columns" = Table.SelectColumns(#"Set Type return value and call function",{"declare", "return value", "call function"}),
    #"Trim Columns return value and call function" = Table.TransformColumns(#"Select Columns",{{"return value", Text.Trim, type text}, {"call function", Text.Trim, type text}}),
    #"Add index Column" = Table.AddIndexColumn(#"Trim Columns return value and call function", "index", 1, 1, Int64.Type)
in
    #"Add index Column"
```

#### Replacement
```html:Replacement
let
    #"Refer Original Data" = #"Original Data",
    #"Select Columns" = Table.SelectColumns(#"Refer Original Data",{"index" , "return value"}),
    #"Filter Replace Rows" = Table.SelectRows(#"Select Columns", each ([return value] <> null)),
    #"Remove LastRow" = Table.RemoveLastN(#"Filter Replace Rows",1),
    #"Add Input Column" = Table.AddColumn(#"Remove LastRow", "replace", each null),
    #"Rename return value" = Table.RenameColumns(#"Add Input Column",{{"return value", "pattern"}})
in
    #"Rename return value"
```