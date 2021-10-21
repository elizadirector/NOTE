Sub 巨集1()
'
' 巨集1 巨集

    'New sheet naming as date w format yyyymmdd
    Dim d
    d = Format(Date, "Long Date")
    d = Format(d, "yyyymmdd")
    Sheets.Add(After:=ActiveSheet).Name = d
    Range("A1").Select
    ActiveCell.FormulaR1C1 = d

    Range("A5").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A5:A6").Select
    Selection.AutoFill Destination:=Range("A5:A54"), Type:=xlFillDefault
    Range("A5:A54").Select
    
    '摩根大通
    'Application.CutCopyMode = False
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "摩根大通"

    ActiveWorkbook.Queries.Add Name:="Table 6", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=8440&b=8440""))," & Chr(13) & "" & Chr(10) & "    Data6 = 來源{6}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data6,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 6"";Extended Properties=""""" _
        , Destination:=Range("$B$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 6]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_6"
        .Refresh BackgroundQuery:=False
    End With
    
    'Application.CutCopyMode = False
    ActiveWorkbook.Queries.Add Name:="Table 7", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=8440&b=8440""))," & Chr(13) & "" & Chr(10) & "    Data7 = 來源{7}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data7,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 7"";Extended Properties=""""" _
        , Destination:=Range("$F$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 7]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_7"
        .Refresh BackgroundQuery:=False
    End With
    
    '美林
    'Application.CutCopyMode = False
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "美林"

    ActiveWorkbook.Queries.Add Name:="Table 6 (2)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=1440&b=1440""))," & Chr(13) & "" & Chr(10) & "    Data6 = 來源{6}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data6,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 6 (2)"";Extended Properties=""""" _
        , Destination:=Range("$J$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 6 (2)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_6__2"
        .Refresh BackgroundQuery:=False
    End With

    ActiveWorkbook.Queries.Add Name:="Table 7 (2)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""hhttps://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=1440&b=1440""))," & Chr(13) & "" & Chr(10) & "    Data7 = 來源{7}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data7,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 7 (2)"";Extended Properties=""""" _
        , Destination:=Range("$N$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 7 (2)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_7__2"
        .Refresh BackgroundQuery:=False
    End With

    '凱基台北
    'Application.CutCopyMode = False
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "凱基台北"
    
    ActiveWorkbook.Queries.Add Name:="Table 6 (3)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9268""))," & Chr(13) & "" & Chr(10) & "    Data6 = 來源{6}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data6,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 6 (3)"";Extended Properties=""""" _
        , Destination:=Range("$R$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 6 (3)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_6__3"
        .Refresh BackgroundQuery:=False
    End With
    
    'Application.CutCopyMode = False
    ActiveWorkbook.Queries.Add Name:="Table 7 (3)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9268""))," & Chr(13) & "" & Chr(10) & "    Data7 = 來源{7}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data7,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 7 (3)"";Extended Properties=""""" _
        , Destination:=Range("$V$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 7 (3)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_7__3"
        .Refresh BackgroundQuery:=False
    End With

    '凱基松山
    'Application.CutCopyMode = False
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "凱基松山"
    
    ActiveWorkbook.Queries.Add Name:="Table 6 (4)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9217""))," & Chr(13) & "" & Chr(10) & "    Data6 = 來源{6}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data6,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 6 (4)"";Extended Properties=""""" _
        , Destination:=Range("$Z$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 6 (4)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_6__4"
        .Refresh BackgroundQuery:=False
    End With

    'Application.CutCopyMode = False
    ActiveWorkbook.Queries.Add Name:="Table 7 (4)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9200&b=9217""))," & Chr(13) & "" & Chr(10) & "    Data7 = 來源{7}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data7,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 7 (4)"";Extended Properties=""""" _
        , Destination:=Range("$AB$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 7 (4)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_7__4"
        .Refresh BackgroundQuery:=False
    End With

    '富邦
    'Application.CutCopyMode = False
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "富邦"
    
    ActiveWorkbook.Queries.Add Name:="Table 6 (5)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9600&b=9600""))," & Chr(13) & "" & Chr(10) & "    Data6 = 來源{6}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data6,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 6 (5)"";Extended Properties=""""" _
        , Destination:=Range("$AH$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 6 (5)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_6__5"
        .Refresh BackgroundQuery:=False
    End With
    
    'Application.CutCopyMode = False
    ActiveWorkbook.Queries.Add Name:="Table 7 (5)", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    來源 = Web.Page(Web.Contents(""https://fubon-ebrokerdj.fbs.com.tw/z/zg/zgb/zgb0.djhtm?a=9600&b=9600""))," & Chr(13) & "" & Chr(10) & "    Data7 = 來源{7}[Data]," & Chr(13) & "" & Chr(10) & "    已變更類型 = Table.TransformColumnTypes(Data7,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    已變更類型"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 7 (5)"";Extended Properties=""""" _
        , Destination:=Range("$AL$3")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 7 (5)]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table_7__5"
        .Refresh BackgroundQuery:=False
    End With



    Selection.EntireRow.Hidden = True
End Sub

