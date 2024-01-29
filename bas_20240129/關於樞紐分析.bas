Attribute VB_Name = "關於樞紐分析"
Sub 樞紐分析_資料庫()
DATA_ALL_ROWS = Sheets("Data").Range("H1").End(xlDown).Row

'刪除舊表並新增此表
Windows(ThisWorkbook.Name).Activate
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("樞紐分析").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "樞紐分析"
    
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Data!R1C1:R" & DATA_ALL_ROWS & "C26", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="樞紐分析!R1C1", TableName:="樞紐分析表1", DefaultVersion:= _
        xlPivotTableVersion14
    Sheets("樞紐分析").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("DESCRIPTION")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SIZE1")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SIZE2")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SIZE_TEXT")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("AREA")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("LINE NO")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SHEET")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_THK")
        .Orientation = xlRowField
        .Position = 8
    End With
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_THK").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_TYPE")
        .Orientation = xlRowField
        .Position = 8
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_THK")
        .Orientation = xlRowField
        .Position = 9
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("CLASS")
        .Orientation = xlRowField
        .Position = 10
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("PAINT CODE")
        .Orientation = xlRowField
        .Position = 11
    End With
    With ActiveSheet.PivotTables("樞紐分析表1").PivotFields("REMARK")
        .Orientation = xlRowField
        .Position = 12
    End With
    ActiveSheet.PivotTables("樞紐分析表1").AddDataField ActiveSheet.PivotTables("樞紐分析表1" _
        ).PivotFields("QTY"), "加總 - QTY", xlSum
    ActiveSheet.PivotTables("樞紐分析表1").RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables("樞紐分析表1")
        .ColumnGrand = False
        .RowGrand = False
    End With

    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("REMARK").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("PAINT CODE").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("CLASS").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_THK").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("INSU_TYPE").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("PAINT CODE").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SHEET").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("LINE NO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("AREA").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SIZE1").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("SIZE2").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("樞紐分析表1").PivotFields("DESCRIPTION").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)

End Sub
