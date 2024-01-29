Attribute VB_Name = "����ϯä��R"
Sub �ϯä��R_��Ʈw()
DATA_ALL_ROWS = Sheets("Data").Range("H1").End(xlDown).Row

'�R���ª�÷s�W����
Windows(ThisWorkbook.Name).Activate
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("�ϯä��R").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "�ϯä��R"
    
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Data!R1C1:R" & DATA_ALL_ROWS & "C26", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="�ϯä��R!R1C1", TableName:="�ϯä��R��1", DefaultVersion:= _
        xlPivotTableVersion14
    Sheets("�ϯä��R").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("DESCRIPTION")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SIZE1")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SIZE2")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SIZE_TEXT")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("AREA")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("LINE NO")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SHEET")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_THK")
        .Orientation = xlRowField
        .Position = 8
    End With
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_THK").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_TYPE")
        .Orientation = xlRowField
        .Position = 8
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_THK")
        .Orientation = xlRowField
        .Position = 9
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("CLASS")
        .Orientation = xlRowField
        .Position = 10
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("PAINT CODE")
        .Orientation = xlRowField
        .Position = 11
    End With
    With ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("REMARK")
        .Orientation = xlRowField
        .Position = 12
    End With
    ActiveSheet.PivotTables("�ϯä��R��1").AddDataField ActiveSheet.PivotTables("�ϯä��R��1" _
        ).PivotFields("QTY"), "�[�` - QTY", xlSum
    ActiveSheet.PivotTables("�ϯä��R��1").RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables("�ϯä��R��1")
        .ColumnGrand = False
        .RowGrand = False
    End With

    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("REMARK").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("PAINT CODE").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("CLASS").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_THK").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("INSU_TYPE").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("PAINT CODE").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SHEET").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("LINE NO").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("AREA").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SIZE1").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("SIZE2").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("�ϯä��R��1").PivotFields("DESCRIPTION").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)

End Sub
