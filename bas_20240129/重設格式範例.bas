Attribute VB_Name = "重設格式範例"
Sub resetForm_資料庫()
Dim mainSN As String
Dim dataStartRow As Integer, dataEndRow As Integer

mainSN = "CvKv計算表"
dataStartRow = 2
dataEndRow = 500

Application.ScreenUpdating = False

    Sheets(mainSN).Select


'設回預設值======================
'跨欄置中
    Range("A:I").Select
    Selection.UnMerge
'下拉式選單
    Cells.Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
'內容============================


'字顏色
    '程式自動填上的欄
    Range("A:J").Select
    With Selection.Font
        .Color = -4165632
        .TintAndShade = 0
    End With
    '使用者手動輸入的欄
    Range("B:G").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    '使用者使用下拉式選單輸入者
    Range("C:C,E:E,G:G").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    

    
'字體
    
    '內文
   Range("A:I").Select
    With Selection.Font
        .Name = "新細明體"
        .Size = 12
    End With
    '標題列
    Rows("1:1").Select
    With Selection.Font
        .Name = "標楷體"
        .Size = 12
    End With
'字對齊
    '置中
    Range("A:A,C:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    '靠左
    Range("B:B,E:E,G:G").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    '靠右
    Range("D:D,F:F,H:I").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    '換行與縮小
    Cells.Select
    With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
'欄寬列高自動調整
    Cells.EntireColumn.AutoFit
    
'欄寬
    Columns("A:D").ColumnWidth = 24
    Cells.Select

    
 '框線
    '全部
    Range("A:J").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    '特殊處理
    Columns("D:E").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("F:G").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'下拉式選單
    Range("C:C").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="15A,20A,25A,32A,40A,50A,65A,80A,100A,125A"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "錯誤"
        .InputMessage = ""
        .ErrorMessage = "輸入值必須與下拉式選單相同"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("E:E").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="GPM,LPM,LPS,m3/hr"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "錯誤"
        .InputMessage = ""
        .ErrorMessage = "輸入值必須與下拉式選單相同"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("G:G").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="psi,ft-W.G.,M-W.G.,kPa,kg/cm2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "錯誤"
        .InputMessage = ""
        .ErrorMessage = "輸入值必須與下拉式選單相同"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    
'標題============================

'字內容
    i = 1
    Cells(1, i) = "項次"
    i = i + 1
    Cells(1, i) = "TAG NAME"
    i = i + 1
    Cells(1, i) = "管路尺寸"
    i = i + 1
    Cells(1, i) = "Q流量"
    i = i + 1
    Cells(1, i) = ""
    i = i + 1
    Cells(1, i) = "△P壓差"
    i = i + 1
    Cells(1, i) = ""
    i = i + 1
    Cells(1, i) = "Cv"
    i = i + 1
    Cells(1, i) = "Kv"
    i = i + 1
    Cells(1, i) = "<提醒>"
'跨欄置中
    Range("D1:E1").Select
    Selection.Merge
    Range("F1:G1").Select
    Selection.Merge
'字對齊
    '置中
    Range("A1:J1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
'字顏色
    Rows("1:1").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
'儲存格填滿
    Range("A1:I1").Select
    With Selection.Interior
        .Color = 5296274
    End With
    Range("J1:J1").Select
    With Selection.Interior
        .Color = 65535
    End With
    
'下拉式選單(清除)
    Range("A1:J1").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
Application.ScreenUpdating = True
End Sub

