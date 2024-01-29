Attribute VB_Name = "關於Chart的操作"
Sub 基本_資料庫()
Dim myChartName As String
Dim myChart As Chart
    '新增
    myChartName = "Sales"
    Charts.Add after:=Sheets(Sheets.Count)
    ActiveChart.Name = myChartName
    Set myChart = Charts(myChartName)
    '移動
    Charts(myChartName).Move after:=Sheets(Sheets.Count)
    'for each
    For Each iVar In Charts
        If (iVar.Name = "Sales") Then
            MsgBox "YES"
        End If
    Next
    
    '編輯內容
    '   https://learn.microsoft.com/zh-tw/office/vba/api/excel.chart(object)
    '   https://learn.microsoft.com/en-us/office/vba/api/excel.chart(object)
    '設定資料範圍
    '   以折線圖為例，資料每個數值是由左至右展示在圖表上
    '       xlRows 將範圍內的資料以Row為單位視為一組並由左至右依序取出，然後由左至右展示在圖表上
    '           會將範圍內的資料由上到下，將無法顯示在圖表上的都當作水平項次標題
    '       xColumnss 將範圍內的資料以Column為單位視為一組並由上至下依序取出，然後由左至右展示在圖表上
    '           會將範圍內的資料由左到右，將無法顯示在圖表上的都當作水平項次標題
    myChart.SetSourceData Source:=Sheets("TEST1").Range("D2010:BM2011"), PlotBy:=xlRows
    myChart.SetSourceData Source:=Sheets("TEST1").Range("H2010:H2011"), PlotBy:=xlColumns
    '水平/垂直軸的大標題，譬如 "日期"，而非每筆資料的標題( "1/1","2/1"...)
    myChart.ChartWizard CategoryTitle:="水平軸大標題"
    myChart.ChartWizard ValueTitle:="垂直軸大標題"
    '設定的資料範圍，第幾組當作標題(0:沒有標題，每組都是資料；2:第1~2組的都是標題，第3組開始為資料)
    myChart.ChartWizard CategoryLabels:=0
    '是否只顯示看的見的儲存格 (false 才會將隱藏的儲存格的值秀出，不論是CategoryLabels或SeriesLabels)
    myChart.PlotVisibleOnly = False
    '水平/垂直座標的值控制
    myChart.Axes(xlValue).MaximumScale = 1
    
    myChart.ChartType = xlLine
    myChart.ClearToMatchStyle
    myChart.ChartStyle = 23
    
    
    '刪除
    Charts(myChartName).Delete
End Sub


