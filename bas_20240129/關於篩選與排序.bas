Attribute VB_Name = "關於篩選與排序"
Sub 啟動篩選_資料庫()
    ROWS("1:1").Select
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
        Selection.AutoFilter
        Else
        Selection.AutoFilter
    End If
End Sub

Sub 不解除篩選狀態下取消篩選_資料庫()
On Error Resume Next
Sheets("123").Select
ROWS("1:1").Select
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
End If
On Error GoTo 0
End Sub

Sub 自訂排序功能_資料庫()
Dim mySht As Worksheet
Dim customListOriginalCount As Long
Dim i As Long
'方式(1)使用一個或多個寫死的值 做排序依據
    mySht.Sort.SortFields.Add Key:=Range("A2:A10"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="汽車,機車", DataOption:=xlSortNormal
    '   執行排序
    With mySht.Sort
        .SetRange Range("A2:C10")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'方式(2)使用變動的值 做排序依據
'   原理 (1)將要排序的值 加入 自訂清單 (同 檔案>選項>進階>一般>編輯自訂排序)
'        (2)在排序方法內使用自訂清單做排序依據
'        (3)在自訂清單內刪除該值
    '計算原始自訂清單內有多少筆資料
    customListOriginalCount = Application.CustomListCount
    '將要加入 自訂清單 的值寫入Array
    sortOrderArray(1) = "汽車"
    sortOrderArray(2) = "機車"
    '新增 自訂清單
    Application.AddCustomList ListArray:=sortOrderArray
    '使用 自訂清單 內最後一筆資料(也就是上述新增的)做排序依據
    mySht.Sort.SortFields.Add Key:=Range("A2:A10"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=Application.CustomListCount, DataOption:=xlSortNormal
    '   執行排序
    With mySht.Sort
        .SetRange Range("A2:C10")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '刪除新增的CustomList
    '   當使用自訂清單功能後可能一存檔excel就當機，網路上說加上此行就不會(已經過測試證實)
    mySht.Sort.SortFields.Clear
    For i = Application.CustomListCount To customListOriginalCount + 1 Step -1
        Application.DeleteCustomList i
    Next i

End Sub
