Attribute VB_Name = "關於凍結視窗"
Sub 關於凍結視窗_資料庫()
Attribute 關於凍結視窗_資料庫.VB_ProcData.VB_Invoke_Func = " \n14"
ActiveWindow.FreezePanes = False
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 0
End With
'機此此儲存格上面&左邊凍結
Range("B3").Select
ActiveWindow.FreezePanes = True
End Sub
