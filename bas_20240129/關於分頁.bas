Attribute VB_Name = "關於分頁"
'設定分頁線後如果列印時系統仍自動改變分頁線，則是要檢查該列印方式是否有另外控制 "縮放比例",譬如ADOBE會自動計算縮放比例並調整EXCEL分頁線

Sub 加入水平分頁線_資料庫()
'如果是楚裡許多SHEET的迴圈,要在每張SHEET產生時將分頁線號碼歸1
HPageBreaks_NUM = HPageBreaks_NUM + 1
'加入分頁線
'會在CLASS_ROW此列上方加上分頁線
Set ActiveSheet.HPageBreaks(HPageBreaks_NUM).Location = Range("A" & CLASS_ROW)
HPageBreaks_NUM = HPageBreaks_NUM + 1
End Sub
Sub 設定列印範圍_資料庫()
'設定列印範圍
Sheets(LAST_CLASS_NAME).PageSetup.PrintArea = "A1:AY" & LAST_ROW
End Sub
