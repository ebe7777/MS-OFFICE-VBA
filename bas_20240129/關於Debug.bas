Attribute VB_Name = "關於Debug"
Sub test()

myVal = 1

'將運算結果顯示在 "即時運算" 視窗
Debug.Print myVal
'設條件、檢查變數的值是否等於條件==>不等於時會pause在此列
Debug.Assert myVal = 2

End Sub
