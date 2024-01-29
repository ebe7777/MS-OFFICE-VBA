Attribute VB_Name = "A_關於加速vba運算"
Sub 加快執行速度的方式()
'避免在螢幕上顯示資料變動
'   停止螢幕更新
Application.ScreenUpdating = False
Application.ScreenUpdating = True
'   停止儲存格運算
Application.Calculation = xlCalculationManual
Application.Calculation = xlCalculationAutomatic
'   避免使用Application.StatusBar

'避免重複性的和工作表互動
'   將資料寫進變數(陣列)再做計算處理
    i = Range("A1").Value
    For ii = 1 To 1000000
        i = i + ii
    Next ii
'   讀取/寫入儲存格資料時使用陣列一次性執行
Dim myArray() As Variant
    myArray = Sheets("test").Range("A3:B4").Value
    '可在陣列中編輯資料
    myArray(1, 1) = 31
    Sheets("test").Range("A6:B7").Value = myArray

'自己寫程式替代Application.WorksheetFunction

'避免使用定義為Variants的變數

'避免判斷 文字 資料
'   舉例：if (myText = "abc") then
'       select case myText : case  "abc"
'   可將文字轉為數字來判斷，譬如Enum
Public Enum enumGender
    Male = 0
End Enum
Dim Gender As enumGender
    Select Case Gender
        Case Male
    End Select

'避免選取(.select)工作表後再取儲存格的值；應直接取用儲存格的值
    myValue = Worksheets("sheet1").Cells(1, 1).Value
    
'避免重複執行數學運算
    For i = 1 To 100
        '可以將(3 * 10 / 12)拆出先算好
        'myValue = myValue + (3 * 10 / 12)
        a = (3 * 10 / 12)
        myValue = myValue + a
    Next i
    
'不用使用會編輯儲存格既有方法，如.copy .paste .ClearContents
'   應操作儲存格的屬性，如 .value = XXX
End Sub
