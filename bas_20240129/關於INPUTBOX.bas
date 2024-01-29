Attribute VB_Name = "關於inputbox"
Sub INPUTBOX_資料庫() '
    'InputBox "Tell user what to do", "Title of window", "default value in input box"
    '將input的值設帶入變數
    quotnSN = InputBox("請輸入工作表名稱", "輸入訊息", "選機表")
    '如果使用者案取消離開程式
    If (quotnSN = "") Then
        Exit Sub
    End If
End Sub
