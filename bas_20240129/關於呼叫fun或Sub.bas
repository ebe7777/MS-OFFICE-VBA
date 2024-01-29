Attribute VB_Name = "關於呼叫fun或Sub"
Sub 關於呼叫_資料庫()
    
    'sub要呼叫另一個Module的程式
    Call Module1.loadPubVar
    
    '呼叫另一個工作表內的程式
    Call ThisWorkbook.Worksheets("工作表1").loadPubVar

End Sub
