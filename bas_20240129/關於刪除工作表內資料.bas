Attribute VB_Name = "關於刪除工作表內資料"
Sub 清空範圍內儲存格資料_資料庫()
'只是刪除資料並非刪除欄列
Sheets("SYSTEM").Range("A:G").ClearContents
End Sub

Public Function clearRange_資料庫(myRange As Range)
'清除儲存格 內容、底色、字色
    With myRange
        .Formula = ""
        '.ClearComments
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
End Function
