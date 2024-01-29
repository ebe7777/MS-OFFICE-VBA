Attribute VB_Name = "計算出每ROW的列高"
Sub 計算出列高並輸出到另一張表_資料庫()
ANSWER = Sheets("SHEET1").Range("A1").End(xlDown).Row
For A = 1 To ANSWER
Sheets("SHEET2").Range("A" & A).Value = Sheets("SHEET1").Rows(A).RowHeight
Next
End Sub
