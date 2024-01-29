Attribute VB_Name = "不重複篩選"
Sub 不重複篩選_資料庫()
'資料在D欄，由上往下檢查，只要此筆資料和上方的一樣就跳過，不一樣就帶回L欄
'需先確保相同資料(D欄)是上下串聯在一起的
N = 0
For Each EBE In Range("D1:D5")
    If EBE.Value <> EBE.Offset(-1, 0) Then
        Range("L" & 2 + N) = EBE.Value
        N = N + 1
    End If
Next
End Sub
