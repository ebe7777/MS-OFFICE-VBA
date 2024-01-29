Attribute VB_Name = "不用FUNCTION也可做到的事"
Sub 加總計算_資料庫()
'*注意! 經過實驗證明，此方式算出來的值會有小數點以下好幾位的零星數字出現，請慎用!
'EBE1是條件值,EBE2是資料值
'兩者條件相同則將EBE2資料表的QTY加總在SUM_M
'將SUM_M帶回EBE1資料表
For Each EBE1 In Sheets(DATA_SHEET).Range("L2:L" & ALL_DATAL_ROWS)
    SUM_M = 0
    For Each EBE2 In Sheets(MAIN_SHEET).Range("O2:O" & ALL_MAIN_ROWS)
        If EBE1.Value = EBE2.Value Then
            SUM_M = SUM_M + EBE2.Offset(0, 2).Value
        End If
    Next
    EBE1.Offset(0, 1).Value = SUM_M
Next

End Sub

Sub 刪除全部符合條件的整列資料_資料庫()

'從主表資料(MAIN)的P欄找尋符合值等於"O"者,如符合該列刪除
ALL_MAIN_ROWS = Worksheets(MAIN_SHEET).Range("H1").End(xlDown).Row

For Each EBE In Sheets(MAIN_SHEET).Range("P2:P" & ALL_MAIN_ROWS)
    If EBE.Value = "O" Then
        Rows(EBE.Row).Select
        Selection.Delete Shift:=xlUp
    End If
Next
End Sub

Sub 使用MID抓出字串()
MTO_ALLROWS = Worksheets("MTO_整理").Range("A1").End(xlDown).Row
On Error Resume Next
For CLASS_ROW = 2 To MTO_ALLROWS
    Range("L" & CLASS_ROW).Value = Mid(Range("A" & CLASS_ROW), 2, InStr(2, Range("A" & CLASS_ROW), "/", vbTextCompare) - 2)
Next
On Error GoTo 0
End Sub
