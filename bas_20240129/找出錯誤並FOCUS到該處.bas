Attribute VB_Name = "找出錯誤並FOCUS到該處"
Sub 找到錯誤_資料庫()
  If WorksheetFunction.CountIf(Sheets("MTO4FPCC").Range("D2:D" & TOTAL_ROW2), "儀表沒有命名") > 0 Then
  MsgBox "有儀表尚未命名，請通知PIPING人員修改"
  ADD_WRONG = Sheets("MTO4FPCC").Range("D2:D" & TOTAL_ROW2).Find("儀表沒有命名").Address
  Range(ADD_WRONG).Select

End Sub

