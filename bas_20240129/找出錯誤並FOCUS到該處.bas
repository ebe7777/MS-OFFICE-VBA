Attribute VB_Name = "��X���~��FOCUS��ӳB"
Sub �����~_��Ʈw()
  If WorksheetFunction.CountIf(Sheets("MTO4FPCC").Range("D2:D" & TOTAL_ROW2), "����S���R�W") > 0 Then
  MsgBox "������|���R�W�A�гq��PIPING�H���ק�"
  ADD_WRONG = Sheets("MTO4FPCC").Range("D2:D" & TOTAL_ROW2).Find("����S���R�W").Address
  Range(ADD_WRONG).Select

End Sub

