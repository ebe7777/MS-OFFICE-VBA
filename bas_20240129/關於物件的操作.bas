Attribute VB_Name = "���󪫥󪺾ާ@"
Sub �ק�֨�������x�s��s��()
Dim iVar As CheckBox
Dim iCol As Long, iRow As Long

For Each iVar In ActiveSheet.CheckBoxes
    '���o�֨�����Ҧb���(����֨������A��P�򭭩w�L��������I�n�����b���x�s�檺��C���ɤ�)
    iCol = iVar.BottomRightCell.Column
    iRow = iVar.BottomRightCell.Row
    '�]�w�֨�������x�s��s��
    iVar.LinkedCell = Cells(iRow, iCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
Next
MsgBox "���槹��"
End Sub
