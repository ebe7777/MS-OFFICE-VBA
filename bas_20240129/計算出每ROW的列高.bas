Attribute VB_Name = "�p��X�CROW���C��"
Sub �p��X�C���ÿ�X��t�@�i��_��Ʈw()
ANSWER = Sheets("SHEET1").Range("A1").End(xlDown).Row
For A = 1 To ANSWER
Sheets("SHEET2").Range("A" & A).Value = Sheets("SHEET1").Rows(A).RowHeight
Next
End Sub
