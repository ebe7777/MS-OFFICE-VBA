Attribute VB_Name = "����FUNCTION�]�i���쪺��"
Sub �[�`�p��_��Ʈw()
'*�`�N! �g�L�����ҩ��A���覡��X�Ӫ��ȷ|���p���I�H�U�n�X�쪺�s�P�Ʀr�X�{�A�зV��!
'EBE1�O�����,EBE2�O��ƭ�
'��̱���ۦP�h�NEBE2��ƪ�QTY�[�`�bSUM_M
'�NSUM_M�a�^EBE1��ƪ�
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

Sub �R�������ŦX���󪺾�C���_��Ʈw()

'�q�D����(MAIN)��P���M�ŦX�ȵ���"O"��,�p�ŦX�ӦC�R��
ALL_MAIN_ROWS = Worksheets(MAIN_SHEET).Range("H1").End(xlDown).Row

For Each EBE In Sheets(MAIN_SHEET).Range("P2:P" & ALL_MAIN_ROWS)
    If EBE.Value = "O" Then
        Rows(EBE.Row).Select
        Selection.Delete Shift:=xlUp
    End If
Next
End Sub

Sub �ϥ�MID��X�r��()
MTO_ALLROWS = Worksheets("MTO_��z").Range("A1").End(xlDown).Row
On Error Resume Next
For CLASS_ROW = 2 To MTO_ALLROWS
    Range("L" & CLASS_ROW).Value = Mid(Range("A" & CLASS_ROW), 2, InStr(2, Range("A" & CLASS_ROW), "/", vbTextCompare) - 2)
Next
On Error GoTo 0
End Sub
