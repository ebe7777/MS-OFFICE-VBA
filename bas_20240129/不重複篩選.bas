Attribute VB_Name = "�����ƿz��"
Sub �����ƿz��_��Ʈw()
'��ƦbD��A�ѤW���U�ˬd�A�u�n������ƩM�W�誺�@�˴N���L�A���@�˴N�a�^L��
'�ݥ��T�O�ۦP���(D��)�O�W�U���p�b�@�_��
N = 0
For Each EBE In Range("D1:D5")
    If EBE.Value <> EBE.Offset(-1, 0) Then
        Range("L" & 2 + N) = EBE.Value
        N = N + 1
    End If
Next
End Sub
