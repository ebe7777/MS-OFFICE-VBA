Attribute VB_Name = "�����x�s�檺�˦�Style"
Public Function ��Ʈw_deleteCustomStyle()
'�R���Ҧ��ۭq�˦�
'[�`��] > [�˦�] > [�s�W�x�s��˦�]
Dim iVar As Variant
    For Each iVar In ThisWorkbook.Styles
        If (iVar.BuiltIn = False) Then
            iVar.Delete
        End If
    Next
End Function
