Attribute VB_Name = "�R���Ҧ��ۭq�˦�"

Sub �R���Ҧ��ۭq�˦�()
'[�`��] > [�˦�] > [�s�W�x�s��˦�]
Dim iVar As Variant
    For Each iVar In ThisWorkbook.Styles
        If (iVar.BuiltIn = False) Then
            iVar.Delete
        End If
    Next
    
End Sub

