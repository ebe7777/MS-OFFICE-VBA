Attribute VB_Name = "����inputbox"
Sub INPUTBOX_��Ʈw() '
    'InputBox "Tell user what to do", "Title of window", "default value in input box"
    '�Ninput���ȳ]�a�J�ܼ�
    quotnSN = InputBox("�п�J�u�@��W��", "��J�T��", "�����")
    '�p�G�ϥΪ̮ר������}�{��
    If (quotnSN = "") Then
        Exit Sub
    End If
End Sub
