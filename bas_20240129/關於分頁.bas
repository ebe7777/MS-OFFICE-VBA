Attribute VB_Name = "�������"
'�]�w�����u��p�G�C�L�ɨt�Τ��۰ʧ��ܤ����u�A�h�O�n�ˬd�ӦC�L�覡�O�_���t�~���� "�Y����",Ĵ�pADOBE�|�۰ʭp���Y���Ҩýվ�EXCEL�����u

Sub �[�J���������u_��Ʈw()
'�p�G�O���̳\�hSHEET���j��,�n�b�C�iSHEET���ͮɱN�����u���X�k1
HPageBreaks_NUM = HPageBreaks_NUM + 1
'�[�J�����u
'�|�bCLASS_ROW���C�W��[�W�����u
Set ActiveSheet.HPageBreaks(HPageBreaks_NUM).Location = Range("A" & CLASS_ROW)
HPageBreaks_NUM = HPageBreaks_NUM + 1
End Sub
Sub �]�w�C�L�d��_��Ʈw()
'�]�w�C�L�d��
Sheets(LAST_CLASS_NAME).PageSetup.PrintArea = "A1:AY" & LAST_ROW
End Sub
