Attribute VB_Name = "����ErrorHandle"
'�s����D:VBA��J����~�B�ϥΪ̦��UOn Error�ɡA�|�i�J�ϥΪ̦ۭq�����~�B�z�Ҧ��A�åB�����~�|���B�z�����|�A�׶i�JOn Error
'  �|��: ��for��P "�ϥΪ̫��w�W��"�P�W���u�@���B�z�A�èϥ�on error�Ө���䤣��P�W���u�@��
'           ��Ĥ@���J����D�ɡA�|�i�Jon error���w���B�z�覡�A���A�׹J��ɴN�|���X�t�ο��~�T���ӫD�i�Jon error
'��]:�bon error��ϥΪ̶��n�ϥΥH�U���@�Ӥ覡�ӧi��VBA�����~�w�g�B�z����
'   Resume Next ' go to the line following error
'   Resume ' go back to the same line of code
'   Exit Sub ' go out of this routine
Sub ���TerrorHandle�d��()
    For i = 1 To 10
        haveErr = False
        '���b-�Ӥu�@���s�b
        On Error GoTo 880
        '�i��|���~�B - �p�G�u�@���s�b�i�J880�аO�����D
        Set nowSht = Sheets(Cells(i, 1))
        '�Non error�令���B�z-�_�h���U�ӹJ�쪺������~���|�Hgoto 880�B�z
        On Error GoTo 0
        '�S���D�~����ʧ@
        If (haveErr = False) Then
           'when no error,do something
        End If
        GoTo 881
880
        '�аO�����D
        haveErr = True
        '�i�D�{����誺���~�w�g�B�z����
        Resume Next
881
    Next i
End Sub
Sub ����991�P999_��Ʈw()
Dim msgTitle As String, msgText As String, msgStyle As String

991
    msgTitle = "ĵ�i            "   ' �w�q���D�C
    msgText = "  �Э��s���榹�{�� !"
    MsgBox msgText, vbExclamation, msgTitle
    Exit Sub

999
'�i��user���槹��
    msgTitle = "�T��            "    ' �w�q���D�C
    msgText = "  �{�����槹�� !"
    MsgBox msgText, vbOKOnly, msgTitle
    

End Sub

