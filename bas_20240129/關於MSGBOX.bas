Attribute VB_Name = "����MSGBOX"
Sub MSGBOX���e_��Ʈw()
Dim msgTitle As String, msgText As String, msgStyle As String

'�}����
    msgTitle = "�����~����            "    ' �w�q���D�C

    msgText = "�����p�U:" + vbLf   ' �w�q�T���C
    msgText = msgText + "�o�{������SUP TYPE�P�u�@�� [" & sNType & "] ���ꪺ��Ƥ��šC" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�а���H�U�ʧ@�G" + vbLf   ' �w�q�T��
    msgText = msgText + "(1)�ѷӤu�@�� [" & sNPdmsReport & "] ��L��PQ��Х� ERROR �B�íק��Ʋz�C" + vbLf  ' �w�q�T��
    msgText = msgText + "(2)�A�����榹�{���C"    ' �w�q�T��
        
'    msgStyle = vbOKOnly '��ܽT�w
'    msgStyle = vbOKCancel '��ܽT�w/����
'    msgStyle = vbYesNo '��ܬO/�_
'    msgStyle = vbYesNoCancel '��ܬO/�_/����
'    msgStyle = vbCritical '���"X"�Ϯ�
'    msgStyle = vbQuestion '���"?"�Ϯ�
'    msgStyle = vbExclamation '���"!"�Ϯ�
'    msgStyle = vbInformation '���"i"�Ϯ�
    
    MsgBox msgText, msgStyle, msgTitle
End Sub
Sub ���槹��()
    msgTitle = "�T��" ' �w�q���D�C
    msgText = "���槹��"
    msgStyle = vbInformation '���"i"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
End Sub
Sub VBOKCANCEL�g�k_��Ʈw()
Dim msgTitle As String, msgText As String, msgStyle As String, answer As Variant

msgTitle = "���n�T��            "    ' �w�q���D�C
msgText = "�ϥΥ��{�����H�U���󭭨� :" + vbLf + vbCrLf + vbLf  ' �w�q�T���C
msgText = msgText + "   1.SPEC�����ŦXPM�榡" + vbLf + vbCrLf  ' �w�q�T���C
msgText = msgText + "     (�ѷӥ��ɤ����u�@��""SPEC�榡�d��"",A~K�楲���P���ۦP)" + vbLf + vbCrLf  ' �w�q�T��
msgText = msgText + "   2.SPEC���ۦPCLASS���ޥ�ݱƧǦb�@�_" + vbLf + vbCrLf + vbLf  ' �w�q�T��
msgText = msgText + "�p�G�T�w�L�~�Ы�[ �T�w ]�~��A�Ϊ̫�[ ���� ]���}�{�� !" + vbLf + vbCrLf ' �w�q�T���C

answer = MsgBox(msgText, vbOKCancel + vbExclamation, msgTitle)

If answer = vbCancel Then
    Exit Sub
    Else
End If
End Sub

Sub VBYESNO�g�k_��Ʈw()
Dim msgTitle As String, msgText As String, msgStyle As String, answer As Variant

msgTitle = "���n�T��            "    ' �w�q���D�C
msgText = "�Ҳ��ͪ�MTO1~MTO4�`��O�_�n�C�L ? �p�G�n�C�L�п�� [�O] " + vbLf + vbCrLf + vbLf  ' �w�q�T���C
msgText = msgText + "   ==>�p���[�O]�A�{���|���F�C�L�i��ƪ��C" + vbLf + vbCrLf  ' �w�q�T���C"
msgText = msgText + "       *�ݯӮ� 30�� �H�W(��ƶq�V�h�V�[)" + vbLf + vbCrLf   ' �w�q�T��


answer = MsgBox(msgText, vbYesNo + vbQuestion, msgTitle)

If answer = vbYes Then
    
    Else
End If
End Sub
