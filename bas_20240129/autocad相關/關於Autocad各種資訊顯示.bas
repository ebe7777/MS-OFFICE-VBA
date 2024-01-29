Attribute VB_Name = "����Autocad�U�ظ�T���"
Sub �`��()
    Dim acad As AcadApplication
    Set acad = GetObject(, "AutoCAD.Application")
    
    '����vba������autocad()
    acad.Visible = True
    
    '������m()
    '   �̤j��
    acad.WindowState = acMax
    '   ���o�ثe�������ؤo
    i = acad.Application.Height
    ii = acad.Application.Width
    '   �]�w�����������W���Z���ù����W����m ;�`�N�A���i�b�̤j�ƪ����p�U�ϥ�
    acad.WindowState = acNorm
    acad.WindowTop = 1
    acad.WindowLeft = (ii / 2) - 1
    '   �]�w�����ؤo
    acad.Height = i
    acad.Width = ii / 2
    '   windows foucus�����
    AppActivate acad.Caption
End Sub
Sub �}�ɮɷ|�J�쪺ĵ�i�T��()
Dim acad As AcadApplication, dwgFile As AcadDocument
Dim preferences As AcadPreferences, currShowProxyDialogBox As Boolean
    Set acad = GetObject(, "AutoCAD.Application")

'===�]custom objects���ͪ�ĵ�i�T��
    Set preferences = acad.Application.preferences
    '�]�m �}�ɮɤ��n��ܦ]custom objects���ͪ�ĵ�i�T��
    '   Retrieve the current ShowProxyDialogBox value
    currShowProxyDialogBox = preferences.OpenSave.ShowProxyDialogBox
    '   Change the value for ShowProxyDialogBox
    preferences.OpenSave.ShowProxyDialogBox = Not (currShowProxyDialogBox)

    '�}��autocad
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    '   do something...
    
'===�]��AEC����(if a drawing has AEC object references)
'   �L�k�HVBA�B�z�B�z
'   ��������
'       https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Error-opening-a-drawing-This-application-has-detected-a-mixed-version-of-AEC-objects.html
'       https://forums.autodesk.com/t5/net/how-to-disable-aec-warning-message/td-p/4703443

'===�ɮת������s(This drawing contains objects from a newer version..)
'   �d�LVBA�O�_�i�B�z��T
'   ��������
'       https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Error-This-drawing-contains-objects-from-a-newer-version-when-opening-drawing-in-AutoCAD.html
End Sub
