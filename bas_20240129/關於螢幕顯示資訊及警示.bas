Attribute VB_Name = "����ù���ܸ�T��ĵ��"
Sub ��Ʈw()
'�����ù���ܧ�s
Application.ScreenUpdating = False

'����ĵ�i�ܵ����
Application.DisplayAlerts = False

'����events����-�b�u�@��ϥ�event�ɥi�ϥΦ��קK�L���j��
Application.EnableEvents = False

'�]�t�ק�Listbox��rawsource��Ū�����u�@��ɭPĲ�olistbox
'   �����۰ʭp��
Application.Calculation = xlCalculationManual
'   �}�Ҧ۰ʭp��
Application.Calculation = xlCalculationAutomatic

'excel�����ؤo
'   ����
Application.Visible = False
Application.Visible = True
'   �̤j��
Application.WindowState = xlMaximized
'   �ثe��
myHeig = Application.Height
'   �ثe�e
myWid = Application.Width
'   �������W���a�ù���m�F�`�N�A�������i�b�̤j�ƪ����A
Application.WindowState = xlNormal
Application.Top = 1
Application.Left = 1

'focus��excel
AppActivate Application.Caption

'�Ȱ�
Application.Wait Now + TimeValue("00:00:05")

'���T�w�A�����ӬO�����i���i�H�ާ@
Application.Interactive = False '  �T��椬�Ҧ� :

'���T�w
Application.StatusBar = False '�������A�C

'excel������ �۩w�q���u�@��t�Ʈ�,�j�������u�@��A�\��P���UCTRL+ALT+F9
'   https://docs.microsoft.com/en-us/office/troubleshoot/excel/custom-function-calculate-wrong-value
Application.Volatile
End Sub
