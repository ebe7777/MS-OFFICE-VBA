Attribute VB_Name = "��sTAONO��"
 Sub ��sTAONO_��Ʈw()
 Application.ScreenUpdating = False
 '���o�{�b���ʤ���SHEET NAME
 ORIGIN_SHEET = ActiveSheet.Name
 '���oDATEABASE���ɦW
 DATE_PATH = Application.GetOpenFilename(FileFilter:="Excel�ɮ�(*.xls;*.xlsx),*.xls;*.xlsx", Title:="�}��TAONO��DataBase��..")
 DATE_NAME = Right(DATE_PATH, Len(DATE_PATH) - InStrRev(DATE_PATH, "\"))

 '�}��DATEBASE�ӽƻs���
 '==>�P�_DATABASE�ɮ׬O�_�w�g�}��
 For Each EBE In Workbooks
    If EBE.Name = DATE_NAME Then
        TEMP_A = 1
    End If
 Next
 
 If TEMP_A = 0 Then
     Workbooks.Open Filename:=DATE_PATH, ReadOnly:=True
 End If
'==>�ƻs��ƨ��л\
Windows(ThisWorkbook.Name).Activate
    Sheets("TAONO��").Select
    Cells.Select
    Selection.ClearContents
    Windows(DATE_NAME).Activate
    Sheets("TAONO��").Select
    Cells.Select
    Selection.Copy
    Windows(ThisWorkbook.Name).Activate
    Range("A1").Select
ActiveSheet.Paste

'����DATABASE
 If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(DATE_NAME).Close SaveChanges:=False
 End If
'==>�NFOCUS�a�^UPDATE�e��m
    Windows(ThisWorkbook.Name).Activate
    Sheets(ORIGIN_SHEET).Select
Application.DisplayAlerts = True
MsgBox "   ��s����!            ", vbInformation
 End Sub
