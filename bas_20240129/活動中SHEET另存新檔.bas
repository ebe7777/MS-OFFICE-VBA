Attribute VB_Name = "���ʤ�SHEET�t�s�s��"
Sub SATH()
'SATH = SAVE AS THIS SHEET

SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, FileFilter:="Excel2007�榡�ɮ�(*.xlsx),*.xlsx", Title:="�t�s�s�ɦW��")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).Copy
    ActiveWorkbook.SaveAs Filename:=SAVE_NAME, FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Windows(ActiveWorkbook.Name).Close
    
    Else
    MsgBox "�ާ@�����å��s��!"
End If
    
    
End Sub

Sub �u�@��t�s�s��FUNCTION()

Dim mySheet As Worksheet
Dim savePath As String

savePath = Sheets("SYSTEM").Range("B4")
For Each mySheet In Sheets
    If mySheet.Name = "PROJECT" Or Right(mySheet.Name, 6) = "SYSTEM" Then
        '�N�u�@����s���ɮ�
        Sheets(mySheet.Name).Copy
        '�S�O�B�z�APROJECT�NH��R��
        If mySheet.Name = "PROJECT" Then
            Columns("H:H").ClearContents
        End If
        '�ɮץO�s�s��
        Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs Filename:=savePath & "\BOM_SUP_" & mySheet.Name _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            Windows(ActiveWorkbook.Name).Close
        Application.DisplayAlerts = True
    End If
Next

End Sub
