Attribute VB_Name = "B_TOOL"
Sub SATH()
'SATH = SAVE AS THIS SHEET

SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, FileFilter:="Excel�ɮ�(*.xls),*.xls", Title:="�t�s�s�ɦW��")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).Copy
    ActiveWorkbook.SaveAs Filename:=SAVE_NAME, FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Windows(ActiveWorkbook.Name).Close
    
    Else
    MsgBox "�ާ@�����å��s��!"
End If
    
    
End Sub

Sub GET_SHEET()
UserForm1.Show 0
Unload ProgressBar
End Sub
