Attribute VB_Name = "B_TOOL"
Sub SATH()
'SATH = SAVE AS THIS SHEET

SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, FileFilter:="Excel檔案(*.xls),*.xls", Title:="另存新檔名稱")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).Copy
    ActiveWorkbook.SaveAs Filename:=SAVE_NAME, FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Windows(ActiveWorkbook.Name).Close
    
    Else
    MsgBox "操作取消並未存檔!"
End If
    
    
End Sub

Sub GET_SHEET()
UserForm1.Show 0
Unload ProgressBar
End Sub
