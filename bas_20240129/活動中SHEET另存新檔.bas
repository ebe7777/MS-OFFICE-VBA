Attribute VB_Name = "活動中SHEET另存新檔"
Sub SATH()
'SATH = SAVE AS THIS SHEET

SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, FileFilter:="Excel2007格式檔案(*.xlsx),*.xlsx", Title:="另存新檔名稱")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).Copy
    ActiveWorkbook.SaveAs Filename:=SAVE_NAME, FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
        CreateBackup:=False
    Windows(ActiveWorkbook.Name).Close
    
    Else
    MsgBox "操作取消並未存檔!"
End If
    
    
End Sub

Sub 工作表另存新檔FUNCTION()

Dim mySheet As Worksheet
Dim savePath As String

savePath = Sheets("SYSTEM").Range("B4")
For Each mySheet In Sheets
    If mySheet.Name = "PROJECT" Or Right(mySheet.Name, 6) = "SYSTEM" Then
        '將工作表放到新的檔案
        Sheets(mySheet.Name).Copy
        '特別處理，PROJECT將H欄刪除
        If mySheet.Name = "PROJECT" Then
            Columns("H:H").ClearContents
        End If
        '檔案令存新檔
        Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs Filename:=savePath & "\BOM_SUP_" & mySheet.Name _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            Windows(ActiveWorkbook.Name).Close
        Application.DisplayAlerts = True
    End If
Next

End Sub
