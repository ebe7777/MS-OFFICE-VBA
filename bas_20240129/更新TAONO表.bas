Attribute VB_Name = "更新TAONO表"
 Sub 更新TAONO_資料庫()
 Application.ScreenUpdating = False
 '取得現在活動中的SHEET NAME
 ORIGIN_SHEET = ActiveSheet.Name
 '取得DATEABASE的檔名
 DATE_PATH = Application.GetOpenFilename(FileFilter:="Excel檔案(*.xls;*.xlsx),*.xls;*.xlsx", Title:="開啟TAONO表的DataBase檔..")
 DATE_NAME = Right(DATE_PATH, Len(DATE_PATH) - InStrRev(DATE_PATH, "\"))

 '開啟DATEBASE來複製資料
 '==>判斷DATABASE檔案是否已經開啟
 For Each EBE In Workbooks
    If EBE.Name = DATE_NAME Then
        TEMP_A = 1
    End If
 Next
 
 If TEMP_A = 0 Then
     Workbooks.Open Filename:=DATE_PATH, ReadOnly:=True
 End If
'==>複製資料並覆蓋
Windows(ThisWorkbook.Name).Activate
    Sheets("TAONO表").Select
    Cells.Select
    Selection.ClearContents
    Windows(DATE_NAME).Activate
    Sheets("TAONO表").Select
    Cells.Select
    Selection.Copy
    Windows(ThisWorkbook.Name).Activate
    Range("A1").Select
ActiveSheet.Paste

'關閉DATABASE
 If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(DATE_NAME).Close SaveChanges:=False
 End If
'==>將FOCUS帶回UPDATE前位置
    Windows(ThisWorkbook.Name).Activate
    Sheets(ORIGIN_SHEET).Select
Application.DisplayAlerts = True
MsgBox "   更新完畢!            ", vbInformation
 End Sub
