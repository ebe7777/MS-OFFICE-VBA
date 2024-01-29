Attribute VB_Name = "關於Autocad開檔關檔存檔"
Sub 開關檔_資料庫()

Dim oldFilePath As String, oldFileName As String, newFilePath As String, newFileName As String
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
    
    Set acad = GetObject(, "AutoCAD.Application")
    '既有路徑與檔名
    oldFileName = Cells(1, 3)
    oldFilePath = Cells(1, 2)
    filePathName = oldFilePath & oldFileName
    newFilePath = "c:\"
    newFileName = "test"
    '開檔-可編輯
    Set dwgFile = acad.Documents.Open(filePathName, False)
    ' 如果autocad檔案的sdi設為1，代表一次只能開一個檔案，開另一檔代表關此檔；但我們要開關開關，所以要設為0
    sdi = dwgFile.GetVariable("sdi")
    dwgFile.SetVariable "sdi", 0
    '轉檔
    With dwgFile
      .SendCommand "-PURGE" & vbCr & "A" & vbCr & "*" & vbCr & "N" & vbCr
      .SendCommand "-PURGE" & vbCr & "R" & vbCr & "*" & vbCr & "N" & vbCr
      acad.ZoomExtents
      .SaveAs newFilePath & "\" & newFileName & ".dwg", ac2004_dwg
    End With
    '關檔-不存檔
    dwgFile.Close False
    
End Sub

