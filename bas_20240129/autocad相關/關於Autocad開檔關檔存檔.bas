Attribute VB_Name = "����Autocad�}�����ɦs��"
Sub �}����_��Ʈw()

Dim oldFilePath As String, oldFileName As String, newFilePath As String, newFileName As String
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
    
    Set acad = GetObject(, "AutoCAD.Application")
    '�J�����|�P�ɦW
    oldFileName = Cells(1, 3)
    oldFilePath = Cells(1, 2)
    filePathName = oldFilePath & oldFileName
    newFilePath = "c:\"
    newFileName = "test"
    '�}��-�i�s��
    Set dwgFile = acad.Documents.Open(filePathName, False)
    ' �p�Gautocad�ɮת�sdi�]��1�A�N��@���u��}�@���ɮסA�}�t�@�ɥN�������ɡF���ڭ̭n�}���}���A�ҥH�n�]��0
    sdi = dwgFile.GetVariable("sdi")
    dwgFile.SetVariable "sdi", 0
    '����
    With dwgFile
      .SendCommand "-PURGE" & vbCr & "A" & vbCr & "*" & vbCr & "N" & vbCr
      .SendCommand "-PURGE" & vbCr & "R" & vbCr & "*" & vbCr & "N" & vbCr
      acad.ZoomExtents
      .SaveAs newFilePath & "\" & newFileName & ".dwg", ac2004_dwg
    End With
    '����-���s��
    dwgFile.Close False
    
End Sub

