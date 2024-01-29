Attribute VB_Name = "關於取得WindowsUserName"
Sub 資料庫_取得windowsUserName()
a = getWindowsUserName
MsgBox a
End Sub
Function getWindowsUserName()
'取得使用者名稱
getWindowsUserName = VBA.Interaction.Environ$("UserName")
'f取得使用者資料夾
Environ ("Userprofile")
End Function
Sub getWindowsUserName資料來源()
'https://officetricks.com/excel-vba-get-username-windows-system/
Dim idx As Integer
'To Directly the value of a Environment Variable with its Name
MsgBox VBA.Interaction.Environ$("UserName")

'To get all the List of Environment Variables
For idx = 1 To 255
    strEnvironVal = VBA.Interaction.Environ$(idx)
    ThisWorkbook.Sheets(1).Cells(idx, 1) = strEnvironVal
Next idx

End Sub
