Attribute VB_Name = "������oWindowsUserName"
Sub ��Ʈw_���owindowsUserName()
a = getWindowsUserName
MsgBox a
End Sub
Function getWindowsUserName()
'���o�ϥΪ̦W��
getWindowsUserName = VBA.Interaction.Environ$("UserName")
'f���o�ϥΪ̸�Ƨ�
Environ ("Userprofile")
End Function
Sub getWindowsUserName��ƨӷ�()
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
