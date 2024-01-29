Attribute VB_Name = "刪除所有自訂樣式"

Sub 刪除所有自訂樣式()
'[常用] > [樣式] > [新增儲存格樣式]
Dim iVar As Variant
    For Each iVar In ThisWorkbook.Styles
        If (iVar.BuiltIn = False) Then
            iVar.Delete
        End If
    Next
    
End Sub

