Attribute VB_Name = "關於儲存格註解"
Sub 關於儲存格註解()
Attribute 關於儲存格註解.VB_ProcData.VB_Invoke_Func = " \n14"
'新增註解
Range("A1").AddComment
'刪除註解
Cells.ClearComments
'判斷註解是否存在
If (Range("A1").Comment Is Nothing) Then
    
End If
'顯示註解
Range("A1").Comment.Visible = True
'註解寫入內容
Range("A1").Comment.Text Text:="Bruce Chen 陳彥錡:" & Chr(10) & "123" & Chr(10) & "456"
End Sub

Sub 將註解格式位置重設()
Dim myCell As Variant
    For Each myCell In Cells.SpecialCells(xlCellTypeComments)
        myCell.Comment.Shape.Placement = 1
        myCell.Comment.Shape.Top = myCell.Top
        myCell.Comment.Shape.Left = myCell.Left + myCell.Width
        myCell.Comment.Shape.TextFrame.AutoSize = True
        myCell.Comment.Visible = False
    Next
End Sub

Sub 判斷註解存不存在()

    If ActiveSheet.Cells(1, 1).Comment Is Nothing Then
        MsgBox "not exist"
    Else
        MsgBox "exist"
    End If

End Sub
