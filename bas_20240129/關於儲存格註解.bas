Attribute VB_Name = "�����x�s�����"
Sub �����x�s�����()
Attribute �����x�s�����.VB_ProcData.VB_Invoke_Func = " \n14"
'�s�W����
Range("A1").AddComment
'�R������
Cells.ClearComments
'�P�_���ѬO�_�s�b
If (Range("A1").Comment Is Nothing) Then
    
End If
'��ܵ���
Range("A1").Comment.Visible = True
'���Ѽg�J���e
Range("A1").Comment.Text Text:="Bruce Chen �����B:" & Chr(10) & "123" & Chr(10) & "456"
End Sub

Sub �N���Ѯ榡��m���]()
Dim myCell As Variant
    For Each myCell In Cells.SpecialCells(xlCellTypeComments)
        myCell.Comment.Shape.Placement = 1
        myCell.Comment.Shape.Top = myCell.Top
        myCell.Comment.Shape.Left = myCell.Left + myCell.Width
        myCell.Comment.Shape.TextFrame.AutoSize = True
        myCell.Comment.Visible = False
    Next
End Sub

Sub �P�_���Ѧs���s�b()

    If ActiveSheet.Cells(1, 1).Comment Is Nothing Then
        MsgBox "not exist"
    Else
        MsgBox "exist"
    End If

End Sub
