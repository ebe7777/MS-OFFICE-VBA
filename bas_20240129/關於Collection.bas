Attribute VB_Name = "����Collection"
'   https://www.codevba.com/help/collection.htm#.ZAWmkXZBx9U



Sub ��Ʈw_collection()
'�w�q�ɻݭn�ϥ�new,���D��collection�w�g�s�b
Dim iColllection As New Collection
Dim iBool As Boolean
    '�s�W value,key(����string)
    iColllection.Add "Value", "Key"
    '�O�_�s�b
    On Error Resume Next
        iBool = iColllection.Item(Key)
        If (iBool <> Empty) Then
            Exists = True
        End If
    On Error GoTo 0
    '�M��
    Set iColllection = Nothing
End Sub

