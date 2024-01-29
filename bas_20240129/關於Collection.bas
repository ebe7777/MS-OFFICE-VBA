Attribute VB_Name = "關於Collection"
'   https://www.codevba.com/help/collection.htm#.ZAWmkXZBx9U



Sub 資料庫_collection()
'定義時需要使用new,除非該collection已經存在
Dim iColllection As New Collection
Dim iBool As Boolean
    '新增 value,key(須為string)
    iColllection.Add "Value", "Key"
    '是否存在
    On Error Resume Next
        iBool = iColllection.Item(Key)
        If (iBool <> Empty) Then
            Exists = True
        End If
    On Error GoTo 0
    '清空
    Set iColllection = Nothing
End Sub

