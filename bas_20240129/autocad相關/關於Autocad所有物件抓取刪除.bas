Attribute VB_Name = "����Autocad�Ҧ��������R��"

Private Sub getBlockRefAttributeValue(blockName As String, attributeName As String, returnAttributeVal As String)
Dim acad As AcadApplication, dwg As AcadDocument
Dim obj As Object
Dim allAttributes As Variant
Dim i As Integer
 
    Set acad = GetObject(, "AutoCAD.Application")
    With acad
        If .Documents.Count = 0 Then
          Set dwg = .Documents.Add
        Else
          Set dwg = .Documents(0)
        End If
    End With
    '���o�Ҧ�[�϶��Ѧ�]���󤤪�
    For Each obj In dwg.ModelSpace
        If TypeOf obj Is AcadBlockReference Then
            If (obj.Name = blockName) Then
                '��[�ݩ�] (���ϥΪ̥i�H�b�϶�����J�ϸ��B�ϦW�Bñ�W...��)
                allAttributes = obj.GetAttributes
                For i = 0 To UBound(allAttributes, 1)
                    If allAttributes(i).TagString = attributeName Then
                        returnAttributeVal = allAttributes(i).TextString
                        Exit For
                    End If
                Next i
                '��[�ʺA�϶����]�w��]
                allAttributes = obj.GetDynamicBlockProperties
                Exit For
            End If
        End If
    Next
    '���ocad�ɤ��]�t���Ҧ�block(���צ��L�Φb�ϭ��W)
    Dim myBlock As AcadBlock
    For Each myBlock In dwg.Blocks
       'do something...
    Next
End Sub
