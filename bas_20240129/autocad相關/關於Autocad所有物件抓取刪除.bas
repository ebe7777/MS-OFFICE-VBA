Attribute VB_Name = "關於Autocad所有物件抓取刪除"

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
    '取得所有[圖塊參考]物件中的
    For Each obj In dwg.ModelSpace
        If TypeOf obj Is AcadBlockReference Then
            If (obj.Name = blockName) Then
                '抓[屬性] (讓使用者可以在圖塊中輸入圖號、圖名、簽名...等)
                allAttributes = obj.GetAttributes
                For i = 0 To UBound(allAttributes, 1)
                    If allAttributes(i).TagString = attributeName Then
                        returnAttributeVal = allAttributes(i).TextString
                        Exit For
                    End If
                Next i
                '抓[動態圖塊的設定值]
                allAttributes = obj.GetDynamicBlockProperties
                Exit For
            End If
        End If
    Next
    '取得cad檔中包含的所有block(不論有無用在圖面上)
    Dim myBlock As AcadBlock
    For Each myBlock In dwg.Blocks
       'do something...
    Next
End Sub
