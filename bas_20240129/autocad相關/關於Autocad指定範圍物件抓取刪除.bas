Attribute VB_Name = "關於Autocad指定範圍物件抓取刪除"
Private Sub getTextValue(selectedText As String, p1X As Double, p1Y As Double, p2X As Double, p2Y As Double, selectionSetName As String)
Dim acad As AcadApplication, dwg As AcadDocument
Dim selectionSet As AcadSelectionSet
Dim p1(0 To 2) As Double, p2(0 To 2) As Double
Dim gpCode(0) As Integer, dataValue(0) As Variant
Dim i As Integer, txt As Variant, s As String
Dim totalRows As Integer
Dim sysSN As String

Set acad = GetObject(, "AutoCAD.Application")
With acad
    If .Documents.Count = 0 Then
      Set dwg = .Documents.Add
    Else
      Set dwg = .Documents(0)
    End If
End With
'抓CAD相關資訊
'p1是抓範圍的起點
p1(0) = p1X: p1(1) = p1Y: p1(2) = 0#
'p2是抓範圍的終點
p2(0) = p2X: p2(1) = p2Y: p2(2) = 0#
gpCode(0) = 0: dataValue(0) = "TEXT"
'如果不讓畫面縮小到指定位置，可能會抓錯資料(譬如縮太小導致精度不足)
acad.ZoomWindow p1, p2
Set selectionSet = dwg.SelectionSets.Add(selectionSetName)
'物件 全部 在p1 p2 範圍內才選取 - p1在左下,p2在右上
'selectionSet.Select acSelectionSetWindow, p1, p2, gpCode, dataValue
'物件 部分 在p1 p2 範圍內就選取 - p1在右上,p2在左下
selectionSet.Select acSelectionSetCrossing, p1, p2, gpCode, dataValue

For Each txt In selectionSet
    selectedText = Trim(txt.textString)
Next

End Sub
Private Function deleteDwgObject(objectName As String, p1X As Double, p1Y As Double, p2X As Double, p2Y As Double, selectionSetName As String)
Dim acad As AcadApplication, dwg As AcadDocument
Dim selectionSet As AcadSelectionSet
Dim p1(0 To 2) As Double, p2(0 To 2) As Double
Dim gpCode(0) As Integer, dataValue(0) As Variant
Dim deletedObj As Variant


    Set acad = GetObject(, "AutoCAD.Application")
    With acad
        If .Documents.Count = 0 Then
          Set dwg = .Documents.Add
        Else
          Set dwg = .Documents(0)
        End If
    End With
    '抓BCAD相關資訊
    'p1是抓範圍的起點
    p1(0) = p1X: p1(1) = p1Y: p1(2) = 0#
    'p2是抓範圍的終點
    p2(0) = p2X: p2(1) = p2Y: p2(2) = 0#
    gpCode(0) = 0: dataValue(0) = objectName
    '如果不讓畫面縮小到指定位置，可能會抓錯資料(譬如縮太小導致精度不足)
    acad.ZoomWindow p1, p2
    Set selectionSet = dwg.SelectionSets.Add(selectionSetName)
    '物件 全部 在p1 p2 範圍內才選取 - p1在左下,p2在右上
    'selectionSet.Select acSelectionSetWindow, p1, p2, gpCode, dataValue
    '物件 部分 在p1 p2 範圍內就選取 - p1在右上,p2在左下
    selectionSet.Select acSelectionSetCrossing, p1, p2, gpCode, dataValue
    
    
    For Each deletedObj In selectionSet
''for test:know which text is deleted
'Dim TEST As String
'TEST = Trim(deletedObj.textString)
        deletedObj.Delete
    Next
End Function
