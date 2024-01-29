Attribute VB_Name = "����Autocad���w�d�򪫥����R��"
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
'��CAD������T
'p1�O��d�򪺰_�I
p1(0) = p1X: p1(1) = p1Y: p1(2) = 0#
'p2�O��d�򪺲��I
p2(0) = p2X: p2(1) = p2Y: p2(2) = 0#
gpCode(0) = 0: dataValue(0) = "TEXT"
'�p�G�����e���Y�p����w��m�A�i��|������(Ĵ�p�Y�Ӥp�ɭP��פ���)
acad.ZoomWindow p1, p2
Set selectionSet = dwg.SelectionSets.Add(selectionSetName)
'���� ���� �bp1 p2 �d�򤺤~��� - p1�b���U,p2�b�k�W
'selectionSet.Select acSelectionSetWindow, p1, p2, gpCode, dataValue
'���� ���� �bp1 p2 �d�򤺴N��� - p1�b�k�W,p2�b���U
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
    '��BCAD������T
    'p1�O��d�򪺰_�I
    p1(0) = p1X: p1(1) = p1Y: p1(2) = 0#
    'p2�O��d�򪺲��I
    p2(0) = p2X: p2(1) = p2Y: p2(2) = 0#
    gpCode(0) = 0: dataValue(0) = objectName
    '�p�G�����e���Y�p����w��m�A�i��|������(Ĵ�p�Y�Ӥp�ɭP��פ���)
    acad.ZoomWindow p1, p2
    Set selectionSet = dwg.SelectionSets.Add(selectionSetName)
    '���� ���� �bp1 p2 �d�򤺤~��� - p1�b���U,p2�b�k�W
    'selectionSet.Select acSelectionSetWindow, p1, p2, gpCode, dataValue
    '���� ���� �bp1 p2 �d�򤺴N��� - p1�b�k�W,p2�b���U
    selectionSet.Select acSelectionSetCrossing, p1, p2, gpCode, dataValue
    
    
    For Each deletedObj In selectionSet
''for test:know which text is deleted
'Dim TEST As String
'TEST = Trim(deletedObj.textString)
        deletedObj.Delete
    Next
End Function
