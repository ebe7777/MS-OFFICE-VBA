Attribute VB_Name = "����Autocad��text����"
Private Sub Acad��text�Ѧ�_��Ʈw()
Dim textObj As AcadText
Dim textString As String
Dim insertionPoint(0 To 2) As Double, alignmentPoint(0 To 2) As Double
Dim height As Double
    ' Define the new Text object
    '��r>���e
    textString = "Hello, World."
    '�X��ϧ�>��m X Y Z
    insertionPoint(0) = 3: insertionPoint(1) = 3: insertionPoint(2) = 0
    '��r>��r��� X Y Z
    alignmentPoint(0) = 3: alignmentPoint(1) = 3: alignmentPoint(2) = 0
    '��r>����
    height = 0.5
    ' Create the Text object in model space
    Set textObj = thisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
    ' Set the text alignment to a value other than acAlignmentLeft, which is the default.
    ' Create a point that will act as an alignment reference point
    '��r>�勵
    textObj.Alignment = acAlignmentRight
    ' Create the text alignment reference point and the text will automatically
    ' align to the right of this point, because the text
    ' alignment was set to acAlignmentRight
    textObj.TextAlignmentPoint = alignmentPoint
    thisDrawing.Regen acActiveViewport
    '��r>�Φ�
    textObj.StyleName = "HTX1"
    textObj.Update

End Sub
Private Sub autocad�����r_getDwgTextValue(selectedText As String, p1X As Double, p1Y As Double, p2X As Double, p2Y As Double, selectionSetName As String)
Dim acad As AcadApplication, dwg As AcadDocument
Dim selectionSet As AcadSelectionSet
Dim p1(0 To 2) As Double, p2(0 To 2) As Double
Dim gpCode(0) As Integer, dataValue(0) As Variant
Dim txt As Variant

    '�M��selectedText�J�����-�קK�����줣��ȤϦӧ�ǤJ���­ȦA�Ǧ^
    selectedText = ""
    
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
Private Function �إߤ�r_drawDwgText(insertionPointX As Double, insertionPointY As Double, alignmentPointX As Double, alignmentPointY As Double, textHeight As Double, alignmenType As Integer, textStyle As String, textString As String)
'alignmenType��J���O�@��integer�A�W�ٻP�Ʀr��Ӧp�U
'acAlignmentLeft 0
'acAlignmentCenter 1
'acAlignmentRight 2
'acAlignmentAligned 3
'acAlignmentMiddle 4
'acAlignmentFit 5
'acAlignmentTopLeft 6
'acAlignmentTopCenter 7
'acAlignmentTopRight 8
'acAlignmentMiddleLeft 9
'acAlignmentMiddleCenter 10
'acAlignmentMiddleRight 11
'acAlignmentBottomLeft 12
'acAlignmentBottomCenter 13
'acAlignmentBottomRight 14
Dim acad As AcadApplication, dwg As AcadDocument
Dim textObj As AcadText
Dim insertionPoint(0 To 2) As Double, alignmentPoint(0 To 2) As Double

    
    Set acad = GetObject(, "AutoCAD.Application")
    With acad
        If .Documents.Count = 0 Then
          Set dwg = .Documents.Add
        Else
          Set dwg = .Documents(0)
        End If
    End With
    
    ' Define the new Text object
    '��r>���e
    '�X��ϧ�>��m X Y Z
    insertionPoint(0) = insertionPointX: insertionPoint(1) = insertionPointY: insertionPoint(2) = 0
    '��r>��r��� X Y Z
    alignmentPoint(0) = alignmentPointX: alignmentPoint(1) = alignmentPointY: alignmentPoint(2) = 0
    
    Set textObj = dwg.ModelSpace.AddText(textString, insertionPoint, textHeight)
    '��r>�勵
    'textObj.Alignment = acAlignmentRight
    textObj.Alignment = alignmenType
    textObj.TextAlignmentPoint = alignmentPoint

    '��r>�Φ�
    textObj.StyleName = textStyle

End Function
