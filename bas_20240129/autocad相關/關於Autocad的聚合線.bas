Attribute VB_Name = "����Autocad���E�X�u"
Sub Example_PolarPoint_���@���u()
    ' This example finds the coordinate of a point that is a given
    ' distance and angle from a base point.
    
    Dim polarPnt As Variant
    Dim basePnt(0 To 2) As Double
    Dim angle As Double
    Dim distance As Double
    
    basePnt(0) = 2#: basePnt(1) = 2#: basePnt(2) = 0#
    'angle = ��radians��ܪ�����
    '   ��k: radians = degree  * pi  / 180
    '         degree  = radians * 180 / pi
    angle = 0.1744444   ' 45 degrees <--���Ӻ⦡45degree����=0.785 radians,�ݴ���
    distance = 5
    '�H�Ĥ@�I����ǫ��w�U�@�I
    polarPnt = thisDrawing.Utility.PolarPoint(basePnt, angle, distance)
    
    ' Create a line from the base point to the polar point
    Dim lineObj As AcadLine
    Set lineObj = thisDrawing.ModelSpace.AddLine(basePnt, polarPnt)
    ZoomAll
    
End Sub
Private Function drawDwgRegularTriangle_�e���T����_��Ʈw(topPointX As Double, topPointY As Double)
Dim acad As AcadApplication, dwg As AcadDocument
Dim basePnt(0 To 2) As Double
Dim angle As Double
Dim distance As Double
Dim polarPnt1 As Variant, polarPnt2 As Variant
Dim lineObj As AcadLine
    
    Set acad = GetObject(, "AutoCAD.Application")
    With acad
        If .Documents.Count = 0 Then
          Set dwg = .Documents.Add
        Else
          Set dwg = .Documents(0)
        End If
    End With
    
    basePnt(0) = topPointX: basePnt(1) = topPointY: basePnt(2) = 0#
    'angle = ��radians��ܪ�����
    '   ��k: radians = degree  * pi  / 180
    '         degree  = radians * 180 / pi
    '         45 degrees = 0.785
    '         60 degrees = 1.046
    angle = 1.046
    distance = 10.8558
    '�H�Ĥ@�I����ǫ��w�U�@�I
    '   angle����ڤ�V :0�O3�I����V,90�O12�I����V,180�O9�I����V,270�O6�I����V ; �f�����O����,�������O�t��
    '   ���Ҥ�,basePnt�O�b�W�� ,polarPnt1�b�k�U,polarPnt2�b���U
    polarPnt1 = dwg.Utility.PolarPoint(basePnt, angle * 5, distance)
    polarPnt2 = dwg.Utility.PolarPoint(polarPnt1, angle * 3, distance)
    
    ' Create a line from the base point to the polar point
    Set lineObj = dwg.ModelSpace.AddLine(basePnt, polarPnt1)
    Set lineObj = dwg.ModelSpace.AddLine(polarPnt1, polarPnt2)
    Set lineObj = dwg.ModelSpace.AddLine(polarPnt2, basePnt)
End Function
