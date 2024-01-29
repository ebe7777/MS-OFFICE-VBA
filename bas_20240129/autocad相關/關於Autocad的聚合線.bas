Attribute VB_Name = "關於Autocad的聚合線"
Sub Example_PolarPoint_劃一條線()
    ' This example finds the coordinate of a point that is a given
    ' distance and angle from a base point.
    
    Dim polarPnt As Variant
    Dim basePnt(0 To 2) As Double
    Dim angle As Double
    Dim distance As Double
    
    basePnt(0) = 2#: basePnt(1) = 2#: basePnt(2) = 0#
    'angle = 用radians表示的角度
    '   算法: radians = degree  * pi  / 180
    '         degree  = radians * 180 / pi
    angle = 0.1744444   ' 45 degrees <--按照算式45degree應該=0.785 radians,待測試
    distance = 5
    '以第一點為基準指定下一點
    polarPnt = thisDrawing.Utility.PolarPoint(basePnt, angle, distance)
    
    ' Create a line from the base point to the polar point
    Dim lineObj As AcadLine
    Set lineObj = thisDrawing.ModelSpace.AddLine(basePnt, polarPnt)
    ZoomAll
    
End Sub
Private Function drawDwgRegularTriangle_畫正三角形_資料庫(topPointX As Double, topPointY As Double)
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
    'angle = 用radians表示的角度
    '   算法: radians = degree  * pi  / 180
    '         degree  = radians * 180 / pi
    '         45 degrees = 0.785
    '         60 degrees = 1.046
    angle = 1.046
    distance = 10.8558
    '以第一點為基準指定下一點
    '   angle的實際方向 :0是3點鐘方向,90是12點鐘方向,180是9點鐘方向,270是6點鐘方向 ; 逆時鐘是正數,順時鐘是負數
    '   此例中,basePnt是在上方 ,polarPnt1在右下,polarPnt2在左下
    polarPnt1 = dwg.Utility.PolarPoint(basePnt, angle * 5, distance)
    polarPnt2 = dwg.Utility.PolarPoint(polarPnt1, angle * 3, distance)
    
    ' Create a line from the base point to the polar point
    Set lineObj = dwg.ModelSpace.AddLine(basePnt, polarPnt1)
    Set lineObj = dwg.ModelSpace.AddLine(polarPnt1, polarPnt2)
    Set lineObj = dwg.ModelSpace.AddLine(polarPnt2, basePnt)
End Function
