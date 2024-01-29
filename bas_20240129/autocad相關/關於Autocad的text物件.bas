Attribute VB_Name = "關於Autocad的text物件"
Private Sub Acad建text參考_資料庫()
Dim textObj As AcadText
Dim textString As String
Dim insertionPoint(0 To 2) As Double, alignmentPoint(0 To 2) As Double
Dim height As Double
    ' Define the new Text object
    '文字>內容
    textString = "Hello, World."
    '幾何圖形>位置 X Y Z
    insertionPoint(0) = 3: insertionPoint(1) = 3: insertionPoint(2) = 0
    '文字>文字對齊 X Y Z
    alignmentPoint(0) = 3: alignmentPoint(1) = 3: alignmentPoint(2) = 0
    '文字>高度
    height = 0.5
    ' Create the Text object in model space
    Set textObj = thisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
    ' Set the text alignment to a value other than acAlignmentLeft, which is the default.
    ' Create a point that will act as an alignment reference point
    '文字>對正
    textObj.Alignment = acAlignmentRight
    ' Create the text alignment reference point and the text will automatically
    ' align to the right of this point, because the text
    ' alignment was set to acAlignmentRight
    textObj.TextAlignmentPoint = alignmentPoint
    thisDrawing.Regen acActiveViewport
    '文字>形式
    textObj.StyleName = "HTX1"
    textObj.Update

End Sub
Private Sub autocad抓取文字_getDwgTextValue(selectedText As String, p1X As Double, p1Y As Double, p2X As Double, p2Y As Double, selectionSetName As String)
Dim acad As AcadApplication, dwg As AcadDocument
Dim selectionSet As AcadSelectionSet
Dim p1(0 To 2) As Double, p2(0 To 2) As Double
Dim gpCode(0) As Integer, dataValue(0) As Variant
Dim txt As Variant

    '清除selectedText既有資料-避免此次抓不到值反而把傳入的舊值再傳回
    selectedText = ""
    
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
Private Function 建立文字_drawDwgText(insertionPointX As Double, insertionPointY As Double, alignmentPointX As Double, alignmentPointY As Double, textHeight As Double, alignmenType As Integer, textStyle As String, textString As String)
'alignmenType輸入的是一個integer，名稱與數字對照如下
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
    '文字>內容
    '幾何圖形>位置 X Y Z
    insertionPoint(0) = insertionPointX: insertionPoint(1) = insertionPointY: insertionPoint(2) = 0
    '文字>文字對齊 X Y Z
    alignmentPoint(0) = alignmentPointX: alignmentPoint(1) = alignmentPointY: alignmentPoint(2) = 0
    
    Set textObj = dwg.ModelSpace.AddText(textString, insertionPoint, textHeight)
    '文字>對正
    'textObj.Alignment = acAlignmentRight
    textObj.Alignment = alignmenType
    textObj.TextAlignmentPoint = alignmentPoint

    '文字>形式
    textObj.StyleName = textStyle

End Function
