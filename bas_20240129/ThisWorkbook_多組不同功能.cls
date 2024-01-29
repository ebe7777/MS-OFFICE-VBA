VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Deactivate()
  On Error Resume Next
  Application.CommandBars(ToolBarName1).Visible = False
  Application.CommandBars(ToolBarName2).Visible = False
  Application.CommandBars(ToolBarName3).Visible = False
  On Error GoTo 0
End Sub

Private Sub Workbook_Activate()
   On Error Resume Next
  Application.CommandBars(ToolBarName1).Visible = True
  Application.CommandBars(ToolBarName2).Visible = True
  Application.CommandBars(ToolBarName3).Visible = True
  On Error GoTo 0
End Sub
