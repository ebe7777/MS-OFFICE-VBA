VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Deactivate()
    On Error Resume Next
        Application.CommandBars(ToolBarName).Visible = False
        Application.CommandBars(ToolBarName2).Visible = False
    On Error GoTo 0
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
        Application.CommandBars(ToolBarName).Visible = True
        Application.CommandBars(ToolBarName2).Visible = True
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    'Call LoadMod
   
End Sub

Private Sub LoadMod()
    Dim fp$
    
    fp = "D:\Bruce工作資料\M5報價發包採購資訊系統\bas存放測試\"
    
    KillMod
    
    With Application.VBE.ActiveVBProject.VBComponents
        .Import fp & "A_產生工具列_一組.bas"
        .Import fp & "B_getDataShtInfo.bas"
    End With
  
End Sub
Private Sub KillMod()
    Dim vbcs As VBComponents, vbc As VBComponent
    
    Set vbcs = Application.VBE.ActiveVBProject.VBComponents
    For Each vbc In vbcs
        Select Case vbc.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
            vbcs.Remove vbc
        End Select
    Next
End Sub
