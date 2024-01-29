VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
'  Call ChkVer
'End Sub

'Private Sub Workbook_Open()
'  Call LoadMod
'  Call ChkVer
'  Worksheets("dwg_list").Range("B2").Value = ThisWorkbook.Path & "\"
End Sub

Private Sub KillMod()
'  Dim vbcs As VBComponents, vbc As VBComponent
'
'  Set vbcs = Application.VBE.ActiveVBProject.VBComponents
'  For Each vbc In vbcs
'    Select Case vbc.Type
'      Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
'        vbcs.Remove vbc
'    End Select
'  Next
End Sub

Private Sub LoadMod()
'  Dim fp$
'
'  fp = "\\s516\tblk\Modules\"
'
'  KillMod
'
'  With Application.VBE.ActiveVBProject.VBComponents
'    .Import fp & "1907 P3D ISOπœ¿ÀÆ∆ & SUPT MODEL DBF¿ÀÆ∆.bas"
'  End With
  
End Sub

Private Sub Workbook_Deactivate()
  On Error Resume Next
  Application.CommandBars(ToolBarName).Visible = False
  On Error GoTo 0
End Sub

Private Sub Workbook_Activate()
   On Error Resume Next
  Application.CommandBars(ToolBarName).Visible = True
  On Error GoTo 0
End Sub
