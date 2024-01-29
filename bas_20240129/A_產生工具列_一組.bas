Attribute VB_Name = "A_產生工具列_一組"
'詳 [ThisWorkbook] 設定如何使此工具列在其他EXCEL裡看不到
'-------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Const ToolBarName As String = "萬用無敵檔"
Private Sub Auto_Open()
    Call MENU_BAR
End Sub
Private Sub MENU_BAR()

    Dim iCtr As Long

    Dim MacNames As Variant
    Dim CapNamess As Variant
    Dim TipText As Variant

    Call RemoveMenubar

    MacNames = Array("SATH", "GET_SHEET")

    CapNamess = Array("SHEET另存新檔", "GET_SHEET")

    TipText = Array("將目前選定的SHEET另存新檔", "載入它檔的指定工作表")

    With Application.CommandBars.Add
        .Name = ToolBarName
        .Left = 500
        .Top = 200
        .Protection = msoBarNoProtection
        .Visible = True
        '.Position = msoBarFloating
        .Position = msoBarTop

        For iCtr = LBound(MacNames) To UBound(MacNames)
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & MacNames(iCtr)
                .Caption = CapNamess(iCtr)
                .Style = msoButtonIconAndCaption
                .FaceId = 71 + iCtr
                .TooltipText = TipText(iCtr)
            End With
        Next iCtr
    End With
End Sub
Private Sub RemoveMenubar()
    On Error Resume Next
    Application.CommandBars(ToolBarName).Delete
    On Error GoTo 0
End Sub
Private Sub Auto_Close()
    Call RemoveMenubar
End Sub
