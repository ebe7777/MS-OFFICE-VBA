Attribute VB_Name = "A_���ͤu��C_�h��"

Public Const ToolBarName As String = "PDMS-ISO�ϭק�"
Public Const ToolBarName2 As String = "PDMS-ISO�ϭק�2"
Private Sub Auto_Open()
    Call MENU_BAR
End Sub
Private Sub MENU_BAR()

    Dim iCtr As Long

    Dim MacNames As Variant
    Dim CapNamess As Variant
    Dim TipText As Variant
 Dim MacNames2 As Variant
    Dim CapNamess2 As Variant
    Dim TipText2 As Variant
    
    Call RemoveMenubar

    MacNames = Array("program1", "program2", "program3")

    CapNamess = Array("[���J���]�n�ק諸��-���|��CAD���e", "[���J���]�ΥH�ѦҪ���-���|", "[���J���]�ΥH�ѦҪ���-CAD���e")

    TipText = Array("11", "12", "13")
    
    MacNames2 = Array("program4", "program5", "program6", "program7")

    CapNamess2 = Array("[�۰ʶi��]�����i��", "[�۰ʶi��]ABC��>0��,��l�i��", "�^�_��W�@���۰ʶi���e", "ISO�ϭק�")

    TipText2 = Array("21", "22", "23", 24)


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
    
    With Application.CommandBars.Add
        .Name = ToolBarName2
        .Left = 500
        .Top = 200
        .Protection = msoBarNoProtection
        .Visible = True
        '.Position = msoBarFloating
        .Position = msoBarTop

        For iCtr = LBound(MacNames2) To UBound(MacNames2)
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & MacNames2(iCtr)
                .Caption = CapNamess2(iCtr)
                .Style = msoButtonIconAndCaption
                .FaceId = 71 + iCtr
                .TooltipText = TipText2(iCtr)
            End With
        Next iCtr
    End With
End Sub
Private Sub RemoveMenubar()
    On Error Resume Next
    Application.CommandBars(ToolBarName).Delete
    Application.CommandBars(ToolBarName2).Delete
    On Error GoTo 0
End Sub
Private Sub Auto_Close()
    Call RemoveMenubar
End Sub


