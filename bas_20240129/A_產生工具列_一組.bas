Attribute VB_Name = "A_���ͤu��C_�@��"
'�� [ThisWorkbook] �]�w�p��Ϧ��u��C�b��LEXCEL�̬ݤ���
'-------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Const ToolBarName As String = "�U�εL����"
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

    CapNamess = Array("SHEET�t�s�s��", "GET_SHEET")

    TipText = Array("�N�ثe��w��SHEET�t�s�s��", "���J���ɪ����w�u�@��")

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
