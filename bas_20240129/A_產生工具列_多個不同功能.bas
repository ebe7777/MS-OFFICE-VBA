Attribute VB_Name = "A_���ͤu��C_�h�Ӥ��P�\��1"
'�� [ThisWorkbook] �]�w�p��Ϧ��u��C�b��LEXCEL�̬ݤ���
'-------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Const ToolBarName1 As String = "�M�׶i�ת�1" '"�U�εL����"
Public Const ToolBarName2 As String = "�M�׶i�ת�2" '"�U�εL����"
Public Const ToolBarName3 As String = "�M�׶i�ת�3" '"�U�εL����"
Private Sub Auto_Open_��Ʈw()
    Call MENU_BAR
End Sub
Private Sub MENU_BAR_��Ʈw()
'���Ϊk�ӷ�
'   https://zhuanlan.zhihu.com/p/81161115
'�U�ث��s(�g�L���եu���H�U�X�إi�H���)
'   https://docs.microsoft.com/zh-tw/office/vba/api/office.msocontroltype
'msoControlPopup            ����-�I���k���X�{�U�@�����s�M��(�p�P�b����W���ƹ��k��)
'msoControlDropdown         �U�Ԧ��M��-���i��J
'msoControlComboBox         �U�Ԧ����-�i��J,�ŦX�̷|�۰ʱa�X�F��J�D��椤���|�X��
'msoControlButton           �R�O���s
'msoControlEdit             ��r���(�i��J��r)
'�U�ث��s��Style
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211207(v=office.11)
'Ū��msoControlDropdown����
'   https://club.excelhome.net/thread-223737-1-1.html
    Dim subName As Variant
    Dim captionText As Variant
    Dim tipText As Variant

    Call RemoveMenubar
    
    '���s1 �]�w
    With Application.CommandBars.Add
        .Name = ToolBarName1
        .Left = 0
        .Top = 0
        .Protection = msoBarNoProtection
        .Visible = True
'        .Position = msoBarFloating
        .Position = msoBarTop
'        .Position = msoBarBottom
        
        '   msoControlButton
        subName = "doSetting"
        captionText = "1.�]�w"
        tipText = "��{���\��B�u�@��W�ٵ����]�w"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoButtonCaption 'msoButtonIconAndCaption
            '�H�Ʀr1�B2�B3...���
            '71���Ʀr1�B71���Ʀr2...
            .FaceId = 71
            .TooltipText = tipText
        End With
    End With
    
    
    '���s2 ���]�u�@��
    With Application.CommandBars.Add
        .Name = ToolBarName2
        .Left = 0
        .Top = 0
        .Protection = msoBarNoProtection
        .Visible = True
'        .Position = msoBarFloating
        .Position = msoBarTop
'        .Position = msoBarBottom
        
        '   msoControlButton
        subName = "test2"
        captionText = "2.���]�u�@��"
        tipText = "�ϥΦ��\��N�U�u�@��B�U�x�s��^�_���Τ@�榡"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoButtonCaption 'msoButtonIconAndCaption
            '�H�Ʀr1�B2�B3...���
            '71���Ʀr1�B71���Ʀr2...
            .FaceId = 72
            .TooltipText = tipText
        End With
    End With
    
    '���s3 �u�@���� ���h�]�w
    With Application.CommandBars.Add
        .Name = ToolBarName3
        .Left = 0
        .Top = 0
        .Protection = msoBarNoProtection
        .Visible = True
        .Position = msoBarFloating
'        .Position = msoBarTop
'        .Position = msoBarBottom

        '   msoControlDropdown
        subName = "test3"
        captionText = "3.�u�@���h"
        tipText = "�N�i�ת�B���Ʀ۰ʰ����h�]�w"
        With .Controls.Add(Type:=msoControlDropdown)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoComboLabel
            .TooltipText = tipText
            'msoControlDropdown/msoControlComboBox �U�Ԧ�����
            .AddItem "�Ĥ@��", 1
            .AddItem "�ĤG��", 2
            .AddItem "�ĤT��", 3
            .DropDownLines = 3
            .DropDownWidth = 75
            .ListIndex = 0
        End With
'���o�ĴX�����W��
'    With CommandBars(ToolBarName3).Controls(1)
'   ���oitem�W��(�Ĥ@���B�ĤG��...)
'        MsgBox .List(.ListIndex)
'   ���oitem���X(1,2....)
'       MsgBox .ListIndex
'    End With
    End With
End Sub
Private Sub RemoveMenubar_��Ʈw()
    On Error Resume Next
    Application.CommandBars(ToolBarName1).Delete
    Application.CommandBars(ToolBarName2).Delete
    Application.CommandBars(ToolBarName3).Delete
    On Error GoTo 0
End Sub
Private Sub Auto_Close_��Ʈw()
    Call RemoveMenubar
End Sub
