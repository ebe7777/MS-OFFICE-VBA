Attribute VB_Name = "A_產生工具列_多個不同功能1"
'詳 [ThisWorkbook] 設定如何使此工具列在其他EXCEL裡看不到
'-------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Const ToolBarName1 As String = "專案進度表1" '"萬用無敵檔"
Public Const ToolBarName2 As String = "專案進度表2" '"萬用無敵檔"
Public Const ToolBarName3 As String = "專案進度表3" '"萬用無敵檔"
Private Sub Auto_Open_資料庫()
    Call MENU_BAR
End Sub
Private Sub MENU_BAR_資料庫()
'此用法來源
'   https://zhuanlan.zhihu.com/p/81161115
'各種按鈕(經過測試只有以下幾種可以選擇)
'   https://docs.microsoft.com/zh-tw/office/vba/api/office.msocontroltype
'msoControlPopup            快顯-點選後右側出現下一階按鈕清單(如同在物件上按滑鼠右鍵)
'msoControlDropdown         下拉式清單-不可輸入
'msoControlComboBox         下拉式方塊-可輸入,符合者會自動帶出；輸入非選單中的會出錯
'msoControlButton           命令按鈕
'msoControlEdit             文字方塊(可輸入文字)
'各種按鈕的Style
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211207(v=office.11)
'讀取msoControlDropdown的值
'   https://club.excelhome.net/thread-223737-1-1.html
    Dim subName As Variant
    Dim captionText As Variant
    Dim tipText As Variant

    Call RemoveMenubar
    
    '按鈕1 設定
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
        captionText = "1.設定"
        tipText = "對程式功能、工作表名稱等做設定"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoButtonCaption 'msoButtonIconAndCaption
            '以數字1、2、3...顯示
            '71為數字1、71為數字2...
            .FaceId = 71
            .TooltipText = tipText
        End With
    End With
    
    
    '按鈕2 重設工作表
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
        captionText = "2.重設工作表"
        tipText = "使用此功能將各工作表、各儲存格回復為統一格式"
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoButtonCaption 'msoButtonIconAndCaption
            '以數字1、2、3...顯示
            '71為數字1、71為數字2...
            .FaceId = 72
            .TooltipText = tipText
        End With
    End With
    
    '按鈕3 工作項目 階層設定
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
        captionText = "3.工作階層"
        tipText = "將進度表的B欄資料自動做分層設定"
        With .Controls.Add(Type:=msoControlDropdown)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & subName
            .caption = captionText
            .Style = msoComboLabel
            .TooltipText = tipText
            'msoControlDropdown/msoControlComboBox 下拉式選單用
            .AddItem "第一階", 1
            .AddItem "第二階", 2
            .AddItem "第三階", 3
            .DropDownLines = 3
            .DropDownWidth = 75
            .ListIndex = 0
        End With
'取得第幾階的名稱
'    With CommandBars(ToolBarName3).Controls(1)
'   取得item名稱(第一階、第二階...)
'        MsgBox .List(.ListIndex)
'   取得item號碼(1,2....)
'       MsgBox .ListIndex
'    End With
    End With
End Sub
Private Sub RemoveMenubar_資料庫()
    On Error Resume Next
    Application.CommandBars(ToolBarName1).Delete
    Application.CommandBars(ToolBarName2).Delete
    Application.CommandBars(ToolBarName3).Delete
    On Error GoTo 0
End Sub
Private Sub Auto_Close_資料庫()
    Call RemoveMenubar
End Sub
