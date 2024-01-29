VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormShowCellInfoSimple 
   Caption         =   "錯誤訊息列表"
   ClientHeight    =   3555
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4425
   OleObjectBlob   =   "UserFormShowCellInfoSimple.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormShowCellInfoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========
'開發者     brucechen1@micb2b.com
'開發起始日 2023-03-09
'修改日期   2023-11-10
'=========


''用途：
''   在程式運作過程，將特定狀態的儲存格顯示在form上，點選form上的資料可跳到該處

''使用方式：
''   設定以下全域變數
''顯示狀況工具
'Public showCellInfoCount As Long
'Public showCellInfoArray()
'    '[1,n] 工作表名稱 [2,n]欄名稱 [3,n]列號碼 [4,n]狀況說明 [5,n]狀況代碼
'    '[n,#]第幾筆資料
'Public showCellInfoListArray()
'    '[#,0]第幾筆資料要顯示的資訊、從0開始
'    '[n,0]沒意義，配合ListBox可以讀取的陣列型態而設
''   是否有檢查到狀況
'Public showCellInfoHaveData As Boolean
''   檢查的方式
'Public userFormShowCellInfoCheckMode As String
''   避免form執行過程中反覆執行不必要的動作
'Public skipUserFormShowCellInfo As Boolean
''   在UserFormShowCellInfoSimple的左下角告訴使用者處理方式
'Public showCellInfoMsg
''   檢查模式代碼(依據專案不同修改)
'Public checkModePriceAdjustCoefEmpty As String, checkModePickPriceEmpty As String

'===放在呼叫此form的程式裡===
''  放在程式執行最上面
''  防呆-使用者沒關掉錯誤訊息視窗==>關掉
'    If (IsFormInitialized("UserFormShowCellInfoSimple") = True) Then
'        Unload UserFormShowCellInfoSimple
'    End If

''  呼叫檢查程式
''  設定檢查模式
'   userFormShowCellInfoCheckMode = checkMode1
'   Call userFormShowCellInfoExcuteCheck(userFormShowCellInfoCheckMode)

''   執行過程將狀況說明寫入showCellInfoArray
''   showCellInfoCount 初始值為0
'   showCellInfoCount = showCellInfoCount + 1
'   ReDim Preserve showCellInfoArray(5,iCount)
'   showCellInfoArray(1, showCellInfoCount) = mySht.Name
'   showCellInfoArray(2, showCellInfoCount) = myCN
'   showCellInfoArray(3, showCellInfoCount) = i
'   showCellInfoArray(4, showCellInfoCount) = "狀況說明"
'   showCellInfoArray(5, showCellInfoCount) = "狀況代碼"


        
''   如果showCellInfoArray有寫入資訊，將其轉化為ListBox可以讀取的陣列型態 ==> 放在該段程式最後面
'    If (UBound(showCellInfoArray, 2) <> 0) Then
'        ReDim showCellInfoListArray(UBound(showCellInfoArray, 2) - 1, 0)
'        Call transformArryToList(showCellInfoArray, showCellInfoListArray)
'        showCellInfoHaveData = True
'    End If

''  告訴使用者檢查到了什麼訊息
'    If (showCellInfoHaveData = True) Then
'        msgTitle = "注意            "    ' 定義標題。
'
'        msgText = "發現狀況說明" + vbLf   ' 定義訊息。
'        msgText = msgText + "====================================" + vbLf ' 定義訊息。
'        msgText = msgText + "請執行以下動作：" + vbLf   ' 定義訊息
'        msgText = msgText + "(1)修改視窗顯示問題" + vbLf   ' 定義訊息
'        msgText = msgText + "(2)再次執行此程式"    ' 定義訊息
''        msgStyle = vbCritical '顯示"X"圖案
''        msgStyle = vbExclamation '顯示"!"圖案
''        msgStyle = vbInformation '顯示"i"圖案
'        MsgBox msgText, msgStyle, msgTitle
'        UserFormShowCellInfoSimple.Show False
'        GoTo 999
'    End If

''===以下funciton為執行所必須，放到存放Func的模組中===
'Public Function transformArryToList(originalArray(), listArray())
''將指定陣列的資料轉化為ListBox可以讀取的陣列型態
'Dim i As Long
'Dim iStr1 As String
'    For i = 1 To UBound(originalArray, 2)
'        '將資訊串成一個字串
'        '   [1,n] 工作表名稱 [2,n]欄名稱 [3,n]列號碼 [4,n]狀況說明 [5,n]狀況代碼
'        '   [n,#]第幾筆資料
'        iStr1 = originalArray(1, i) & "-" & originalArray(2, i) & originalArray(3, i) & " : " & originalArray(4, i)
'        listArray(i - 1, 0) = iStr1
'    Next i
'
'End Function

'Public Function IsFormInitialized(FormName As String) As Boolean
'    '檢查UserFormShowCellInfoSimple是否被initialized
'    'Does not have the side effect of needing to load the form just to see if it's loaded.
'    Dim myForm As Variant
'    For Each myForm In UserForms
'        If myForm.Name = FormName Then
'            IsFormInitialized = True
'            Exit Function
'        End If
'    Next
'End Function


'Public Function userFormShowCellInfoExcuteCheck(checkMode)
''   檢查資料的運作寫在此function中
''   如果有多個檢查方式，使checkMode(全域變數為userFormShowCellInfoCheckMode)控制控制
'Dim totalRow As Long
'Dim iStr1 As String
'Dim i As Long
''====以下為範例中藥改要的東西====
''myWorkSheet 被檢查的工作表
''myCheckThisColumn 被檢查的欄
''checkMode1/checkMode2 檢查模式名稱-各種檢查功能要獨立一個檢查模式名稱
''dataStartRow 要檢查的範圍的起始列
''================================
'    Erase showCellInfoArray
'    showCellInfoCount = 0
'    '防呆-使用者沒關掉錯誤訊息視窗==>關掉
'    If (IsFormInitialized("UserFormShowCellInfoSimple") = True) Then
'        Unload UserFormShowCellInfoSimple
'    End If
'    With myWorkSheet
'        totalRow = myDataRows(ThisWorkbook.Name, .Name, myCheckThisColumn, 65536)
'        If (checkMode = checkMode1) Then
'            '檢查1 <什麼什麼狀況...>
'            For i = dataStartRow To totalRow
'                iStr1 = .Range(myCheckThisColumn & i).Value
'                If (Trim(iStr1) = "") Then
'                    showCellInfoCount = showCellInfoCount + 1
'                    ReDim Preserve showCellInfoArray(5, showCellInfoCount)
'                    '[1,n] 工作表名稱 [2,n]欄名稱 [3,n]列號碼 [4,n]狀況說明 [5,n]狀況代碼
'                    '[n,#]第幾筆資料
'                    showCellInfoArray(1, showCellInfoCount) = myWorkSheet
'                    showCellInfoArray(2, showCellInfoCount) = myCheckThisColumn
'                    showCellInfoArray(3, showCellInfoCount) = i
'                    showCellInfoArray(4, showCellInfoCount) = "告訴使用者什麼狀況"
'                    showCellInfoArray(5, showCellInfoCount) = checkMode
'                End If
'            Next i
'        ElseIf (userFormShowCellInfoCheckMode = checkMode2) Then
'            '檢查1 <什麼什麼狀況...>
'            '   重複撰寫檢查狀況，以及有問題時要輸入showCellInfoArray的值
'            Next i
'        End If
'    End With
'End Function


'^^^^說明結束^^^^


Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
Private Sub UserForm_Initialize()
    ListBox1.Clear
    ListBox1.List = showCellInfoListArray
    skipUserFormShowCellInfo = False
    ListBox1.ListIndex = 0
End Sub

Private Sub ListBox1_Click()
'''範本
'''點選列表時的動作
''Dim i As Long
''Dim myMode As String, showCellInfoMsg As String
''    If (skipUserFormShowCellInfo = False) Then
''        '說明做變化
''        i = ListBox1.ListIndex + 1
''        '   showCellInfoArray
''        '       [1,n] 工作表名稱 [2,n]欄名稱 [3,n]列號碼 [4,n]錯誤狀況 [5,n]錯誤代碼
''        '       [n,#]第幾筆資料
''        '   listbox index號碼等於原始第二維-1
''        myMode = showCellInfoArray(5, i)
''        Select Case myMode
'''依據狀況對各種檢查模式制定一段訊息，放在UserFormShowCellInfoSimple左下角告訴使用者處理方式
''            Case checkModePriceAdjustCoefEmpty
''                showCellInfoMsg = "微調係數不可是空白，請修改"
''            Case checkModePickPriceEmpty
''                showCellInfoMsg = "此格需要人工處理"
''        End Select
''        Label2.Caption = showCellInfoMsg
''        '移動到該處
''        Worksheets(showCellInfoArray(1, i)).Select
''        Range(showCellInfoArray(2, i) & showCellInfoArray(3, i)).Select
''        With Selection
''            .Borders(xlDiagonalDown).LineStyle = xlContinuous
''            .Borders(xlDiagonalDown).Weight = xlThick
''            .Borders(xlDiagonalUp).LineStyle = xlContinuous
''            .Borders(xlDiagonalUp).Weight = xlThick
''            Application.Wait Now + 1 / (24 * 60 * 60# * 1)
''            .Borders(xlDiagonalDown).LineStyle = xlNone
''            .Borders(xlDiagonalUp).LineStyle = xlNone
''        End With
''    End If

'點選列表時的動作
Dim i As Long
Dim myMode As String, showCellInfoMsg As String
    If (skipUserFormShowCellInfo = False) Then
        '說明做變化
        i = ListBox1.ListIndex + 1
        '   showCellInfoArray
        '       [1,n] 工作表名稱 [2,n]欄名稱 [3,n]列號碼 [4,n]錯誤狀況 [5,n]錯誤代碼
        '       [n,#]第幾筆資料
        '   listbox index號碼等於原始第二維-1
        myMode = showCellInfoArray(5, i)
        Select Case myMode
'依據狀況對各種檢查模式制定一段訊息，放在UserFormShowCellInfoSimple左下角告訴使用者處理方式
            Case checkModePriceAdjustCoefEmpty
                showCellInfoMsg = "微調係數不可是空白，請修改"
            Case checkModePickPriceEmpty
                showCellInfoMsg = "此格需要人工處理"
        End Select
        Label2.Caption = showCellInfoMsg
        '移動到該處
        Worksheets(showCellInfoArray(1, i)).Select
        Range(showCellInfoArray(2, i) & showCellInfoArray(3, i)).Select
        With Selection
            .Borders(xlDiagonalDown).LineStyle = xlContinuous
            .Borders(xlDiagonalDown).Weight = xlThick
            .Borders(xlDiagonalUp).LineStyle = xlContinuous
            .Borders(xlDiagonalUp).Weight = xlThick
            Application.Wait Now + 1 / (24 * 60 * 60# * 1)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
    End If
End Sub
Private Sub CommandButton1_Click()
'對目前指定的檢查方式再重新檢查
    Me.Hide
    Call userFormShowCellInfoExcuteCheck(userFormShowCellInfoCheckMode)
    If (UBound(showCellInfoArray, 2) <> 0) Then
        ListBox1.Clear
        ListBox1.List = showCellInfoListArray
        skipUserFormShowCellInfo = False
        ListBox1.ListIndex = 0
        UserFormShowCellInfoSimple.Show False
    Else
        '修正完成後，再度檢查後如果沒問題顯示的訊息
''範本
'        msgTitle = "訊息            "    ' 定義標題。
'
'        msgText = "沒有發現錯誤" + vbLf   ' 定義訊息。
'        msgText = msgText + "====================================" + vbLf  ' 定義訊息。
'        msgText = msgText + "請再次執行以下功能：" + vbLf  ' 定義訊息
'        msgText = msgText + "工作表 [" & operateSN & "] => 按鈕 [產生報表]"   ' 定義訊息
'        msgStyle = vbInformation '顯示"i"圖案
'        MsgBox msgText, msgStyle, msgTitle
'        Unload Me
        '        msgTitle = "訊息            "    ' 定義標題。
        If (userFormShowCellInfoCheckMode = checkModePriceAdjustCoefEmpty) Then
            msgText = "沒有發現錯誤" + vbLf   ' 定義訊息。
            msgText = msgText + "====================================" + vbLf  ' 定義訊息。
            msgText = msgText + "請再次執行以下功能：" + vbLf  ' 定義訊息
            msgText = msgText + "工作表 [" & operateSN & "] => 按鈕 [產生報表]"   ' 定義訊息
            msgStyle = vbInformation '顯示"i"圖案
            MsgBox msgText, msgStyle, msgTitle
            Unload Me
        ElseIf (userFormShowCellInfoCheckMode = checkModePickPriceEmpty) Then
            msgText = "沒有發現空白處"    ' 定義訊息。
            msgStyle = vbInformation '顯示"i"圖案
            MsgBox msgText, msgStyle, msgTitle
            Unload Me
        End If
    End If
End Sub

