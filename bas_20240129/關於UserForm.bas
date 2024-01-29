Attribute VB_Name = "關於UserForm"
Sub 通用知識()
'combobox、listbox如果修改後的值和修改前相同，不會觸發物件的_Click/_Change...等sub
'   ==>先變成其他值再變成修改後的值
'   唯一能在沒變動也能執行的sub是 DropButtonClick

'許多物件(combobox、listbox、scrollbar)在撰寫其他功能時可能會改到這些物件並觸發_Click/_Change...等sub，但可能當下不該觸發
'   ==>在這些sub中預先設定好開關 (if true do something,if false do nothing)
End Sub
Sub 開啟form()
    UserForm1.Show False
    '如果userForm的內容(label的caption等)沒顯示內容
    DoEvents
End Sub
Sub 隱藏form()
    UserForm1.Hide
End Sub
Sub 移除form()
    '當form可能需要在多個excel檔中開開關關時(一個excel存放程式，其他excel打開使用)，如果只用hide，當再度show時會active hide此form的excel檔
    '  >在Terminate上放 unload
    Private Sub UserForm_Terminate()
'        Unload UserForm1
'    End Sub
    Unload UserForm1
End Sub

Public Function 資料庫_IsFormInitialized(FormName As String) As Boolean
'測試某名稱的UserForm是否已經Initialized
    Dim myForm As Variant
    For Each myForm In UserForms
        If myForm.Name = FormName Then
            IsFormInitialized = True
            Exit Function
        End If
    Next
End Function

Sub 刷新form()
'有些情況下form會發生不正常的狀況，已知以下要用此處理
'   (1)當image被click後，重新設置picture到該image上時，image上仍會是舊圖片不(雖然實際有替換)
    Me.Repaint
End Sub
Sub FormActivate的通用設定()
    '將Form擺在excel所在畫面中央
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
Sub Form裡面的屬性設定說明()
    
    'picture-圖片設定
    myObj.Picture = LoadPicture("c:\123.png")
    '   圖片是否有設定
    If (Not Image1.Picture Is Nothing) Then
        'do something
    End If
    'foreColor-字體顏色設定,bas資料夾中有RGB表
    myObj.ForeColor = RGB(204, 204, 204)
    
End Sub
Sub 雜項()
'檢查textbox是否輸入integer
'https://stackoverflow.com/questions/26138833/making-vba-form-textbox-accept-numbers-only-including-and

'呼叫其他工作表的 Private Sub Worksheet_Change(ByVal Target As Range)
Dim myWS As Worksheet
Dim myShtCodeName As String 'myShtCodeName:工作表的CodeName屬性值
Dim myRg As Range   '工作表改變時，正在修改的範圍
Set myWS = Sheet("某張工作表")
Set myRg = myWS.Range("A1")
    myShtCodeName = myWS.CodeName
    Application.Run myShtCodeName & ".Worksheet_Change", myRg

'呼叫其他form的commandbutton_click
'   if 'the module' is the same form module as commandbutton1_click then yes you can.
'   If it is a different module you need to replace private with public for the commandbutton1_click sub and the form itself needs to be open.
'   And to call it you would use
'   參考資料
'       https://www.access-programmers.co.uk/forums/threads/calling-commandbutton_click-event-from-a-module.275200/

Form_whatevermyformiscalled.CommandButton1_Click
End Sub
Sub 找到form上的物件()
'直接指定
Dim myLB As MSForms.ListBox
    Set myLB = Form03Inquiry.ListBox010101
'使用物件名稱
Dim myForm As UserForm, TextBoxName As String, myObj As Object
Dim i As Integer

    Set myForm = UserForm1
    TextBoxName = "TextBox1"
    
    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
End Sub
Function 將Form的物件帶入function的變數中(ByRef TextBoxName As String)
    With Me.Controls(TextBoxName)
            .BackStyle = fmBackStyleTransparent
    End With
End Function
Sub 設定textBox的值()

Dim myForm As UserForm, TextBoxName As String, myObj As Object
Dim i As Integer

Set myForm = UserForm1
TextBoxName = "TextBox1"

    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
    myObj.Value = "123"
    
    '要讓顯示值換行要設定 MultiLine 為 True.
End Sub
Sub 設定TextBox的顯示()
    'textbox可以操作時
    TextBox1.Enabled = True
    TextBox1.BackStyle = fmBackStyleOpaque
    'textbox不能操作時
    TextBox202.Enabled = False
    TextBox202.BackStyle = fmBackStyleTransparent
End Sub
Sub Scrollbar設定內容()
    '設定scrollbar
    ScrollBar1.min = 1
    ScrollBar1.max = i
    ScrollBar1.Value = 1
    '如果換頁後有image要替換，則每次換頁後須執行此
    Me.Repaint
End Sub
Sub multipage設定內容()
    'activate page (0~n)
    MultiPage1.Value = 0
    '設定某page的屬性
    MultiPage1.Pages(1).Enabled = False
    

End Sub
Sub 資料庫_MultiPage1_Change()
'各輸入頁切換時，前面加個check mark
Dim i As Long
    i = MultiPage1.Value
    '先設為原始值
    MultiPage1.Pages(0).Caption = "主頁"
    MultiPage1.Pages(1).Caption = "主管"
    '目前選取的page名稱修改
    Select Case i
        Case 0
            MultiPage1.Pages(0).Caption = ChrW(&H2611) & "主頁"
        Case 1
            MultiPage1.Pages(1).Caption = ChrW(&H2611) & "主管"
    End Select
'選不同標籤時form的高度、寬度改變
'   在Form的 Private Sub UserForm_Initialize() 也要做相同的設定
Dim myHei As Long, myWidth As Long
    Select Case MultiPage1.Value
        Case 0
            myHei = 92.5
            myWidth = 161
        Case 1
            myHei = 127.5
            myWidth = 183
    End Select
    Me.Height = myHei
    Me.Width = myWidth
End Sub
Sub 設定comboBox的下拉式清單值()
'如果要讓使用者無法手動修改combobox內容，找到屬性[Style]，改成 [2-fmStyleDropDownList]
'ListIndex從0開始

    '手動填入
    ComboBox1.AddItem "帆宣"
    ComboBox1.AddItem "元通"
    ComboBox1.ListIndex = 0
    '用array填入-直接指定特定的combobox
    ComboBox1.List = iArray
    '取得選取的Combobox的值
    myVal = ComboBox1.Value
    '用array填入-form上有許多combobox-->找到特定的
Dim myForm As UserForm, comboBoxName As String, myObj As Object
Dim myArray()
Dim i As Integer

Set myForm = UserForm1
comboBoxName = "ComboBox1"
    
    ReDim myArray(2)
    For i = 1 To 3
        myArray(i - 1) = i
    Next i
    
    For i = 0 To myForm.Controls.Count
        If (myForm.Controls.Item(i).Name = comboBoxName) Then
            Set myObj = myForm.Controls.Item(i)
            Exit For
        End If
    Next i
    myObj.List = myArray
    '設定初始值
    ComboBox1.ListIndex = 0

End Sub
Sub listBox設定內容()
    '取得listboxt目前select物件的值
    '   如果第一次開啟form且沒有人工點選ListBox(以ListBox.ListIndex取代人工)
    '       會使得.value抓不到值；使用.SetFocus可解決此問題
    listBoxObj.SetFocus
    myVal = listBoxObj.Value
    '設定columns數目，沒設定以listbox物件的原始值
    listBoxObj.ColumnCount = 10
    '指定column的寬度
    '   各別指定，如無指定者統一讓系統設置
    listBoxObj.ColumnWidths = "50,50,50"
    
    '增添資料
    '   方式1(不能使用ColumnHeads)，使用.list，但限制最多10 columns
    listBoxObj.AddItem
    listBoxObj.List(i - 1, ii) = dataArray(i, ii)
    '       清空listbox方式
    listBoxObj.Clear
    '   方式2(不能使用ColumnHeads)，使用array，沒column限制
    '       使用的array需為2維陣列 [列,欄] 輸入值
    '           每一維的起始值為0
    '       欄的數量需自行設定，否則只顯示1欄(預設值)
    'listBoxObj.ColumnCount = UBound(dataArray, 2)
    '           由於欄寬不好控制，資料如有多欄建議多設幾個ListBox分別放各欄的資料
    '       要讓水平捲動軸出現，需修改屬性 ColumnWidths ，改成一個大於屬性 Width 的值
    listBoxObj.List = dataArray

    '   方式3(可以使用ColumnHeads)，使用工作表中的資料，沒column限制
    '       Adding Column Headers
    '       You can only display column headers when you use the RowSource property, not when you use an array or add items individually.
    '       To display column headers set the ColumnHeads property to True.
    '       Do not include the column headings on the worksheet in the range defined for RowSource.
    '       The row directly above the first row of the RowSource will be automatically used.
    '       !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '       注意!!當ListBox的rawsource有設定時，刪除與填入資料到rawsource的工作表時會使LIST內容不斷更新而拖累程式執行速度
    '           所以當要刪除或新增資料到該工作表時，應先清空RAWSOURCE設定
    '       !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    listBoxObj.ColumnHeads = True
    '       指定範圍的上一列<range("A1:E1")>會自動抓取成為ColumnHeads
    '       *後面要使用(External:=True)，否則只會抓到active工作表的該range資訊
    listBoxObj.RowSource = Sheets("sysListDataTempSht").Range("A2:E15").address(External:=True)
    '       清空listbox方式
    listBoxObj.RowSource = vbNullString
    
    '判定目前listbox選定的是第幾筆資料 (第一筆資料是0)
    For i = 0 To ListBox1.ListCount - 1
        If (ListBox1.Selected(i) = True) Then
            listboxSelectedRow = i
        End If
    Next i
    
    '將ListBox顯示為沒有選取任何資料的狀態 deselect all
    '注意! 如果是在該ListBox的程序中，如 ListBox1_Click，則此方法不會有效果
    ListBox1.Selected(ListBox1.ListIndex) = False
    
    '取得目前選取項的值
    '   ListBox是從0開始
    ListBox1.List (0)
End Function
Private Function putDict1DimArryIntoListBoxArray_資料庫(dictArray(), ListBoxArray())
'將ebeDictionary的getKeys/getValues的值給ListBox用
'   一維陣列是從1開始，但ListBox需求的是從0開始；此功能將dict的資料陣列轉為ListBox用
Dim i As Long
    ReDim ListBoxArray(UBound(dictArray, 1) - 1)
    For i = 1 To UBound(dictArray, 1)
        ListBoxArray(i - 1) = dictArray(i)
    Next i
End Function
Private Function listBoxSeleAllorNot_資料庫(myListBox As MSForms.ListBox, seleAllorNot As Boolean)
'將指定的ListBox內的物件全選(True)/全不選(False)
    For i = 0 To myListBox.ListCount - 1
        myListBox.Selected(i) = seleAllorNot
    Next i
End Function

Sub optionButton的控制()
'同一個畫面中的optionButton會自動控制只讓你選的到一個
'   不同的MultiPage，當中的optionButton會各自自成一組

'判斷現在是哪個optionButton被選取 假設有兩個,名稱是 optionButton1/optionButton2
'   被選取的值會是true
OptionButton1.Value = True
OptionButton2.Value = False

End Sub

