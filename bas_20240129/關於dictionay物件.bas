Attribute VB_Name = "關於dictionay物件"
Sub Dict的操作_資料庫()
Attribute Dict的操作_資料庫.VB_ProcData.VB_Invoke_Func = " \n14"
'注意!!
'使用 監看式 在檢查dict內容時，只能看到256筆資料

'https://excelmacromastery.com/vba-dictionary/
'Dict用來記錄資料
'ArrayList用來排序
Dim iVar As Variant
Dim iDict As Object, iDictNew As Object

    '參考   https://excelmacromastery.com/vba-dictionary/
    '       https://excelmacromastery.com/vba-arraylist/
'===late binding
    Set iDict = CreateObject("Scripting.Dictionary")
    Set iDictNew = CreateObject("Scripting.Dictionary")

'===dicionary是否區分大小寫
    '設定iDict是否case sensitive
    '   是(預設)
    iDict.CompareMode = vbBinaryCompare
    '   否
    iDict.CompareMode = vbTextCompare

'***製造一些資料測試用
    iDict.Add "key1", "Value1"
    iDict.Add "key2", "Value2"
    iDict.Add "key3", "Value3"
    
'===dicitonary的資料筆數
'   沒資料(被.RemoveAll)=0,add過1次=1
    iVar = dict.Count
    
'===查詢某個key是否存在
    If (iDict.Exists("key1") = True) Then
        '...
    End If
'===將iDict的內容看一遍
    For Each iVar In iDict
        '取得key
        mykey = iVar
        '取得Value
        myVal = iDict(iVar)
    Next iVar
    
    '驗證dict內容
    i = 0
    For Each iVar In iDict
        i = i + 1
        With Worksheets("test")
            '取得key
            .Cells(i, 1) = iVar
            '取得Value
            .Cells(i, 2) = iDict(iVar)
        End With
    Next iVar

'===其他
    '取得dict中某key的value
    myVal = iDict("key1")
    'dict中的資料數量(沒有任何資料時是0)
    myVal = iDict.Count
    '移除dict中特定資料(dict資料數量會改變)
    iDict.Remove "key1"
    '移除dict中所有資料(dict資料數量會改變)
    iDict.RemoveAll
    
    '修改特定Key裡的Value
    '   沒辦法，只能將該key移除再重新加入
    
    '取得dict當中所有key的Value的最大值/最小值
    MsgBox Application.max(iDict.items)
    MsgBox Application.min(iDict.items)
    '將dict物件設為nothing - 不是將資料移除，而是將此變數設為不再是字典物件
    '   先重新定義這個物件再設為nothing可減少執行時間
    Set iDict = CreateObject("Scripting.Dictionary")
    Set iDict = Nothing
    If (iDict Is Nothing = True) Then
       MsgBox 1
    End If
    
'===使用ArrayList將dictionay排序
'**注意** 如果電腦沒安裝 .NET Framwork3.5 會無法使用arrlist
'https://stackoverflow.com/questions/40625618/automation-error-2146232576-80131700-on-creating-an-array
Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")

    For Each iVar In iDict
        arrList.Add iVar
    Next iVar
    '排序-1~9,A~B
    arrList.Sort
    '將目前結果顛倒排列
    arrList.Reverse
    '將排序結果放到新的dictionary物件
    For Each iVar In arrList
        iDictNew.Add iVar, iDict(iVar)
    Next iVar
    
End Sub

Sub test()
Dim iDict As Object
    If (iDict Is Nothing = True) Then
       MsgBox 1
    End If
    Set iDict = CreateObject("Scripting.Dictionary")
    Set iDict = Nothing
    If (iDict Is Nothing = True) Then
       MsgBox 2
    End If
    
    iDict.Add "A", 1
    iDict.Add "B", 1
    iDict.Remove "A"
    myVal = iDict.Count
    
End Sub
