VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ebeDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========
'開發者     brucechen1@micb2b.com
'開發日期   2023-03-10
'修改日期   2023-12-22
'=========

'''Sub 資料庫_所有功能使用範例()
'''Dim iDict1 As New ebeDictionary, iDict2 As New ebeDictionary
'''Dim iCount As Long
'''Dim iValue As String
'''Dim iBool As Boolean
'''Dim iArray()
'''    '新增資料到dict中 key,value (皆為string；雖然允許輸入純數字，但加入dict後會被強制轉為string)
'''    ' **key不可以重複**
'''    iDict1.Add 100.1, 1.5
'''    iDict1.Add "20", "B"
'''    iDict1.Add "AA", "C"
'''    iDict1.Add "王", "D"
'''    iDict1.Add "B", "E"
'''    '將一個dict的值複製到另一個dict
'''    iDict2.copy iDict1
'''    '取得屬性Count - 資料筆數
'''    iCount = iDict1.Count
'''    '由key取得value
'''    iValue = iDict1.GetValue("AA")
'''    '檢查某key是否存在
'''    iBool = iDict1.Exists("B")
'''    '取得所有的Value，以1維陣列的型態；需指定1維陣列的起始號碼
'''    Erase iArray
'''    iArray = iDict1.GetValues(3)
'''    '取得所有的Key，以1維陣列的型態；需指定1維陣列的起始號碼
'''    Erase iArray
'''    iArray = iDict1.GetKeys(3)
'''    '取得所有的Key和Value，以2維陣列的型態
'''    Erase iArray
'''    iArray = iDict1.GetKeysAndValues
'''    '取得某key在ItemsArray中的2維號碼
'''    iCount = iDict1.GetPageNo("20")
'''    '藉由PageNo取得key值
'''    iValue = iDict1.GetKeyByPageNo(iCount)
'''    '藉由PageNo取得Value值
'''    iValue = iDict1.GetValueByPageNo(iCount)
'''    '排序key - 要輸入ascend來由小至大排序，或descend來有大至小排序(ascend、Descend大小寫不拘)
'''    '   排序大小的規則為 中文>英文>純數字 (純中文/純英文/中英數交雜則是一個一個字比對)
'''    '       如果是降序descend(越來越小)，排列結果為是 中文=>英文=>純數字 (純中文/純英文/中英數交雜則是一個一個字比對)
'''    '       如果是昇序ascend(越來越大) ，排列結果為 純數字=>英文=>中文   (純中文/純英文/中英數交雜則是一個一個字比對)
'''    iDict1.SortKey ("ascend")
'''    '以key刪除特定資料
'''    iDict1.Remove ("A")
'''    '刪除所有輸入值
'''    iDict1.RemoveAll
'''
'''End Sub

'===設定屬性值方式
'Public Property Let iname(ByVal strName As String)
'    mystrName = strName
'End Property

'===取得屬性值方式
'Public Property Get iname() As String
'    iname = mystrName
'End Property


Private ItemsCollection As New Collection
Private ItemsArray()
'   使用者取用時需自訂一個2維振烈來存放myDict資料
'   [1,n]key [2,n]value
'   [n,#]第幾筆資料
Private xxxCount As Long

'20231222 測試發現，collection Add時，大小寫M會被視為同一物;部分特殊字元excel認為不同但collection認為相同
'   這些狀況對ebeDictionary執行起來不算是問題，但可能對使用ebeDictionary的程式會造成影響
'   ==>寫一個變數，將程式遇到的狀況寫入，只要使用ebeDictionary都要在程式端查看此變數傳回的狀況是否影響執行

'   還有其他狀況

'=====在sub或function中使用到此class時的起始設置
Private Sub Class_Initialize()
    xxxCount = 0
    ReDim ItemsArray(2, xxxCount)
End Sub
'=====使用到此class時的sub或function結束時的結束設置
Private Sub Class_Terminate()

End Sub
Public Property Get Count() As Long
'注意，xxxCounts需隨著Remove更新
    Count = xxxCount
End Property

'======方法
Friend Function Add(ByVal Key As String, ByVal Value As String)
'新增資料，key不可重複，value可重複
'會將key & value加入屬性
    'xxxCount
    xxxCount = xxxCount + 1
    'ItemsArray
    ReDim Preserve ItemsArray(2, xxxCount)
    ItemsArray(1, xxxCount) = CStr(Key)
    ItemsArray(2, xxxCount) = CStr(Value)
    'temsCollection
    ItemsCollection.Add CStr(Value), CStr(Key)
End Function
Friend Function copy(ByVal DictName As ebeDictionary)
'將另一dict的值複製到此dict
Dim i As Long
Dim iVal As String, iKey As String
    'xxxCount
    xxxCount = DictName.Count
    'ItemsArray
    ItemsArray = DictName.GetKeysAndValues
    'temsCollection
    Set ItemsCollection = Nothing
    For i = 1 To DictName.Count
        iKey = DictName.GetKeyByPageNo(i)
        iVal = DictName.GetValueByPageNo(i)
        ItemsCollection.Add iVal, iKey
    Next i
End Function
Friend Function Exists(ByVal Key As String) As Boolean
'探測某Key是否存在
Dim iVar As Variant, iAns As Variant
    Exists = True
    iHaveErr = False
    On Error GoTo 880
        iAns = ItemsCollection.Item(Key)
    On Error GoTo 0
    GoTo 881
880
    Exists = False
881
End Function
Friend Function GetKeys(ArrayIndexStartNumber As Integer) As Variant
'取得所有的Key，以一維陣列型態回傳給GegKeys；一維陣列從ArrayIndexStartNumber開始
Dim myKeys()
Dim i As Long, ii As Long
    For i = 1 To UBound(ItemsArray, 2)
        ii = i - (1 - ArrayIndexStartNumber)
        ReDim Preserve myKeys(ii)
        myKeys(ii) = ItemsArray(1, i)
    Next i
    GetKeys = myKeys
End Function

Friend Function GetPageNo(ByVal Key As String) As Long
'取得某key在ItemsArray中的2維號碼
Dim i As Long
Dim myKey
'20231222:原本使用屬性ItemsArray來計算,但發現部分特殊字元excel認為不同但collection認為相同，故將使用ItemsArray的判斷式改成使用collection自己的方式
'    For i = 1 To UBound(ItemsArray, 2)
'        If (ItemsArray(1, i) = Key) Then
'            GetPageNo = i
'            Exit For
'        End If
'    Next i
    
    For i = 1 To xxxCount
        If (ItemsCollection.Item(i) = ItemsCollection.Item(Key)) Then
            GetPageNo = i
            Exit Function
        End If
    Next
'20231222 continue from here
'   繼續修改其他的，將使用ItemsArray的判斷式改成使用collection自己的方式
End Function
Friend Function GetKeyByPageNo(ByVal pageNo As Long) As Variant
'藉由PageNo(ItemArray第二維的號碼)取得key值
Dim i As Long
    GetKeyByPageNo = ItemsArray(1, pageNo)
End Function
Friend Function GetValueByPageNo(ByVal pageNo As Long) As Variant
'藉由PageNo(ItemArray第二維的號碼)取得Value值
Dim i As Long
    GetValueByPageNo = ItemsArray(2, pageNo)
End Function
Friend Function GetValue(ByVal Key As String)
'由key值取得其value
'如果該key不存在collection中，則回傳空值""
    On Error Resume Next
        GetValue = ItemsCollection.Item(Key)
    On Error GoTo 0
End Function
Friend Function GetValues(ArrayIndexStartNumber As Integer) As Variant
'取得所有的Value，以一維陣列型態回傳給GetValues；一維陣列從ArrayIndexStartNumber開始
Dim myValues()
Dim i As Long, ii As Long
    For i = 1 To UBound(ItemsArray, 2)
        ii = i - (1 - ArrayIndexStartNumber)
        ReDim Preserve myValues(ii)
        myValues(ii) = ItemsArray(2, i)
    Next i
    GetValues = myValues
End Function

Friend Function GetKeysAndValues() As Variant
'取得所有的Key和Value，以二維陣列型態回傳給GetKeysAndValues，第1維從1開始
'   [1,n]key [2,n]Value
'   [#,n]共有幾組資料
    GetKeysAndValues = ItemsArray
End Function
Friend Function RemoveAll()
'將所有的key與Value刪除
    'ItemsArray
    Erase ItemsArray
    'ItemsCollection
    Set ItemsCollection = Nothing
    'xxxCount
    xxxCount = 0
End Function
Friend Function Remove(Key As String)
'以key將特定的資料刪除
Dim iArray()
Dim i As Long, iCount As Long
Dim iPageNo As Long
    'ItemsArray
    iCount = 0
    iPageNo = GetPageNo(Key)
    For i = 1 To UBound(ItemsArray, 2)
        If (i <> iPageNo) Then
            iCount = iCount + 1
            ReDim Preserve iArray(2, iCount)
            iArray(1, iCount) = ItemsArray(1, i)
            iArray(2, iCount) = ItemsArray(2, i)
        End If
    Next i
    ItemsArray = iArray
    'xxxCount/ItemsCollection
    If (iCount <> xxxCount) Then
        xxxCount = iCount
        ItemsCollection.Remove (Key)
    End If

End Function
Friend Function SortKey(ascendOrDescend As String)
Dim iArray1(), iArray2()
Dim i As Long, ii As Long, iCount1 As Long, iCount2 As Long
Dim iColllection1 As New Collection, iColllection2 As New Collection
Dim iVar1 As Variant, iVar2 As Variant
Dim iBool1 As Boolean, iDoIt As Boolean
'將key做排序，使用者須輸入"ascend"或"descend"(大小寫不拘)來指定排序方式
    '防呆-使用此function，但變數 ascendOrDescend 內容寫錯
    If (UCase(ascendOrDescend) <> "ASCEND" And UCase(ascendOrDescend) <> "DESCEND") Then
        MsgBox "ebeDictionary的function [SortKey]的輸入變數有誤"
    End If
    '防呆-dict中沒有資料，直接離開function
    If (xxxCount = 0) Then
        Exit Function
    End If
    
    '將ItemsArray的key與其陣列號碼寫進iCollection1供後續取value用 (value = 陣列號碼,key = key)
    For i = 1 To UBound(ItemsArray, 2)
        iColllection1.Add CStr(i), ItemsArray(1, i)
    Next i

    '每2筆資料就在iArray中設為一組(如多出來的自成1組)，放2維入陣列中
    '   [1,n]這組資料中較大的 [2,n]這組資料中較小的 [3,n]目前比較的對象
    '   [n,#]第幾筆資料
    iCount2 = 0
    For i = 1 To UBound(ItemsArray, 2)
        If (i = 1 Or iCount1 = 2) Then
            iCount1 = 1
        Else
            iCount1 = iCount1 + 1
        End If
        '先取出key直，如果是數字就轉為Dbl
        If (iCount1 = 1) Then
            If (IsNumeric(ItemsArray(1, i)) = True) Then
                iVar1 = CDbl(ItemsArray(1, i))
            Else
                iVar1 = ItemsArray(1, i)
            End If
        ElseIf (iCount1 = 2) Then
            If (IsNumeric(ItemsArray(1, i)) = True) Then
                iVar2 = CDbl(ItemsArray(1, i))
            Else
                iVar2 = ItemsArray(1, i)
            End If
        End If
        '將每組數字比大小，大的放[1,n]、小的放[2,n]
        If (iCount1 = 2 Or (iCount1 = 1 And i = UBound(ItemsArray, 2))) Then
            iCount2 = iCount2 + 1
            ReDim Preserve iArray1(3, iCount2)
            If (iCount1 = 2) Then
                If (iVar1 > iVar2) Then
                    iArray1(1, iCount2) = iVar1
                    iArray1(2, iCount2) = iVar2
                Else
                    iArray1(1, iCount2) = iVar2
                    iArray1(2, iCount2) = iVar1
                End If
            ElseIf (iCount1 = 1 And i = UBound(ItemsArray, 2)) Then
                '如果最後一組只有一個值，放到[1,n]
                iArray1(1, iCount2) = iVar1
            End If
            '將要比較得先放到[3,n]
            If (UCase(ascendOrDescend) = "ASCEND") Then
                '昇序,找小的
                If (iArray1(2, iCount2) <> "") Then
                    iArray1(3, iCount2) = iArray1(2, iCount2)
                Else
                    iArray1(3, iCount2) = iArray1(1, iCount2)
                End If
            ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                '降序，找大的
                iArray1(3, iCount2) = iArray1(1, iCount2)
            End If
        End If
    Next i

    '設定一個Collection，Value填入以區分兩兩一組的陣列的第2維號碼,key填入陣列[1,n](key值)
    '   後續比對大小時，已確定順序者從此Collection移除，以供後續程式知道陣列中的該筆資料是否已比對完畢
    For i = 1 To UBound(iArray1, 2)
        iColllection2.Add CStr(i), CStr(iArray1(3, i))
    Next i
    
    '將iArray1中的資料全部比出大小
    '   每次找出一個值(目前最大/最小)，所以總執行次數等於資料筆數
    iCount1 = 0
    For i = 1 To UBound(ItemsArray, 2)
        iVar1 = ""
        For ii = 1 To UBound(iArray1, 2)
            '檢查該陣列是否須執行
            '是否存在
            iBool1 = False
            On Error Resume Next
                iBool1 = iColllection2.Item(CStr(iArray1(3, ii)))
            On Error GoTo 0
            If (iBool1 = False) Then
                GoTo 701
            Else
                If (iVar1 = "") Then
                    iVar1 = iArray1(3, ii)
                Else
                    If (UCase(ascendOrDescend) = "ASCEND") Then
                        '昇序,找小的
                        If (iArray1(3, ii) < iVar1) Then
                            iVar1 = iArray1(3, ii)
                        End If
                    ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                        '降序，找大的
                        If (iArray1(3, ii) > iVar1) Then
                            iVar1 = iArray1(3, ii)
                        End If
                    End If
                    
                End If
            End If
701
        Next ii
        '將目前找到的最大/最小值的key和value寫進陣列
        iCount1 = iCount1 + 1
        ReDim Preserve iArray2(2, iCount1)
        iArray2(1, iCount1) = CStr(iVar1)
        iArray2(2, iCount1) = ItemsArray(2, iColllection1.Item(CStr(iVar1)))
        '重新整理collection 和 重新定義iArray(3,n)
        ii = CLng(iColllection2.Item(CStr(iVar1)))
        iDoIt = False
        If (iArray1(2, ii) <> "") Then
            If (UCase(ascendOrDescend) = "ASCEND") Then
                '昇序,將大的放入[3,n]
                If (iArray1(3, ii) <> iArray1(1, ii)) Then
                    iArray1(3, ii) = iArray1(1, ii)
                    iDoIt = True
                End If
            ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                '降序,將小的放入[3,n]
                If (iArray1(3, ii) <> iArray1(2, ii)) Then
                    iArray1(3, ii) = iArray1(2, ii)
                    iDoIt = True
                End If
            End If
        End If
        If (iDoIt = True) Then
            iColllection2.Add CStr(ii), CStr(iArray1(3, ii))
        End If
        iColllection2.Remove CStr(iVar1)

    Next i
    '將整理後的iArray2資料寫回dict
    '   ItemsArray
    ItemsArray = iArray2
    'xxxCount不變
    'ItemsCollection
    Set ItemsCollection = Nothing
    For i = 1 To UBound(iArray2, 2)
        ItemsCollection.Add CStr(iArray2(2, i)), CStr(iArray2(1, i))
    Next i
End Function
