Attribute VB_Name = "關於Array操作"
Sub 關於Array相關操作()
'將陣列中的字串在一起,注意,陣列是從0開始串
Dim myA(2)
myA(1) = 1
myA(2) = 2
myStr = Join(myA, ",")

'將陣列中的值寫進儲存格
'   陣列所有維度須從0開始輸入
'   可以是多維度的陣列，譬如下列
    Dim iArray1(1, 1)
    iArray1(0, 0) = "00"
    iArray1(0, 1) = "01"
    iArray1(1, 0) = "10"
    iArray1(1, 1) = "11"
    Range("AB1:AC2").Value = iArray1
'   1維陣列說明如下
    '如果range超過1個，陣列只有一筆資料，會將所有range填入該陣列值
    Range("A1:C1").Value = Array("1")
    '如果range超過1個，陣列也超過1個，則會將陣列中的資料依序寫入range；陣列中多餘的會被捨去
    '   A1~C1都被填入相對應的值1~3，最後的4不會被填入
    Range("A1:C1").Value = Array("1", "3", "3", "4")
    '使用Range.value = array，只是用在水平的Range
    '   譬如A1~C1
    '   如果Range是垂直的，須加上Application.Transpose
    Range("A1:A3").Value = Application.Transpose(Array("1", "2", "3", "4"))
End Sub

'======自寫功能(1-1維陣列-1) 對一維陣列排序並移除重複的資料
Public 資料庫_transpose2DimensionArray(myArray)
'將傳入的陣列轉置，譬如 arr(5,100) 變成 arr(100,5)，裡片的值也轉置
Dim i As Long, ii As Long, iNewD1 As Long, iNewD2 As Long
Dim tempArray()
    iNewD1 = UBound(myArray, 2)
    iNewD2 = UBound(myArray, 1)
    ReDim tempArray(iNewD1, iNewD2)
    For i = 0 To iNewD1
        For ii = 0 To iNewD2
            tempArray(i, ii) = myArray(ii, i)
        Next ii
    Next i
    Erase myArray
    myArray = tempArray
End Function


Function 資料庫_myUniqueSort1DimArray(ByVal inputArray As Variant, ascendOrDescend As String)
'注意，判斷是否重複的值如果是純數字時，要使用整數(如LONG)不要使用小數位型態(如DOUBLE)，因為EXCEL在計算DOUBLE時會產生非常小的差異，導致看起來一樣的值對EXCEL來說卻不一樣

'輸入資料:inputArray 原始資料料陣列 / ascendOrDescend 輸入ascend(低至高排序)或descend(高至低排序)
'inputArray中的資料格式: 一維陣列、數字 or 英文 or 混和 (文字型態的數字，譬如"10"，會被當作文字處理)、可有重複資料、可有空格
'排序方式: 依使用者決定要低>高(純數字在前，1 to 9 then A to Z) 或 高<低 (非純數字在前，Z to A then 9 to 1)
'排序後會被剔除者: trim後是空格("")、重複的資料
Dim i As Long, ii As Long, iii As Long
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()

    '排序-同樣大小的排在一起，使得有幾個重複就有幾個空格產生
    ReDim sortWithDuplicateArray(UBound(inputArray, 1))
    For i = 1 To UBound(inputArray, 1)
        iii = 1
        For ii = 1 To UBound(inputArray, 1)
            If (i <> ii) Then
                '由高至低 或 由低至高
                If (UCase(ascendOrDescend) = "ASCEND") Then
                    If (inputArray(i) >= inputArray(ii)) Then
                        iii = iii + 1
                    End If
                ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                    If (inputArray(i) <= inputArray(ii)) Then
                        iii = iii + 1
                    End If
                End If
            End If
        Next ii
        sortWithDuplicateArray(iii) = inputArray(i)
    Next i
    '移除因重複導致的空格
    '   計算有幾個空格
    emptyCount = 0
    For i = 1 To UBound(sortWithDuplicateArray, 1)
        If (Trim(sortWithDuplicateArray(i)) = "") Then
            emptyCount = emptyCount + 1
        End If
    Next i
    '   移除空格
    ReDim sortWithoutDuplicateArray(UBound(sortWithDuplicateArray, 1) - emptyCount)
    ii = 0
    For i = 1 To UBound(sortWithDuplicateArray, 1)
        If (Trim(sortWithDuplicateArray(i)) <> "") Then
            ii = ii + 1
            sortWithoutDuplicateArray(ii) = sortWithDuplicateArray(i)
        End If
    Next i
    myUniqueSort1DimArray = sortWithoutDuplicateArray
End Function

'======自寫功能(1-2維陣列-1) 對二維陣列排序並移除重複的資料
Function 資料庫_myUniqueSort2DimArray(ByVal inputArray As Variant, ascendOrDescend As String, dataInWhichDim As Long, sortByWhichAttribute As Long)
'注意，判斷是否重複的值如果是純數字時，要使用整數(如LONG)不要使用小數位型態(如DOUBLE)，因為EXCEL在計算DOUBLE時會產生非常小的差異，導致看起來一樣的值對EXCEL來說卻不一樣

'輸入資料:inputArray 原始資料料陣列 / ascendOrDescend 輸入ascend(低至高排序)或descend(高至低排序)
'inputArray中的資料格式: 二維陣列、數字 or 英文 or 混和 (文字型態的數字，譬如"10"，會被當作文字處理)、可有重複資料、可有空格
'排序方式: 依使用者決定要低>高(純數字在前，1 to 9 then A to Z) 或 高<低 (非純數字在前，Z to A then 9 to 1)
'排序後會被剔除者: trim後是空格("")、重複的資料
Dim attributeInThisDim
Dim i As Long, ii As Long, iii As Long, iv As Long
Dim iIsEmpty As Boolean
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()
    '取得屬性所在維度方便後續使用
    If (dataInWhichDim = 1) Then
        attributeInThisDim = 2
    ElseIf (dataInWhichDim = 2) Then
        attributeInThisDim = 1
    End If
    '排序-同樣大小的排在一起排進陣列中 > 有幾個重複陣列中就有幾格是空格
    ReDim sortWithDuplicateArray(UBound(inputArray, 1), UBound(inputArray, 2))
    For i = 1 To UBound(inputArray, dataInWhichDim)
        iii = 1

        For ii = 1 To UBound(inputArray, dataInWhichDim)

            If (i <> ii) Then
                '由高至低 或 由低至高
                If (UCase(ascendOrDescend) = "ASCEND") Then
                    If (dataInWhichDim = 1) Then
                        If (inputArray(i, sortByWhichAttribute) >= inputArray(ii, sortByWhichAttribute)) Then
                            iii = iii + 1
                        End If
                    ElseIf (dataInWhichDim = 2) Then
                        If (inputArray(sortByWhichAttribute, i) >= inputArray(sortByWhichAttribute, ii)) Then
                            iii = iii + 1
                        End If
                    End If
                ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                    If (dataInWhichDim = 1) Then
                        If (inputArray(i, sortByWhichAttribute) <= inputArray(ii, sortByWhichAttribute)) Then
                            iii = iii + 1
                        End If
                    ElseIf (dataInWhichDim = 2) Then
                        If (inputArray(sortByWhichAttribute, i) <= inputArray(sortByWhichAttribute, ii)) Then
                            iii = iii + 1
                        End If
                    End If
                End If
            End If
        Next ii
        
        For iv = 0 To UBound(inputArray, attributeInThisDim)
            If (dataInWhichDim = 1) Then
                sortWithDuplicateArray(iii, iv) = inputArray(i, iv)
            ElseIf (dataInWhichDim = 2) Then
                sortWithDuplicateArray(iv, iii) = inputArray(iv, i)
            End If
        Next iv
    Next i
    '移除因重複導致陣列中存在的空格 (每個屬性都是空格才算是)
    '   計算有幾個空格
    emptyCount = 0
    For i = 1 To UBound(sortWithDuplicateArray, dataInWhichDim)
        iIsEmpty = True
        For ii = 0 To UBound(inputArray, attributeInThisDim)
            If (dataInWhichDim = 1) Then
                If (Trim(sortWithDuplicateArray(i, ii)) <> "") Then
                    iIsEmpty = False
                    Exit For
                End If
            ElseIf (dataInWhichDim = 2) Then
                If (Trim(sortWithDuplicateArray(ii, i)) <> "") Then
                    iIsEmpty = False
                    Exit For
                End If
            End If
        Next ii
        If (iIsEmpty = True) Then
            emptyCount = emptyCount + 1
        End If
    Next i
    '   移除空格
    If (dataInWhichDim = 1) Then
        ReDim sortWithoutDuplicateArray(UBound(sortWithDuplicateArray, 1) - emptyCount, UBound(sortWithDuplicateArray, 2))
    ElseIf (dataInWhichDim = 2) Then
        ReDim sortWithoutDuplicateArray(UBound(sortWithDuplicateArray, 1), UBound(sortWithDuplicateArray, 2) - emptyCount)
    End If

    iii = 0
    For i = 1 To UBound(sortWithDuplicateArray, dataInWhichDim)
        '找出空格者 - 由於用來排序的屬性也可能是空值，故設計為所有屬性都 不是 空值才判定為要保留
        iIsEmpty = True
        For ii = 0 To UBound(sortWithDuplicateArray, attributeInThisDim)
            If (dataInWhichDim = 1) Then
                If (Trim(sortWithDuplicateArray(i, ii)) <> "") Then
                    iIsEmpty = False
                    Exit For
                End If
            ElseIf (dataInWhichDim = 2) Then
                If (Trim(sortWithDuplicateArray(ii, i)) <> "") Then
                    iIsEmpty = False
                    Exit For
                End If
            End If
        Next ii
        
        If (iIsEmpty = False) Then
            iii = iii + 1
            
            For iv = 0 To UBound(inputArray, attributeInThisDim)
                If (dataInWhichDim = 1) Then
                    sortWithoutDuplicateArray(iii, iv) = sortWithDuplicateArray(i, iv)
                ElseIf (dataInWhichDim = 2) Then
                    sortWithoutDuplicateArray(iv, iii) = sortWithDuplicateArray(iv, i)
                End If
            Next iv
        End If
    
    Next i
    myUniqueSort2DimArray = sortWithoutDuplicateArray
End Function

'======自寫功能(1-2維陣列-2) 對二維陣列排序-如排序的數性值相同時依原始陣列中前後順序排在一起
Function 資料庫_mySort2DimArray(ByVal inputArray As Variant, ascendOrDescend As String, dataInWhichDim As Long, sortByWhichAttribute As Long)
'注意，判斷是否重複的值如果是純數字時，要使用整數(如LONG)不要使用小數位型態(如DOUBLE)，因為EXCEL在計算DOUBLE時會產生非常小的差異，導致看起來一樣的值對EXCEL來說卻不一樣

'輸入資料:inputArray 原始資料料陣列 / ascendOrDescend 輸入ascend(低至高排序)或descend(高至低排序)
'inputArray中的資料格式: 二維陣列、數字 or 英文 or 混和 (文字型態的數字，譬如"10"，會被當作文字處理)、可有重複資料、可有空格
'排序方式: 依使用者決定要低>高(純數字在前，1 to 9 then A to Z) 或 高<低 (非純數字在前，Z to A then 9 to 1)
'排序後會被剔除者: trim後是空格("")、重複的資料
Dim attributeInThisDim
Dim i As Long, ii As Long, iii As Long, iv As Long
Dim iIsEmpty As Boolean
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()
    '取得屬性所在維度方便後續使用
    If (dataInWhichDim = 1) Then
        attributeInThisDim = 2
    ElseIf (dataInWhichDim = 2) Then
        attributeInThisDim = 1
    End If
    '排序-同樣大小的依照原先前後順序排在一起排進陣列中
    '   *如果要將順序相反，則下面 If (i > ii) Then 改成 If (i < ii) Then
    ReDim sortWithDuplicateArray(UBound(inputArray, 1), UBound(inputArray, 2))
    For i = 1 To UBound(inputArray, dataInWhichDim)
        iii = 1
'If (i = 1022) Then
'iIsEmpty = False
'End If
        '找到排序後的新位置
        For ii = 1 To UBound(inputArray, dataInWhichDim)
'If (ii = 1021) Then
'iIsEmpty = False
'End If
            If (i <> ii) Then
                '由高至低 或 由低至高
                If (UCase(ascendOrDescend) = "ASCEND") Then
                    If (dataInWhichDim = 1) Then
                        If (inputArray(i, sortByWhichAttribute) > inputArray(ii, sortByWhichAttribute)) Then
                            iii = iii + 1
                        ElseIf (inputArray(i, sortByWhichAttribute) = inputArray(ii, sortByWhichAttribute)) Then
                            If (i > ii) Then
                                iii = iii + 1
                            End If
                        End If
                    ElseIf (dataInWhichDim = 2) Then
                        If (inputArray(sortByWhichAttribute, i) > inputArray(sortByWhichAttribute, ii)) Then
                            iii = iii + 1
                        ElseIf (inputArray(sortByWhichAttribute, i) = inputArray(sortByWhichAttribute, ii)) Then
                            If (i > ii) Then
                                iii = iii + 1
                            End If
                        End If
                    End If
                ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                    If (dataInWhichDim = 1) Then
                        If (inputArray(i, sortByWhichAttribute) < inputArray(ii, sortByWhichAttribute)) Then
                            iii = iii + 1
                        ElseIf (inputArray(i, sortByWhichAttribute) = inputArray(ii, sortByWhichAttribute)) Then
                            If (i > ii) Then
                                iii = iii + 1
                            End If
                        End If
                    ElseIf (dataInWhichDim = 2) Then
                        If (inputArray(sortByWhichAttribute, i) < inputArray(sortByWhichAttribute, ii)) Then
                            iii = iii + 1
                        ElseIf (inputArray(sortByWhichAttribute, i) = inputArray(sortByWhichAttribute, ii)) Then
                            If (i > ii) Then
                                iii = iii + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next ii
        '將資料寫入新陣列中的新位置
        For iv = 0 To UBound(inputArray, attributeInThisDim)
            If (dataInWhichDim = 1) Then
                sortWithDuplicateArray(iii, iv) = inputArray(i, iv)
            ElseIf (dataInWhichDim = 2) Then
                sortWithDuplicateArray(iv, iii) = inputArray(iv, i)
            End If
        Next iv
    Next i
    
    mySort2DimArray = sortWithDuplicateArray
End Function

'======自寫功能(2) 對一維陣列做不重複篩選(按照原始先後排序)
Function 資料庫_make1DimArrayUnique(inputArray, arrayStartNum As Long)
Dim tempArray()
Dim i As Long, ii As Long, iCount As Long
Dim iFound As Boolean
    tempArray = inputArray
    Erase inputArray
    iCount = arrayStartNum - 1
    For i = arrayStartNum To UBound(tempArray, 1)
        If (i = 1) Then
            iCount = iCount + 1
            ReDim Preserve inputArray(iCount)
            inputArray(iCount) = tempArray(i)
        Else
            iFound = False
            For ii = arrayStartNum To i - 1
                If (tempArray(i) = tempArray(ii)) Then
                    iFound = True
                    Exit For
                End If
            Next ii
            If (iFound = False) Then
                iCount = iCount + 1
                ReDim Preserve inputArray(iCount)
                inputArray(iCount) = tempArray(i)
            End If
        End If
    Next i
End Function

'======自寫功能(3) 對二維陣列做不重複篩選(按照原始先後排序)
Function 資料庫_make2DimArrayUnique(inputArray, dataInWhichDim As Long, filterInWhichAttribute As Long, arrayStartNum As Long)
Dim attributesTotalNum As Long
Dim tempArray()
Dim i As Long, ii As Long, iCount As Long
Dim iFound As Boolean
    'dataInWhichDim: 1 or 2 ,第n筆資料的n擺在第1維還是第2維
    '   1: myArray(n,1)
    '   2: myArray(1,n)
    'filterInWhichAttribute: 1 to n ,每筆資料第幾個屬性是不重複篩選的依據
    'arrayStartNum: 0 to m ,資料重該維度的第幾個開始放置
    ' <舉例>
    '   第1維是屬性 (共有2個屬性 名子,電話)，第2維是值
    '   myArray(1,0) 第0筆資料的第1個屬性是第一個人的名子
    '   myArray(2,0) 第0筆資料的第2個屬性是第一個人的電話
    '   myArray(1,1) 第1筆資料的第1個屬性是第二個人的名子
    '   myArray(2,1) 第1筆資料的第2個屬性是第二個人的電話
    '   以人名做不重複篩選，則
    '       dataInWhichDim = 2
    '       attributesTotalNum = 2
    '       filterInWhichAttribute = 1
    '       arrayStartNum = 0

    tempArray = inputArray
    Erase inputArray
    
    If (dataInWhichDim = 1) Then
        attributesTotalNum = UBound(tempArray, 2)
    ElseIf (dataInWhichDim = 2) Then
        attributesTotalNum = UBound(tempArray, 1)
    End If
    iCount = arrayStartNum - 1
    For i = arrayStartNum To UBound(tempArray, dataInWhichDim)
        '第1筆資料
        If (i = 1) Then
            iCount = iCount + 1
            If (dataInWhichDim = 1) Then
                ReDim Preserve inputArray(iCount, attributesTotalNum)
                For ii = 1 To attributesTotalNum
                    inputArray(iCount, ii) = tempArray(i, ii)
                Next ii
            ElseIf (dataInWhichDim = 2) Then
                ReDim Preserve inputArray(attributesTotalNum, iCount)
                For ii = 1 To attributesTotalNum
                    inputArray(ii, iCount) = tempArray(ii, i)
                Next ii
            End If
        '其他筆資料
        Else
            iFound = False
            '比對此筆資料和前面的資料是否重複
            For ii = arrayStartNum To i - 1
                If (dataInWhichDim = 1) Then
                    If (tempArray(i, filterInWhichAttribute) = tempArray(ii, filterInWhichAttribute)) Then
                        iFound = True
                        Exit For
                    End If
                ElseIf (dataInWhichDim = 2) Then
                    If (tempArray(filterInWhichAttribute, i) = tempArray(filterInWhichAttribute, ii)) Then
                        iFound = True
                        Exit For
                    End If
                End If
            Next ii
            '沒重複則紀錄
            If (iFound = False) Then
                iCount = iCount + 1
                If (dataInWhichDim = 1) Then
                    ReDim Preserve inputArray(iCount, attributesTotalNum)
                    For ii = 1 To attributesTotalNum
                        inputArray(iCount, ii) = tempArray(i, ii)
                    Next ii
                ElseIf (dataInWhichDim = 2) Then
                    ReDim Preserve inputArray(attributesTotalNum, iCount)
                    For ii = 1 To attributesTotalNum
                        inputArray(ii, iCount) = tempArray(ii, i)
                    Next ii
                End If
            End If
        End If
    Next i
End Function
'======自寫功能(4) 在二維陣列中以兩個條件值找到特定的多筆資料[datas]；比較所有[datas]的某數字屬性的值，看誰的值最大；以該最大值寫回所有[datas]的該屬性
Function 資料庫_getMaxNumberThenOverwriteInArrayWith2Filter(inputArray, dataInWhichDim As Long, dataInWhichAttribute, ByVal filterInWhichAttribute1 As Long, ByVal filterString1 As String, ByVal filterInWhichAttribute2 As Long, ByVal filterString2 As String)
Dim iVal As Double, currentMaxVal As Double
Dim i As Long, ii As Long, iCount As Long
Dim iStr1 As String, iStr2 As String
Dim tempArray()
    'dataInWhichDim: 1 or 2 ,第n筆資料的n擺在第1維還是第2維
    '   1: myArray(n,1)
    '   2: myArray(1,n)
    'dataInWhichAttribute: 1 to n ,每筆資料第幾個屬性是存放要找最大值的數字
    'filterInWhichAttribute: 1 to n ,每筆資料第幾個屬性是不重複篩選的依據
    'filterValue:篩選值
    
    '篩選出陣列中符合條件者並找到最大值
    ReDim tempArray(0)
    iCount = 0
    For i = 0 To UBound(inputArray, dataInWhichDim)
        If (dataInWhichDim = 1) Then
            iStr1 = CStr(inputArray(i, filterInWhichAttribute1))
            iStr2 = CStr(inputArray(i, filterInWhichAttribute2))
            iVal = CDbl(inputArray(i, dataInWhichAttribute))
        Else
            iStr1 = CStr(inputArray(filterInWhichAttribute1, i))
            iStr2 = CStr(inputArray(filterInWhichAttribute2, i))
            iVal = CDbl(inputArray(dataInWhichAttribute, i))
        End If
        If (iStr1 = filterString1 And iStr2 = filterString2) Then
            iCount = iCount + 1
            ReDim Preserve tempArray(iCount)
            tempArray(iCount) = i
            If (iCount = 1) Then
                currentMaxVal = iVal
            Else
                If (iVal > currentMaxVal) Then
                    currentMaxVal = iVal
                End If
            End If
        End If
    Next i
    '用最大值改寫inputArray
    If (UBound(tempArray, 1) <> 0) Then
        For i = 1 To UBound(tempArray, 1)
            If (dataInWhichDim = 1) Then
                inputArray(tempArray(i), dataInWhichAttribute) = currentMaxVal
            ElseIf (dataInWhichDim = 2) Then
                inputArray(dataInWhichAttribute, tempArray(i)) = currentMaxVal
            End If
        Next i
    End If
End Function

Public Function make1DArrayToCsvString_資料庫(myArray)
'將陣列 (一維，從1開始)的值以,串在一起成CSV格式
Dim i As Long
    For i = 1 To UBound(myArray, 1)
        If (i = 1) Then
            make1DArrayToCsvString = myArray(i)
        Else
            make1DArrayToCsvString = make1DArrayToCsvString & "," & myArray(i)
        End If
    Next i
End Function
