Attribute VB_Name = "關於編輯儲存格的值"
Private Sub Worksheet_Change(ByVal Target As Range)
'取得剛剛修改的存儲格的range，並與自訂的一個範圍比較得知 "修改的是否在此特定範圍內"
'此sub只能寫在工作表中
'https://docs.microsoft.com/zh-tw/office/troubleshoot/excel/run-macro-cells-change
'https://docs.microsoft.com/zh-tw/office/vba/api/excel.application.intersect
Dim myRange1 As Range
Dim myRange2 As Range
    Set myRange1 = Range("C10:C11")
    Set myRange2 = Range(Target.address)
        
    If (Application.Intersect(myRange1, myRange2) Is Nothing) Then
        MsgBox "修改的儲存格 不在 指定範圍內"
    Else
        MsgBox "修改的儲存 在 指定範圍內"
    End If
End Sub

Sub 複製貼上_標準語法_資料庫()
    Set mySht = ActiveSheet
    '使用range搭配cell指定複製範圍
    With mySht
        .Range(.Cells(1, 1), .Cells(1, 1)).copy
    End With
    
    '貼上(避免貼上時有任何狀況-如 名稱重複 導致程式暫停)
    Application.DisplayAlerts = False
    mySht.Paste Destination:=mySht.Cells(2, 1)
    Application.DisplayAlerts = True
    
    '一次寫完複製和貼上
    Range("A1").copy Destination:=Range("A2:A3")
End Sub

Sub 複製整張工作表貼上為值_資料庫()


    Cells.Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub 將空格複製上方填滿()
PIVOT_ALL_ROWS = Sheets("樞紐分析").Range("M1").End(xlDown).Row

Range("A1:M" & PIVOT_ALL_ROWS).SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
Sheets("樞紐分析").Cells.Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Cells.Select
    Selection.Replace What:="(空白)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Public Function trimAndReplace2SpaceTo1_資料庫(myString As String)
'將所有相連的多個空格替換成1個空格
Dim iStr1 As String
    iStr1 = Trim(myString)
    iStr1 = Replace(iStr1, "  ", " ")
    If (InStr(1, iStr1, "  ") <> 0) Then
        Call trimAndReplace2SpaceTo1(iStr1)
    Else
        trimAndReplace2SpaceTo1 = iStr1
    End If
End Function

Public Function 資料庫_clearRange(myRange As Range)
'清除儲存格 內容、底色、字色
    With myRange
        .Formula = ""
        '.ClearComments
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
End Function
Public Function 資料庫_varyRangeInSheetFunction1Main(cellOriginalFormula As Variant, copyToThisWB As Workbook, copyToThisWS As Worksheet, varyOnlySameWB As Boolean, varyOnlySameWS As Boolean, varyNoneFixedCol As Boolean, varyFixedCol As Boolean, varyNoneFixedRow As Boolean, varyFixedRow As Boolean, cellOriginalAdress As String, cellNewAdress As String, variedFormula As String, isVariedFormulaOutOfCell As Boolean, Optional allCellColVaryInfoArray As Variant, Optional allCellRowVaryInfoArray As Variant)
'Public Function varyRangeInSheetFunction1Main(cellOriginalFormula As String, copyToThisWB As Workbook, copyToThisWS As Worksheet, varyOnlySameWB As Boolean, varyOnlySameWS As Boolean, varyNoneFixedCol As Boolean, varyFixedCol As Boolean, varyNoneFixedRow As Boolean, varyFixedRow As Boolean, cellOriginalAdress As String, cellNewAdress As String, variedFormula As String, isVariedFormulaOutOfCell As Boolean, Optional allCellColVaryInfoArray As Variant, Optional allCellRowVaryInfoArray As Variant)
'程式功能
'   模仿excel的功能 , 複製儲存格內容時如內容是函數, 函數內的儲存格range會做相對的變動
'algorithm
'   分析函數的組成並將函數當中的range挑出 > 以移動前和移動後的儲存格range差別做計算函數內的range該如何改變 > 將函數內的range更新
'需要以下function
'   convertABCto123 ,convert123toABC ,varyRangeInSheetFunction2Combine
'傳入變數說明
'   cellOriginalFormula：
'       要檢查的儲存格Formula(如非函數會自動跳過檢查)，此func將檢查此函數內的每個range是否要因為 "拷貝前後該函數所在儲存格欄列不同" 而需要做改變
'   copyToThisWB / copyToThisWS：
'       資料要貼在哪個活頁簿 / 工作表
'   varyOnlySameWB / varyOnlySameWS：
'       是否 "複製和貼上的活頁簿/工作表" 名稱相同才處理
'   varyNoneFixedCol / varyFixedCol As Boolean / varyNoneFixedRow / varyFixedRow As Boolean：
'       是否處理 "固定的(有$符號)/不固定的(沒$符號) 的 欄/列"
'   cellOriginalAdress / cellNewAdress：
'       要檢查的函數所在的儲存格在複製時 / 貼上時 的存儲格位址，資料型態為 "A1" 格式
'   variedFormula：
'       改變後的Formula值，呼叫此func者取用此值為最終結果
'   isVariedFormulaOutOfCell：
'       標記處理過程是否發生 "改變後的Col小於A、Row小於1" 的狀況；發生此狀況時variedFormula會等於cellOriginalFormula：呼叫此func者需自行撰寫錯誤訊息
'   [選擇性]allCellColVaryInfoArray() / allCellRowVaryInfoArray()：
'       拷貝前後整張工作表欄列改變的資訊，用以改變函數中的Range用
'           如果該陣列存在則函數中Range的修改以陣列中的資訊處理；如果陣列中找不到該Range的資訊則以儲存格的移動處理Range
'           如果該陣列不存在，則以儲存格的移動處理Range
'       [1,n]複製資料工作表的欄/列號碼(以數字表示) [2,n]同樣的欄/列在貼上資料工作表的欄/列號碼(以數字表示) [3,n]判定欄/列是否改變的資料的值(譬如sn)
'           決定欄列的資料值必須是該欄列的唯一值；如不是則一律以"不明"處理
'       [n,#]第幾筆資料
'           資料有幾欄/列就有幾筆資
'備註
'   此程式不對 "參照的工作表不存在" 做檢查，因為要考慮得太多

Dim funcSplitArray(), fsnArray(), rangeArray()
Dim cellAddressRowMoveValue As Long, cellAddressColMoveValue As Long
Dim i As Long, ii As Long, iii As Long, iv As Long, iStart As Long, iEnd As Long
Dim iQuotationMarkStart As Long, iQuotationMarkEnd As Long
Dim iWordInMid As Long, iNumInMid As Long
Dim iStr1 As String, iStr2 As String, iStr3 As String, iOriginalCN As String
Dim isNum As Boolean, isWord As Boolean, isSymble As Boolean
Dim iCount1 As Long, iCount2 As Long
Dim iOK As Boolean, iFound As Boolean
    '判斷(1)要計算的formula的值為函數 (2)對非固定欄/非固定列/固定欄/固定列 四者其中一者為true 才繼續處理
    If ((Left(cellOriginalFormula, 1) <> "=") Or (varyNoneFixedCol = False And varyNoneFixedRow = False And varyFixedCol = False And varyFixedRow = False)) Then
        variedFormula = cellOriginalFormula
        Exit Function
    Else
        '計算出 row移動值、col移動值
        iStr1 = ""
        iStr2 = ""
        cellAddressRowMoveValue = 0
        cellAddressColMoveValue = 0
        For i = 1 To Len(cellOriginalAdress)
            If (IsNumeric(Mid(cellOriginalAdress, i, 1)) = False) Then
                iStr1 = iStr1 & Mid(cellOriginalAdress, 1, 1)
            Else
                ii = Right(cellOriginalAdress, Len(cellOriginalAdress) - (i - 1))
                Exit For
            End If
        Next i
        For i = 1 To Len(cellNewAdress)
            If (IsNumeric(Mid(cellNewAdress, i, 1)) = False) Then
                iStr2 = iStr2 & Mid(cellNewAdress, 1, 1)
            Else
                iii = Right(cellNewAdress, Len(cellNewAdress) - (i - 1))
                Exit For
            End If
        Next i
        cellAddressRowMoveValue = iii - ii
        ii = convertABCto123(iStr1)
        iii = convertABCto123(iStr2)
        cellAddressColMoveValue = iii - ii
        '將formula解析並放入陣列

        '   找出FSN (File & Sheet Name)
        '       FSN一定開始於運算符號，並結束於 [!]
        '           運算符號: = ( , + - * /
        iStart = 0
        iEnd = 0
        iCount1 = 0
        iCount2 = 0
        iQuotationMarkStart = 0
        iQuotationMarkEnd = 0
        ReDim fsnArray(4, 0)
        For i = 1 To Len(cellOriginalFormula)
            iStr1 = Mid(cellOriginalFormula, i, 1)
            'FSN的值本身也可能包含運算符號,但遇到這些狀況會被單引號(')左右包起
            If (iStr1 = "'") Then
                iCount1 = iCount1 + 1
                If (iCount1 Mod 2 = 1) Then
                    iQuotationMarkStart = i
                    'iQuotationMarkEnd在沒找到前保持和Start一樣
                    iQuotationMarkEnd = i
                Else
                    iQuotationMarkEnd = i
                End If
            End If
            
            If (iStr1 = "=" Or iStr1 = "(" Or iStr1 = "," Or iStr1 = "+" Or iStr1 = "-" Or iStr1 = "*" Or iStr1 = "/") Then
                If (iQuotationMarkStart = 0) Then
                    iStart = i + 1
                End If
            ElseIf (iStr1 = "'") Then
                If (iCount1 Mod 2 = 1) Then
                    iStart = i + 1
                End If
            End If
            If (iStr1 = "!" And iQuotationMarkEnd = 0) Then
                iEnd = i - 1
            ElseIf (iStr1 = "!" And iQuotationMarkEnd <> 0) Then
                iEnd = iQuotationMarkEnd - 1
            End If
            'fsnArray(4,n)
            '   第一維 [1,n]此筆資料從哪一個字開始 [2,n]此筆資料從哪一個字結束 [3,n]此筆資料的值 [4,n]比對FSN規則後此值是否修改
            '   第二維 [n,#]第幾筆資料
            If (iStart <> 0 And iEnd <> 0) Then
                iCount2 = iCount2 + 1
                ReDim Preserve fsnArray(4, iCount2)
                fsnArray(1, iCount2) = iStart
                fsnArray(2, iCount2) = iEnd
                fsnArray(3, iCount2) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '如果不是用 ' 開頭的狀況，避免有多餘空格所以trim
                If (iQuotationMarkStart = 0) Then
                    fsnArray(3, iCount2) = Trim(fsnArray(3, iCount2))
                End If
                
                iStart = 0
                iEnd = 0
                iQuotationMarkStart = 0
                iQuotationMarkEnd = 0
            End If
        Next i

        '   找出RANGE
        '       range的開頭一定是符號、結尾一定是符號或formula最末；range的組成一定為 英文(+英文...)+數字(+數字....)
        iCount1 = 0
        iWordInMid = 0
        iNumInMid = 0
        iStart = 0
        iEnd = 0
        ReDim rangeArray(5, 0)
        For i = 1 To Len(cellOriginalFormula)
            '不在fsn範圍內的才執行
            iOK = True
            If (UBound(fsnArray, 2) = 0) Then
                iOK = True
            Else
                For ii = 1 To UBound(fsnArray, 2)
                    If (((i >= fsnArray(1, ii)) And (i <= fsnArray(2, ii)))) Then
                        iOK = False
                    End If
                Next ii
            End If
            If (iOK = True) Then
                isSymble = False
                iStr1 = Mid(cellOriginalFormula, i, 1)
                'iEnd沒被找到前須隨時歸零
                iEnd = 0
    
                '判斷每一個字是否為英文或數字，如都不是則為符號
                iStr2 = UCase(iStr1)
                If (iStr2 = "A" Or iStr2 = "B" Or iStr2 = "C" Or iStr2 = "D" Or iStr2 = "E" Or iStr2 = "F" Or iStr2 = "G" Or iStr2 = "H" Or iStr2 = "I" Or iStr2 = "J" Or iStr2 = "K" Or iStr2 = "L" Or iStr2 = "M" Or iStr2 = "N" Or iStr2 = "O" Or iStr2 = "P" Or iStr2 = "Q" Or iStr2 = "R" Or iStr2 = "S" Or iStr2 = "T" Or iStr2 = "U" Or iStr2 = "V" Or iStr2 = "W" Or iStr2 = "X" Or iStr2 = "Y" Or iStr2 = "Z") Then
                    isWord = True
                Else
                    isWord = False
                End If
                '   $符號視為英文文字
                If (isWord = True Or iStr1 = "$") Then
                    iWordInMid = i
                Else
                    isNum = IsNumeric(iStr1)
                    If (isNum = True) Then
                        If (iWordInMid = 0) Then
                            iNumInMid = 0
                        Else
                            iNumInMid = i
                        End If
                    Else
                        isSymble = True
                    End If
                End If
                If (isSymble = True) Then
                    '最後一個字是符號
                    If (iNumInMid = 0) Then
                        iStart = i + 1
                        iWordInMid = 0
                    ElseIf (iWordInMid <> 0 And iNumInMid <> 0) Then
                        iEnd = i - 1
                        iWordInMid = 0
                        iNumInMid = 0
                    End If
                ElseIf (i = Len(cellOriginalFormula) And isNum = True And iWordInMid <> 0 And iNumInMid <> 0) Then
                    '最後一個字不是符號且是數字
                    iEnd = i
                    iWordInMid = 0
                    iNumInMid = 0
                End If
    
                
                '抓取range的資料寫進陣列
                'rangeArray(3,n)
                '   第一維 [1,n]此筆資料從哪一個字開始 [2,n]此筆資料從哪一個字結束 [3,n]此筆資料的值(舊值；修改後覆蓋此值) [4,n]比對FSN規則後此值是否修改 [5,n]此range屬於哪一組FSN
                '   第二維 [n,#]第幾筆資料
                If (iStart <> 0 And iEnd <> 0) Then
                    iCount1 = iCount1 + 1
                    ReDim Preserve rangeArray(5, iCount1)
                    rangeArray(1, iCount1) = iStart
                    rangeArray(2, iCount1) = iEnd
                    rangeArray(3, iCount1) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                    iStart = i + 1
                End If
            End If
        Next i
    End If
    '如果有任何range資料才繼續處理
    If (UBound(rangeArray, 2) = 0) Then
        variedFormula = cellOriginalFormula
    Else
        '由FSN規則判定fsnArray的各組是否不需修改
        '   找出參照值是否包含檔名、檔名是否與copyTo相同；工作表名稱是否與copyTo相同
        For i = 1 To UBound(fsnArray, 2)
            iOK = True
            '參照值中的路徑與檔名
            If (varyOnlySameWB = True) Then
                If (InStr(fsnArray(3, i), "\") <> 0) Then
                    '路徑
                    iStr1 = Left(fsnArray(3, i), InStr(fsnArray(3, i), "\[") - 1)
                    If (iStr1 <> copyToThisWB.Path) Then
                        iOK = False
                    End If
                    '檔名
                    ii = InStr(fsnArray(3, i), "[") + 1
                    iii = InStr(fsnArray(3, i), "]") - 1
                    iStr2 = Mid(fsnArray(3, i), ii, iii - ii + 1)
                    If (iStr2 <> copyToThisWB.Name) Then
                        iOK = False
                    End If
                End If
            End If
            '參照值中的工作表名
            If (varyOnlySameWS = True) Then
                '參照值中的工作表名
                ii = InStr(fsnArray(3, i), "]")
                If (ii <> 0) Then
                    '取值-參照值包含有檔名在奈
                    iStr1 = Right(fsnArray(3, i), Len(fsnArray(3, i)) - ii)
                Else
                    '取值-參照值不含檔名
                    iStr1 = fsnArray(3, i)
                End If
                If (iStr1 <> copyToThisWS.Name) Then
                    iOK = False
                End If
            End If
            
            fsnArray(4, i) = iOK
        Next i
        '判斷各range資料是否要修改
        If (UBound(fsnArray, 2) = 0) Then
            'fsnArray沒值時全都判斷為要處理
            For i = 1 To UBound(rangeArray, 2)
                rangeArray(4, i) = True
                rangeArray(5, i) = "NA"
            Next i
        Else
            For i = 1 To UBound(rangeArray, 2)
                '將各range是屬於哪個fsn先找出後，將range設定為與fsn相同的值(在此不考慮varyXXXRow/varyXXXCol的影響)
                
                iStr1 = Mid(cellOriginalFormula, rangeArray(1, i) - 1, 1)
                If (iStr1 = ",") Then
                    '前一個字是逗號, ，則不屬於任一個fsn
                    rangeArray(4, i) = True
                    rangeArray(5, i) = "NA"
                ElseIf (iStr1 = ":") Then
                    '前一個字是冒號 : ，則該range和前一個range屬於同一個fsn
                    rangeArray(4, i) = rangeArray(4, i - 1)
                    rangeArray(5, i) = rangeArray(5, i - 1)
                ElseIf (iStr1 = "!") Then
                    '前一個字是 "!"，則則屬於該fsn
                    iStr2 = Mid(cellOriginalFormula, rangeArray(1, i) - 2, 1)
                    If (iStr2 = "'") Then
                        ii = rangeArray(1, i) - 3
                    Else
                        ii = rangeArray(1, i) - 2
                    End If
                    For iii = 1 To UBound(fsnArray, 2)
                        If (fsnArray(2, iii) = ii) Then
                            rangeArray(4, i) = fsnArray(4, iii)
                            rangeArray(5, i) = iii
                            Exit For
                        End If
                    Next iii
                Else
                    '都不符合以上的狀況，代表 "多個RANGE中其他的有fsn、這個沒有"，要處理
                    rangeArray(4, i) = True
                    rangeArray(5, i) = "NA"
                End If
            Next i
        End If
        
    
    
''======test用
'thisworkbook.Sheets("test").Cells.ClearContents
'For i = 1 To UBound(fsnArray, 2)
'    For ii = 1 To UBound(fsnArray, 1)
'        thisworkbook.Sheets("test").Cells(i, ii).Value = fsnArray(ii, i)
'    Next ii
'Next i
'For i = 1 To UBound(rangeArray, 2)
'    For ii = 1 To UBound(rangeArray, 1)
'        thisworkbook.Sheets("test").Cells(i, ii + 5).Value = rangeArray(ii, i)
'    Next ii
'Next i
''============
    
        '將是range的row/column依照儲存格Range變動做計算
        isVariedFormulaOutOfCell = False
        For i = 1 To UBound(rangeArray, 2)
            '該range需要修要修改才進一步判斷range內容是否要修正
            If (rangeArray(4, i) = True) Then
                iStr1 = ""
                '欄值找出
                If (Left(rangeArray(3, i), 1) = "$") Then
                    '如果欄值有固定(第一個字為$)，從第二個字開始判斷是否為欄值
                    iii = 2
                Else
                    '如果欄值沒有固定，從第一個字開始判斷
                    iii = 1
                End If
                For ii = iii To Len(rangeArray(3, i))
                    If (IsNumeric(Mid(rangeArray(3, i), ii, 1)) = True Or Mid(rangeArray(3, i), ii, 1) = "$") Then
                        Exit For
                    Else
                        iStr1 = iStr1 & Mid(rangeArray(3, i), ii, 1)
                    End If
                Next ii
                '記錄下原始欄名，供後續使用
                iOriginalCN = iStr1
                '   如果符合需要重新計算的條件，則計算
                '       varyNoneFixedCol/varyFixedCol 條件是否符合
                If ((varyNoneFixedCol = True And iii = 1) Or (varyFixedCol = True And iii = 2)) Then
                    '如果選擇性陣列有輸入則以陣列內容計算；陣列中查無資料 或 沒輸入陣列則以cellAddressColMoveValue計算
                    iv = convertABCto123(iOriginalCN)
                    iFound = False
                    If (IsMissing(allCellColVaryInfoArray) = False) Then
                        If (UBound(allCellColVaryInfoArray, 2) <> 0) Then
                            For ii = 1 To UBound(allCellColVaryInfoArray, 2)
                                If (allCellColVaryInfoArray(1, ii) = iv) Then
                                    If (allCellColVaryInfoArray(2, ii) <> 0) Then
                                        iStr2 = convert123toABC(allCellColVaryInfoArray(2, ii))
                                        iFound = True
                                    End If
                                    Exit For
                                End If
                            Next ii
                        End If
                    End If
                    If (iFound = False) Then
                        '如果改變後的欄值小於A，終止所有變動計算
                        If ((iv + cellAddressColMoveValue) < 1) Then
                            isVariedFormulaOutOfCell = True
                            GoTo 881
                        Else
                            iStr2 = convert123toABC(iv + cellAddressColMoveValue)
                        End If
                    End If
                Else
                    iStr2 = iStr1
                End If
                '   如果需要，補回"$"
                If (iii = 2) Then
                    iStr2 = "$" & iStr2
                    iOriginalCN = "$" & iOriginalCN
                End If
        
        
                '列值找出
                iStr1 = Right(rangeArray(3, i), Len(rangeArray(3, i)) - Len(iOriginalCN))
                If (Left(iStr1, 1) = "$") Then
                    '如果列值有固定(第一個字為$)
                    ii = 1
                Else
                    '如果列值沒有固定
                    ii = 0
                End If
                iii = CLng(Right(rangeArray(3, i), Len(rangeArray(3, i)) - Len(iOriginalCN) - ii))
                '   如果符合需要重新計算的條件，則計算
                '       varyNoneFixedRow/varyFixedRow 條件是否符合
                If ((varyNoneFixedRow = True And ii = 0) Or (varyFixedRow = True And ii = 1)) Then
                    '如果選擇性陣列有輸入則以陣列內容計算；陣列中查無資料 或 沒輸入陣列則以cellAddressRowMoveValue計算
                    iFound = False
                    If (IsMissing(allCellRowVaryInfoArray) = False) Then
                        If (UBound(allCellRowVaryInfoArray, 2) <> 0) Then
                            For iv = 1 To UBound(allCellRowVaryInfoArray, 2)
                                If (allCellRowVaryInfoArray(1, iv) = iii) Then
                                    If (allCellRowVaryInfoArray(2, iv) <> 0) Then
                                        iii = allCellRowVaryInfoArray(2, iv)
                                        iFound = True
                                    End If
                                    Exit For
                                End If
                            Next iv
                        End If
                    End If
                    
                    If (iFound = False) Then
                    '如果改變後的列值小於1，終止所有變動計算
                        iii = iii + cellAddressRowMoveValue
                        If (iii < 1) Then
                            isVariedFormulaOutOfCell = True
                            GoTo 881
                        End If
                    End If
                End If
                '   如果需要，補回"$"
                If (ii = 1) Then
                    iStr3 = "$" & iii
                Else
                    iStr3 = iii
                End If
                rangeArray(3, i) = iStr2 & iStr3
            End If
        Next i
881
        '   將陣列內容重新串在一起回傳
        Call varyRangeInSheetFunction2Combine(0, 0, cellOriginalFormula, variedFormula, fsnArray, rangeArray, 1, 1)
        
         
    End If
    
''======test用
'thisworkbook.Sheets("test").Cells(1, 16) = cellAddressColMoveValue
'thisworkbook.Sheets("test").Cells(2, 16) = cellAddressRowMoveValue
'For i = 1 To UBound(rangeArray, 2)
'    For ii = 1 To UBound(rangeArray, 1)
'        thisworkbook.Sheets("test").Cells(i, ii + 10).Value = rangeArray(ii, i)
'    Next ii
'Next i
'thisworkbook.Sheets("test").Cells(1, 17).Value = "'" & cellOriginalFormula
'thisworkbook.Sheets("test").Cells(2, 17).Value = "'" & variedFormula
'thisworkbook.Sheets("test").Cells(1, 18).Value = isVariedFormulaOutOfCell
''============

End Function

Public Function 資料庫_varyRangeInSheetFunction2Combine(ByVal lastArrayWordCounter As Long, ByVal currentWordCounter As Long, ByVal originalFormula As String, ByRef newFormula As Variant, fsnArray, rangeArray, ByVal array1Counter As Long, ByVal array2Counter As Long)
'將sheet function的值拼回去成一串
    'main呼叫此func時 lastArrayWordCounter,currentWordCounter 須為0，fsnArray,rangeArray 須為1
     
    
    'fsnArray/array2須符合以下格式
    '   第一維 [1,n]此筆資料從哪一個字開始 [2,n]此筆資料從哪一個字結束 [3,n]此筆資料的值
    '   第二維 [n,#]第幾筆資料
    
    'algorithm:數數從1開始，每數一個數就依序檢查array1/2的[1,n1]/[1,n2]是否符合該數值
    '   如是
    '   (1)將中間的空白([上一次符合的陣列的[2,n]的值+1]到[目前的數數數值-1])以原來的formula字串填入新formula字串中
    '   (2)將陣列的值[3,n]拚入新formula字串
    '   如否(arry1/2都不符合)則繼續數
Dim skipRest As Boolean
    currentWordCounter = currentWordCounter + 1
    skipRest = False
    '陣列仍有資料才繼續執行
    If (array1Counter <= UBound(fsnArray, 2)) Then
        If (fsnArray(1, array1Counter) = currentWordCounter) Then
            newFormula = newFormula & Mid(originalFormula, lastArrayWordCounter + 1, currentWordCounter - lastArrayWordCounter - 1)
            newFormula = newFormula & fsnArray(3, array1Counter)
            lastArrayWordCounter = fsnArray(2, array1Counter)
            currentWordCounter = fsnArray(2, array1Counter)
            array1Counter = array1Counter + 1

            skipRest = True
        End If
    End If
    If (skipRest = False) Then
        If (array2Counter <= UBound(rangeArray, 2)) Then
            If (rangeArray(1, array2Counter) = currentWordCounter) Then
                newFormula = newFormula & Mid(originalFormula, lastArrayWordCounter + 1, currentWordCounter - lastArrayWordCounter - 1)
                newFormula = newFormula & rangeArray(3, array2Counter)
                lastArrayWordCounter = rangeArray(2, array2Counter)
                currentWordCounter = rangeArray(2, array2Counter)
                array2Counter = array2Counter + 1

            End If
        End If
    End If
    If (currentWordCounter < Len(originalFormula)) Then
        Call varyRangeInSheetFunction2Combine(lastArrayWordCounter, currentWordCounter, originalFormula, newFormula, fsnArray, rangeArray, array1Counter, array2Counter)
    ElseIf (currentWordCounter = Len(originalFormula) And currentWordCounter <> lastArrayWordCounter) Then
        newFormula = newFormula & Right(originalFormula, currentWordCounter - lastArrayWordCounter)
    End If
End Function

Public Function 資料庫_getAllRangesInfoInFormula(ByVal cellOriginalFormula As String, nowSheettName As String) As Variant
'   getAllRangesInfoInFormula回傳一個2維陣列
'       [1,0]   [2,0]   [3,n]   [4,0]是否所有欄位都屬於目前檔案 [5,0]是否所有欄位目前檔案及目前工作表(注意，有!也可能是同一張工作表)
'       [1,n]欄名(A,B,C...) [2,n]欄號碼(1,2,3...) [3,n]列號碼 [4,n]是否屬於目前檔案 [5,n]同檔案，但是否屬於目前工作表(注意，有!也可能是同一張工作表)
'       [n,#]第幾筆資料
'   如果參照範圍是整列(1:1)，[1,n]回傳"ALL"  [2,n]回傳 "1234567890" [3,n]回傳該列號碼
'   如果參照範圍是整欄(A:A)，[1,n]回傳該欄名 [2,n]回傳該欄號碼      [3,n]回傳 "1234567890"
    
'*注意，使用此fun時需先做好防呆，確定使用者輸入的是合法的函數、不是excel系統錯誤值(#REF!之類的)，否則可能會發生無法預期的錯誤
'   iVar = IsError(chkThisWS.Cells(i, ii).Value)


    '將一個儲存格內的函數有參照到的欄位都列出
    '   輸入的變數
    '       cellOriginalFormula 要處理的儲存格內的formula(不是函數也沒關係)
    '       nowSheettName 工作表名稱，會判斷參照範圍是否是在此工作表內
    
    '   getAllRangesInfoInFormula回傳一個2維陣列
    '       [1,0]   [2,0]   [3,n]   [4,0]是否所有欄位都屬於目前檔案 [5,0]是否所有欄位目前檔案及目前工作表(注意，有!也可能是同一張工作表)
    '       [1,n]欄名(A,B,C...) [2,n]欄號碼(1,2,3...) [3,n]列號碼 [4,n]是否屬於目前檔案 [5,n]同檔案，但是否屬於目前工作表(注意，有!也可能是同一張工作表)
    '       [n,#]第幾筆資料
    
    '   如果參照範圍是整列(1:1)，將欄名寫上"ALL",欄號碼"1234567890"
    '   如果參照範圍是整欄(A:A)，將列號碼寫上"1234567890"
    '   如果儲存格的值不包含欄位資訊，getAllRangesInfoInFormula回傳1維是5，二維只有0的陣列，且(4,0) = True、(5,0) = True
    'algorithm
    '   分析函數時拆出檔案路徑、檔案名稱、參照範圍的方式
    '       將是 "檔案路徑 & 檔案名稱" 的字串挑出，紀錄起始位置、結束位置
    '       將是 "參照範圍" 的字串挑出，紀錄起始位置、結束位置
    '       比對 "參照範圍" 的起始位置 與 "檔案路徑 & 檔案名稱" 的結束位置，得知該參照範圍屬於哪個檔案、哪個工作表
    '       得知該筆資料是否為一個範圍的頭/尾，或者只是一個單獨的儲存格
    '   將分析結果的儲存格資料一個一個展開寫入陣列，此函數=該陣列
Dim funcSplitArray(), fsnArray(), cellArray()
Dim cellAddressRowMoveValue As Long, cellAddressColMoveValue As Long
Dim colName As String, colNo As Long, rowNo As Long
Dim endColName As String, endColNo As Long, endRowNo As Long
Dim i As Long, ii As Long, iii As Long, iv As Long, iStart As Long, iEnd As Long, iPureNumStart As Long
Dim iQuotationMarkStart As Long, iQuotationMarkEnd As Long
Dim iWordInMid As Long, iNumInMid As Long
Dim iStr1 As String, iStr2 As String, iStr3 As String, iOriginalCN As String
Dim isNum As Boolean, isWord As Boolean, isDolarSymble As Boolean, isColonSymbol As Boolean, isOtherSymble As Boolean
Dim iCount1 As Long, iCount2 As Long
Dim iNotFsn As Boolean, iFound1 As Boolean, iFound2 As Boolean, ibool1 As Boolean, ibool2 As Boolean
Dim iColonAhead As Boolean
Dim iArray()
Dim iColName As String, iColNo As Long, iRowNo As Long
Dim iCharTypeDict As New ebeDictionary, iFsnCharPosDict As New ebeDictionary
    
   
    '防呆-要計算的formula的值為函數才繼續處理
    If (Left(cellOriginalFormula, 1) <> "=") Then
        ReDim iArray(5, 0)
        iArray(4, 0) = True
        iArray(5, 0) = True
        GoTo 999
    Else
        '將formula解析並放入陣列

        '   找出FSN (File & Sheet Name)
        '       FSN一定開始於運算符號，並結束於 [!]
        '           運算符號: = ( , + - * /
        iStart = 0
        iEnd = 0
        iCount1 = 0
        iCount2 = 0
        iQuotationMarkStart = 0
        iQuotationMarkEnd = 0
        ReDim fsnArray(5, 0)
        For i = 1 To Len(cellOriginalFormula)
            iStr1 = Mid(cellOriginalFormula, i, 1)
            'FSN的值本身也可能包含運算符號,但遇到這些狀況會被單引號(')左右包起
            If (iStr1 = "'") Then
                iCount1 = iCount1 + 1
                If (iCount1 Mod 2 = 1) Then
                    iQuotationMarkStart = i
                    'iQuotationMarkEnd在沒找到前保持和Start一樣
                    iQuotationMarkEnd = i
                Else
                    iQuotationMarkEnd = i
                End If
            End If
            
            If (iStr1 = "=" Or iStr1 = "(" Or iStr1 = "," Or iStr1 = "+" Or iStr1 = "-" Or iStr1 = "*" Or iStr1 = "/") Then
                If (iQuotationMarkStart = 0) Then
                    iStart = i + 1
                End If
            ElseIf (iStr1 = "'") Then
                If (iCount1 Mod 2 = 1) Then
                    iStart = i + 1
                End If
            End If
            If (iStr1 = "!" And iQuotationMarkEnd = 0) Then
                iEnd = i - 1
            ElseIf (iStr1 = "!" And iQuotationMarkEnd <> 0) Then
                iEnd = iQuotationMarkEnd - 1
            End If
            'fsnArray(4,n)
            '   第一維 [1,n]此筆資料從哪一個字開始 [2,n]此筆資料從哪一個字結束 [3,n]此筆資料的值 [4,n]是否屬於同一個檔案 [5,n]是否屬於同一個工作表
            '   第二維 [n,#]第幾筆資料
            If (iStart <> 0 And iEnd <> 0) Then
                iCount2 = iCount2 + 1
                ReDim Preserve fsnArray(5, iCount2)
                fsnArray(1, iCount2) = iStart
                fsnArray(2, iCount2) = iEnd
                fsnArray(3, iCount2) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '如果不是用 ' 開頭的狀況，避免有多餘空格所以trim
                If (iQuotationMarkStart = 0) Then
                    fsnArray(3, iCount2) = Trim(fsnArray(3, iCount2))
                End If
                
                iStart = 0
                iEnd = 0
                iQuotationMarkStart = 0
                iQuotationMarkEnd = 0
            End If
        Next i
        
        '   找出每一個cell/整欄/整列的資料在cellOriginalFormula的字元位置
        '       資料組成方式
        '           cell 的組成
        '               英文(+英文...)+數字(+數字....)
        '           整欄 的組成
        '               英文:英文
        '           整列 的組成
        '               數字:數字
        '           演算法
        '               準備變數...
        '                   iDict來寫入各字元的資料種類
        '                       [key]字元號碼 [value]
        '                                       A(英文字元)
        '                                       1(數字字元)
        '                                       $($字元)
        '                                       :(:字元)
        '                                       =(=字元)
        '                                       other(其他符號字元[包含空格])
        '                   iStart來寫入一個 cell/整欄/整列 的資料在formula裡的起始字元位置
        '                   iEnd來寫入一個 cell/整欄/整列 的資料在formula裡的結束字元位置
        '                   iColonAhead來寫入 "前面有 :字元 ，且尚未找到字串" 時為True
        '                   iFound來標示整個字元的起始結束位置都已找到
        '                   iPureNumStart來寫入 前一個是符號字元 不是英文、不是數字、不是$字元、不是:字元 的數字資料在formula裡的起始字元位置
        '
        '                   !! 任何時候，iFound為True時，紀錄此Cell / 整欄 / 整列 的值；執行紀錄後，然後將iStart & iEnd & iPureNumStart 設為0、iColonAhead 為false
        '
        '               將字元種類分為6種
        '                   (1)英文字元 (2)數字字元 (3)$字元 (4):字元 (5)=字元 (6)其他符號字元[包含空格]
        '               將整個字串每一個字元的種類分辨出
        '               將字串以逐一字元檢視的方式找出cell值
        '                   前置檢查
        '                       第1個字元是 =字元 才處理
        '                       字元的位置不在fsn的範圍內才處理
        '                   遇到 其他符號字元[包含空格]，如果...
        '                       a.iStart & iEnd 都不為0  ---> 將 iFound 設為True
        '                       b.iStart 為0 且 iEnd 不為0，將 iEnd 設為0、iPureNumStart 設為0、iColonAhead 為false
        '                       c.iStart & iEnd 都為0，將 iPureNumStart 設為0、iColonAhead 為false
        '                       d.iStart 不為0 且 iEnd 為0，將 iStart設為0、iEnd 設為0、iPureNumStart 設為0、iColonAhead 為false
        '                   遇到 $字元，如果...
        '                       a.iStart 為 0，如果...
        '                           a.前一個字元是 :字元 ，將 iColonAhead 設為True，將 iStart 設為 目前字元位置，將 iEnd 設為 目前字元位置
        '                           b.前一個字元不是 :字元， iStart 設為 目前字元位置
        '                   遇到 :字元，如果...
        '                       a.iPureNumStart為0，將iEnd設為前一個字元 ---> 將 iFound 設為True
        '                       b.iPureNumStart不為0，將iStart設為iPureNumStart的值，將iEnd設為前一個字元 ---> 將 iFound 設為True
        '                   遇到 英文字元，如果...
        '                       iStart 為 0，如果...
        '                               a.前一個字元是 :字元 ，將 iColonAhead 設為True，將 iStart 設為 目前字元位置，將 iEnd 設為 目前字元位置
        '                               b.前一個字元不是 :字元， iStart 設為 目前字元位置
        '                       iStart 不為 0，且iColonAhead 為True ， iEnd 設為 目前字元位置
        '                   遇到 數字字元，如果...
        '                       a.iStart 為 0 且...
        '                           a-1.iPureNumStart為 0...
        '                               a-1-1.前一個字元是 :字元 ，將 iStart 設為 目前字元位置，將 iEnd 設為 目前字元位置
        '                               a-1-2.前一個字元是 其他符號字元[包含空格] ，將 iPureNumStart 設為 目前字元位置
        '                           a-2.iPureNumStart不為 0 且 下一個字元是 其他符號字元[包含空格]，將iEnd設為目前字元、將 iStart 設為iPureNumStart的值
        '                       b.iStart 不為 0，將iEnd設為目前字元
        '                   最後檢查，不論目前的字元是什麼，只要目前字元是formula最後一個字元，則如果...
        '                       目前字元是 英文字元 ，且 iStart 有值，iEnd沒有值，將iEnd設為目前字元  ---> 將iFound 設為True
        '                       目前字元是 數字字元 ，且 iStart & iEnd 都有值，將iEnd設為目前字元  ---> 將iFound 設為True
         
        
        
        
        '   各字元的資料種類寫入dict
        For i = 1 To Len(cellOriginalFormula)
            iStr1 = UCase(Mid(cellOriginalFormula, i, 1))
            If (iStr1 = "A" Or iStr1 = "B" Or iStr1 = "C" Or iStr1 = "D" Or iStr1 = "E" Or iStr1 = "F" Or iStr1 = "G" Or iStr1 = "H" Or iStr1 = "I" Or iStr1 = "J" Or iStr1 = "K" Or iStr1 = "L" Or iStr1 = "M" Or iStr1 = "N" Or iStr1 = "O" Or iStr1 = "P" Or iStr1 = "Q" Or iStr1 = "R" Or iStr1 = "S" Or iStr1 = "T" Or iStr1 = "U" Or iStr1 = "V" Or iStr1 = "W" Or iStr1 = "X" Or iStr1 = "Y" Or iStr1 = "Z") Then
                iStr2 = "A"
            ElseIf (IsNumeric(iStr1) = True) Then
                iStr2 = "1"
            ElseIf (iStr1 = "$") Then
                iStr2 = "$"
            ElseIf (iStr1 = ":") Then
                iStr2 = ":"
            ElseIf (iStr1 = "=") Then
                iStr2 = "="
            Else
                iStr2 = "OTHER"
            End If
            iCharTypeDict.Add i, iStr2
        Next i
        '   fsn的位置寫入dict
        For i = 1 To UBound(fsnArray, 2)
            For ii = fsnArray(1, i) To fsnArray(2, i)
                iFsnCharPosDict.Add ii, ii
            Next ii
        Next i
        '   將字串以逐一字元檢視的方式找出cell值
        iCount1 = 0
        iStart = 0
        iEnd = 0
        iPureNumStart = 0
        iFound1 = False
        ReDim cellArray(6, 0)
        For i = 1 To Len(cellOriginalFormula)
            '字元的位置不在fsn的範圍內才處理
            iNotFsn = False
            If (UBound(fsnArray, 2) = 0) Then
                iNotFsn = True
            Else
                If (iFsnCharPosDict.Exists(i) = False) Then
                    iNotFsn = True
                End If
            End If
            If (iNotFsn = True) Then
                '開始解析
                If (iCharTypeDict.GetValue(i) <> "=") Then
                    If (iCharTypeDict.GetValue(i) = "OTHER") Then
                        '遇到 其他符號字元[包含空格]，如果...
                        If (iStart <> 0 And iEnd <> 0) Then
                            'a.iStart & iEnd 都不為0  ---> 將 iFound1 設為True
                            iFound1 = True
                        ElseIf (iStart = 0) Then
                            'b.iStart 為0 且 iEnd 不為0，將 iEnd 設為0、iPureNumStart 設為0、iColonAhead 為false
                            'iStart & iEnd 都為0，將 iPureNumStart 設為0、iColonAhead 為false
                            iEnd = 0
                            iPureNumStart = 0
                            iColonAhead = False
                        ElseIf (iStart <> 0 And iEnd = 0) Then
                            'd.iStart 不為0 且 iEnd 為0，將 iStart設為0、iEnd 設為0、iPureNumStart 設為0、iColonAhead 為false
                            iStart = 0
                            iEnd = 0
                            iPureNumStart = 0
                            iColonAhead = False
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "$") Then
                        '遇到 $字元，如果iStart 為 0
                        If (iStart = 0) Then
                            If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                'a.前一個字元是 :字元 ，將 iColonAhead 設為True，將 iStart 設為 目前字元位置，將 iEnd 設為 目前字元位置
                                iColonAhead = True
                                iStart = i
                                iEnd = i
                            ElseIf (iCharTypeDict.GetValue(i - 1) <> ":") Then
                                'b.前一個字元不是 :字元， iStart 設為 目前字元位置
                                iStart = i
                            End If
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = ":") Then
                        '遇到 :字元，如果...
                        If (iPureNumStart = 0) Then
                            'a.iPureNumStart為0 > 將iEnd設為前一個字元 ---> 將 iFound1 設為True
                            iEnd = i - 1
                            iFound1 = True
                        ElseIf (iPureNumStart <> 0) Then
                            'b.iPureNumStart不為0 > 將iStart設為iPureNumStart的值，將iEnd設為前一個字元 ---> 將 iFound1 設為True
                            iStart = iPureNumStart
                            iEnd = i - 1
                            iFound1 = True
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "A") Then
                        '遇到 英文字元，如果...
                        If (iStart = 0) Then
                            'iStart 為 0 ，如果...
                            If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                'a.前一個字元是 :字元 ，將 iColonAhead 設為True，將 iStart 設為 目前字元位置，將 iEnd 設為 目前字元位置
                                iColonAhead = True
                                iStart = i
                                iEnd = i
                            ElseIf (iCharTypeDict.GetValue(i - 1) <> ":") Then
                                'b.前一個字元不是 :字元， iStart 設為 目前字元位置
                                iStart = i
                            End If
                        ElseIf (iStart <> 0 And iColonAhead = True) Then
                            'iStart 不為 0，且iColonAhead 為True ， iEnd 設為 目前字元位置
                            iEnd = i
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "1") Then
                        '遇到 數字字元，如果...
                        If (iStart = 0) Then
                            'a.iStart 為 0
                            If (iPureNumStart = 0) Then
                                'a-1.iPureNumStart為 0...
                                If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                    'a-1-1.前一個字元是 :字元 > 將 iStart 設為 目前字元位置
                                    iStart = i
                                    iEnd = i
                                ElseIf (iCharTypeDict.GetValue(i - 1) = "OTHER") Then
                                    'a-1-2.前一個字元是 其他符號字元[包含空格] 或 =字元 > 將 iPureNumStart 設為 目前字元位置
                                    iPureNumStart = i
                                End If
                            ElseIf (iPureNumStart <> 0) Then
                                'a-2.iPureNumStart不為 0 且 下一個字元是 其他符號字元[包含空格] > 將iEnd設為目前字元、將 iStart 設為iPureNumStart的值
                                If (iCharTypeDict.GetValue(i + 1) = "OTHER") Then
                                    iEnd = i
                                    iStart = iPureNumStart
                                End If
                            End If
                        ElseIf (iStart <> 0) Then
                            'b.iStart 不為 0 > 將iEnd設為目前字元
                            iEnd = i
                        End If
                    End If
                End If
                
                '不論目前的字元是什麼，只要目前字元是formula最後一個字元，則如果...
                If (i = iCharTypeDict.Count) Then
                    If (iCharTypeDict.GetValue(i) = "A" And iStart <> 0 And iEnd = 0) Then
                        '目前字元是 英文字元 ，且 iStart 有值，iEnd沒有值 > 將iEnd設為目前字元  ---> 將iFound 設為True
                        iEnd = i
                        iFound1 = True
                    ElseIf (iCharTypeDict.GetValue(i) = "1" And iStart <> 0 And iEnd <> 0) Then
                        '目前字元是 數字字元 ，且 iStart & iEnd 都有值 > 將iEnd設為目前字元  ---> 將iFound 設為True
                        iEnd = i
                        iFound1 = True
                    End If
                End If
            End If
            
            'iFound為True時，紀錄此Cell / 整欄 / 整列 的值；執行紀錄後，然後將iStart & iEnd & iPureNumStart 設為0、iFound & iColonAhead 設為 false
            '   cellArray(6,n)
            '       第一維 [1,n]此筆資料從哪一個字開始 [2,n]此筆資料從哪一個字結束 [3,n]此筆資料的值 [4,n]是否屬於同一個檔案 [5,n]是否屬於同一個工作表 [6,n]是:的頭尾或都不是(輸入 H/T/NA)
            '       第二維 [n,#]第幾筆資料
            If (iFound1 = True) Then
                iCount1 = iCount1 + 1
                ReDim Preserve cellArray(6, iCount1)
                cellArray(1, iCount1) = iStart
                cellArray(2, iCount1) = iEnd
                cellArray(3, iCount1) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '找資料是前後是否有:
                '   預設都沒有
                cellArray(6, iCount1) = "NA"
                If (iCharTypeDict.GetValue(iStart - 1) = ":") Then
                    '此段cell前面有:
                    cellArray(6, iCount1) = "T"
                Else
                    If (iEnd <> iCharTypeDict.Count) Then
                        '不是最後一個字元，才找此段cell後面是否有:
                        If (iCharTypeDict.GetValue(iEnd + 1) = ":") Then
                            cellArray(6, iCount1) = "H"
                        End If
                    End If
                End If
                iStart = 0
                iEnd = 0
                iPureNumStart = 0
                iFound1 = False
                iColonAhead = False
            End If
        Next i
    End If

    
    
    '如果有任何range資料才繼續處理
    If (UBound(cellArray, 2) = 0) Then
        ReDim iArray(5, 0)
        iArray(4, 0) = True
        iArray(5, 0) = True
        GoTo 999
    Else
        '分析每個FSN是否是其他檔案、或者同檔案但是其他工作表
        If (UBound(fsnArray, 2) <> 0) Then
            For i = 1 To UBound(fsnArray, 2)
                iNotFsn = False
                '參照值中是否有檔名，有的話此fsn一定是其他檔案
                ii = InStr(fsnArray(3, i), "[")
                If (ii <> 0) Then
                    '參照值中有檔名
                    iii = ii
                    ii = InStr(iii + 1, fsnArray(3, i), "]")
                    If (ii <> 0) Then
                        fsnArray(4, i) = False
                        fsnArray(5, i) = False
                    End If
                Else
                    '參照值中不含檔名，代表只有工作表名
                        '取值-參照值不含檔名
                    If (fsnArray(3, i) <> nowSheettName) Then
                        fsnArray(4, i) = True
                        fsnArray(5, i) = False
                    Else
                        fsnArray(4, i) = True
                        fsnArray(5, i) = True
                    End If
                End If
            Next i
        End If
        '整理將各range屬於哪個fsn找出來，寫上[4,n]是否屬於同一個檔案 [5,n]是否屬於同一個工作表
        If (UBound(fsnArray, 2) = 0) Then
            'fsnArray沒值時全都判斷為要處理
            For i = 1 To UBound(cellArray, 2)
                cellArray(4, i) = True
                cellArray(5, i) = True
            Next i
        Else
            For i = 1 To UBound(cellArray, 2)
                iStr1 = Mid(cellOriginalFormula, cellArray(1, i) - 1, 1)
                If (iStr1 = ",") Then
                    '前一個字是逗號, ，則不屬於任一個fsn
                    cellArray(4, i) = True
                    cellArray(5, i) = True
                ElseIf (iStr1 = ":") Then
                    '前一個字是冒號 : ，則該range和前一個range屬於同一個fsn
                    cellArray(4, i) = cellArray(4, i - 1)
                    cellArray(5, i) = cellArray(5, i - 1)
                ElseIf (iStr1 = "!") Then
                    '前一個字是 "!"，則屬於某fsn
                    iStr2 = Mid(cellOriginalFormula, cellArray(1, i) - 2, 1)
                    If (iStr2 = "'") Then
                        ii = cellArray(1, i) - 3
                    Else
                        ii = cellArray(1, i) - 2
                    End If
                    For iii = 1 To UBound(fsnArray, 2)
                        If (fsnArray(2, iii) = ii) Then
                            cellArray(4, i) = fsnArray(4, iii)
                            cellArray(5, i) = fsnArray(5, iii)
                            Exit For
                        End If
                    Next iii
                Else
                    '都不符合以上的狀況，代表 "多個RANGE中其他的有fsn、這個沒有"，要處理
                    cellArray(4, i) = True
                    cellArray(5, i) = True
                End If
            Next i
        End If
    End If
    
    '從rangeArray蒐集所有函數使用到的欄位資訊
    '   如果參照範圍是整列(1:1)，將欄名寫上"ALL",欄號碼"1234567890"
    '   如果參照範圍是整欄(A:A)，將列號碼寫上"1234567890"
    iCount1 = 0
    ReDim iArray(5, 0)
    iArray(4, 0) = True
    iArray(5, 0) = True
    '預設值
    '   全部資料都同檔案、同工作表
    iArray(4, 0) = True
    iArray(5, 0) = True
    '   iFound1代表找到數字資料，iFound2代表找到文字資料
    iFound1 = False
    iFound2 = False
    For i = 1 To UBound(cellArray, 2)
        iStr2 = ""
        '將$符號移除
        iStr1 = WorksheetFunction.Substitute(cellArray(3, i), "$", "")
        '確認範圍值是否為整列或整欄
        '   試著尋找數字資料
        For ii = 1 To Len(iStr1)
            If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                iFound1 = True
                Exit For
            End If
        Next ii
        '   試著尋找文字資料
        For ii = 1 To Len(iStr1)
            If (IsNumeric(Mid(iStr1, ii, 1)) = False) Then
                iFound2 = True
                Exit For
            End If
        Next ii
        If (iFound1 = True And iFound2 = True) Then
            '資料值不是整列或整欄
'            For ii = 1 To Len(iStr1)
'                If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
'                    '遇到是數字，代表此值的欄名已搜索完畢
'                    Exit For
'                Else
'                    iStr2 = iStr2 & Mid(iStr1, ii, 1)
'                End If
'            Next ii
            '   將欄列值找出
            For ii = 1 To Len(iStr1)
                If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                    '遇到是數字，代表此值的欄名已搜索完畢
                    Exit For
                Else
                    iStr2 = iStr2 & Mid(iStr1, ii, 1)
                End If
            Next ii
            colName = iStr2
            colNo = convertABCto123(iStr2)
            '   將列值找出
            rowNo = Right(iStr1, Len(iStr1) - Len(colName))
        ElseIf (iFound1 = True And iFound2 = False) Then
            '資料值是整列
            colName = "ALL"
            colNo = 1234567890
            rowNo = CLng(iStr1)
        ElseIf (iFound1 = False And iFound2 = True) Then
            '資料值是整欄
            colName = iStr1
            colNo = convertABCto123(iStr1)
            rowNo = 1234567890
        End If
        
        '資料寫入陣列
        '   任一筆資料不屬於此檔案/此工作表，做紀錄
        If (cellArray(4, i) = False) Then
            iArray(4, 0) = False
        End If
        If (cellArray(5, i) = False) Then
            iArray(5, 0) = False
        End If
        '   看是否為 : 前的資料
        ibool1 = cellArray(4, i)
        ibool2 = cellArray(5, i)
        If (cellArray(6, i) = "NA") Then
            '否，此range不是一個範圍
            iCount1 = iCount1 + 1
            ReDim Preserve iArray(5, iCount1)
            iArray(1, iCount1) = colName
            iArray(2, iCount1) = colNo
            iArray(3, iCount1) = rowNo
            iArray(4, iCount1) = ibool1
            iArray(5, iCount1) = ibool2
        ElseIf (cellArray(6, i) = "H") Then
            '是，此range是範圍的頭，將範圍內的資料都寫入
            If (colNo = 1234567890 Or rowNo = 1234567890) Then
                '處理是整列或整欄的
                iStr1 = WorksheetFunction.Substitute(cellArray(3, i + 1), "$", "")
                If (rowNo = 1234567890) Then
                    '如果此範圍值是整欄，下一筆T資料一定也是整欄
                    endColName = iStr1
                    endColNo = convertABCto123(iStr1)
                    endRowNo = 1234567890
                ElseIf (colNo = 1234567890) Then
                    '如果此範圍值是整列，下一筆T資料一定也是整列
                    endColName = "ALL"
                    endColNo = 1234567890
                    endRowNo = CLng(iStr1)
                End If
                '將範圍內的寫入陣列
                '   :前
                iCount1 = iCount1 + 1
                ReDim Preserve iArray(5, iCount1)
                iArray(1, iCount1) = colName
                iArray(2, iCount1) = colNo
                iArray(3, iCount1) = rowNo
                iArray(4, iCount1) = ibool1
                iArray(5, iCount1) = ibool2
                '   :後
                iCount1 = iCount1 + 1
                ReDim Preserve iArray(5, iCount1)
                iArray(1, iCount1) = endColName
                iArray(2, iCount1) = endColNo
                iArray(3, iCount1) = endRowNo
                iArray(4, iCount1) = ibool1
                iArray(5, iCount1) = ibool2
            Else
                '處理不是整列也不是整欄的
                '   下一個一定是T，取其值
                iStr2 = ""
                '   將$符號移除
                iStr1 = WorksheetFunction.Substitute(cellArray(3, i + 1), "$", "")
                '   將欄列值找出
                For ii = 1 To Len(iStr1)
                    If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                        '遇到是數字，代表此值的欄名已搜索完畢
                        Exit For
                    Else
                        iStr2 = iStr2 & Mid(iStr1, ii, 1)
                    End If
                Next ii
                endColNo = convertABCto123(iStr2)
                '   列值找出
                endRowNo = Right(iStr1, Len(iStr1) - Len(colName))
                '將範圍內的寫入陣列
                For ii = colNo To endColNo
                    For iii = rowNo To endRowNo
                        iColName = convert123toABC(ii)
                        iColNo = ii
                        iRowNo = iii
                        iCount1 = iCount1 + 1
                        ReDim Preserve iArray(5, iCount1)
                        iArray(1, iCount1) = iColName
                        iArray(2, iCount1) = iColNo
                        iArray(3, iCount1) = iRowNo
                        iArray(4, iCount1) = ibool1
                        iArray(5, iCount1) = ibool2
                    Next iii
                Next ii
            End If
            '下一個跳過
            i = i + 1
        End If
    Next i


999
getAllRangesInfoInFormula = iArray

''======檢查用=====
'Dim mySht As Worksheet
'    Set mySht = ThisWorkbook.Worksheets("TESTX")
'    mySht.Cells.ClearContents
'    '標題
'    mySht.Cells(1, 1) = "fsnArray(1,x)"
'    mySht.Cells(1, 2) = "fsnArray(2,x)"
'    mySht.Cells(1, 3) = "fsnArray(3,x)"
'    mySht.Cells(1, 4) = "fsnArray(4,x)"
'    mySht.Cells(1, 5) = "fsnArray(5,x)"
'
'    mySht.Cells(1, 6) = "cellArray(1,x)"
'    mySht.Cells(1, 7) = "cellArray(2,x)"
'    mySht.Cells(1, 8) = "cellArray(3,x)"
'    mySht.Cells(1, 9) = "cellArray(4,x)"
'    mySht.Cells(1, 10) = "cellArray(5,x)"
'    mySht.Cells(1, 11) = "cellArray(6,x)"
'
'    mySht.Cells(1, 12) = "iArray(1,x)"
'    mySht.Cells(1, 13) = "iArray(2,x)"
'    mySht.Cells(1, 14) = "iArray(3,x)"
'    mySht.Cells(1, 15) = "iArray(4,x)"
'    mySht.Cells(1, 16) = "iArray(5,x)"
'
'    For i = 0 To UBound(fsnArray, 2)
'        For ii = 1 To 5
'            mySht.Cells(i + 2, ii).Value = fsnArray(ii, i)
'        Next ii
'    Next i
'    For i = 0 To UBound(cellArray, 2)
'        For ii = 1 To 6
'            mySht.Cells(i + 2, 5 + ii).Value = cellArray(ii, i)
'        Next ii
'    Next i
'    For i = 0 To UBound(iArray, 2)
'        For ii = 1 To 5
'            mySht.Cells(i + 2, 5 + 6 + ii).Value = iArray(ii, i)
'        Next ii
'    Next i
''=================
End Function
