Attribute VB_Name = "將儲存格的函數轉為文字顯示"
Sub findref1_資料庫()
'將每張工作表內的每個儲存格的公式開頭為 "=" 者，並將顯示改為"fm is 原始函數"
Dim iSht As Worksheet
Dim i As Long, ii As Long, iRow As Long, iCol As Long
Dim iStr1 As String, stringiColStr As String
    For Each iSht In ThisWorkbook.Worksheets
        iCol = 0
        For i = 1 To 9
            ii = myDataColumns(ThisWorkbook.Name, iSht.Name, i, "ZZ")
            If (ii > iCol) Then
                iCol = ii
            End If
        Next i
        iColStr = convert123toABC(iCol)
        iRow = findMaxRowNo(ThisWorkbook.Name, iSht.Name, "A", iColStr)
        
        For i = 1 To iRow
            For ii = 1 To iCol
                iStr1 = iSht.Cells(i, ii).Formula
                If (InStr(1, iStr1, "=") = 1) Then
                    iSht.Cells(i, ii).Value = "fm is " & iStr1
                    With iSht.Cells(i, ii).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With iSht.Cells(i, ii).Font
                        .Color = -1003520
                        .TintAndShade = 0
                    End With
                End If
            Next ii
        Next i
    Next
    
    MsgBox "Done"
End Sub
Sub findref2_資料庫()
'將目前工作表內的每個儲存格的公式開頭為 "=" 者，並將顯示改為"fm is 原始函數"
Dim iSht As Worksheet
Dim i As Long, ii As Long, iRow As Long, iCol As Long
Dim iStr1 As String, stringiColStr As String

    iCol = 0
    For i = 1 To 9
        ii = myDataColumns(ThisWorkbook.Name, ActiveSheet.Name, i, "ZZ")
        If (ii > iCol) Then
            iCol = ii
        End If
    Next i
    iColStr = convert123toABC(iCol)
    iRow = findMaxRowNo(ThisWorkbook.Name, ActiveSheet.Name, "A", iColStr)
    
    For i = 1 To iRow
        For ii = 1 To iCol
            iStr1 = ActiveSheet.Cells(i, ii).Formula
            If (InStr(1, iStr1, "=") = 1) Then
                ActiveSheet.Cells(i, ii).Value = "fm is " & iStr1
                With ActiveSheet.Cells(i, ii).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.399975585192419
                    .PatternTintAndShade = 0
                End With
                With ActiveSheet.Cells(i, ii).Font
                    .Color = -1003520
                    .TintAndShade = 0
                End With
            End If
        Next ii
    Next i

    MsgBox "Done"
End Sub

Public Function myDataColumns(ByVal bookName As String, ByVal sheetName As String, ByVal rowNumber As Long, ByVal countLeftwardFromThisColumn As String)
'以下舊方式如果遇到最後幾欄被隱藏，則該部分不會納入計算範圍
'   myDataColumns = Workbooks(bookName).Sheets(sheetName).Range(countLeftwardFromThisColumn & rowNumber).End(xlToLeft).Column
'需搭配 convertABCto123 使用

Dim myArray As Variant
Dim i As Long, maxCol As Long
    '為加速讀取速度，先將儲存格值寫入陣列
    maxCol = convertABCto123(countLeftwardFromThisColumn)
    With Workbooks(bookName).Sheets(sheetName)
        myArray = .Range(.Cells(rowNumber, 1), .Cells(rowNumber, maxCol)).Formula
    End With
    For i = UBound(myArray, 2) To 1 Step -1
        If (myArray(1, i) <> "") Then
            myDataColumns = i
            Exit For
        End If
    Next i
End Function
Public Function myDataRows(ByVal bookName As String, ByVal sheetName As String, ByVal columnName As String, ByVal countUpwardFromThisRow As Long)
'以下舊方式如果遇到最後幾列被篩選掉/隱藏，則該部分不會納入計算範圍
'    myDataRows = Workbooks(bookName).Sheets(sheetName).Range(columnName & countUpwardFromThisRow).End(xlUp).Row

Dim myArray As Variant
Dim i As Long
    '為加速讀取速度，先將儲存格值寫入陣列
    With Workbooks(bookName).Sheets(sheetName)
        myArray = .Range(columnName & "1:" & columnName & countUpwardFromThisRow).Formula
    End With
    For i = UBound(myArray, 1) To 1 Step -1
        If (myArray(i, 1) <> "") Then
            myDataRows = i
            Exit For
        End If
    Next i
' 如果將function放置在一個sub中，另一個sub要呼叫此sub的function，使用 call Module名稱1.Sub名稱
End Function

Public Function findMaxRowNo(bookName As String, sheetName As String, startColName As String, ByVal endColName As String)
'找到某檔某工作表指定的欄範圍內使用的最多列的號碼
Dim iStart As Long, iEnd As Long, iRows As Long, iMaxRows As Long
Dim i As Long
    iStart = convertABCto123(startColName)
    iEnd = convertABCto123(endColName)
    iMaxRows = 0
    For i = iStart To iEnd
        iRows = myDataRows(bookName, sheetName, convert123toABC(i), 65536)
        If (iRows > iMaxRows) Then
            iMaxRows = iRows
        End If
    Next i
    findMaxRowNo = iMaxRows
End Function
Public Function convert123toABC(inputVal As Long)
Dim quotientNo As Long, remainderNo As Long
Dim leftStr As String, rightStr As String
'目前只支援換成成A~ZZ
    quotientNo = WorksheetFunction.RoundDown(inputVal / 26, 0)
    '最多到ZZ
    If (quotientNo <= 27) Then
        remainderNo = inputVal Mod 26
        If (remainderNo = 0) Then
            quotientNo = quotientNo - 1
        End If
        Select Case quotientNo
            Case 0
            leftStr = ""
            Case 1
                leftStr = "A"
            Case 2
                leftStr = "B"
            Case 3
                leftStr = "C"
            Case 4
                leftStr = "D"
            Case 5
                leftStr = "E"
            Case 6
                leftStr = "F"
            Case 7
                leftStr = "G"
            Case 8
                leftStr = "H"
            Case 9
                leftStr = "I"
            Case 10
                leftStr = "J"
            Case 11
                leftStr = "K"
            Case 12
                leftStr = "L"
            Case 13
                leftStr = "M"
            Case 14
                leftStr = "N"
            Case 15
                leftStr = "O"
            Case 16
                leftStr = "P"
            Case 17
                leftStr = "Q"
            Case 18
                leftStr = "R"
            Case 19
                leftStr = "S"
            Case 20
                leftStr = "T"
            Case 21
                leftStr = "U"
            Case 22
                leftStr = "V"
            Case 23
                leftStr = "W"
            Case 24
                leftStr = "X"
            Case 25
                leftStr = "Y"
            Case 26
                leftStr = "Z"
        End Select
        Select Case remainderNo
            Case 1
                rightStr = "A"
            Case 2
                rightStr = "B"
            Case 3
                rightStr = "C"
            Case 4
                rightStr = "D"
            Case 5
                rightStr = "E"
            Case 6
                rightStr = "F"
            Case 7
                rightStr = "G"
            Case 8
                rightStr = "H"
            Case 9
                rightStr = "I"
            Case 10
                rightStr = "J"
            Case 11
                rightStr = "K"
            Case 12
                rightStr = "L"
            Case 13
                rightStr = "M"
            Case 14
                rightStr = "N"
            Case 15
                rightStr = "O"
            Case 16
                rightStr = "P"
            Case 17
                rightStr = "Q"
            Case 18
                rightStr = "R"
            Case 19
                rightStr = "S"
            Case 20
                rightStr = "T"
            Case 21
                rightStr = "U"
            Case 22
                rightStr = "V"
            Case 23
                rightStr = "W"
            Case 24
                rightStr = "X"
            Case 25
                rightStr = "Y"
            Case 0
                rightStr = "Z"
        End Select
        convert123toABC = leftStr & rightStr
    End If
End Function
Public Function convertABCto123(inputVal As String)
Dim baseStr As String, addStr As String
Dim baseNo As Long, addNo As Long
'目前支援A~ZZ換算成數字

    '將英文欄名拆為基本值與累加值
    '   基本值 = 該欄的前一欄位置,以數字表示
    '   累加值 = 基本值加上累加值等於該欄,以數字表示
    'A~Z：基本值為0,A~Z換算成1~26作為累加值
    'Ax~Zx，左邊的字是基本值，右邊的字是累加值
    '譬如:C欄 =>0 + 3;AC欄 =>26 + 3
    
    '找出基本值
    If (Len(inputVal) = 1) Then
        baseNo = 0
        addStr = inputVal
    ElseIf (Len(inputVal) = 2) Then
        baseStr = UCase(Left(inputVal, 1))
        addStr = UCase(Right(inputVal, 1))
        Select Case baseStr
        Case "A"
            baseNo = 26
        Case "B"
            baseNo = 52
        Case "C"
            baseNo = 78
        Case "D"
            baseNo = 104
        Case "E"
            baseNo = 130
        Case "F"
            baseNo = 156
        Case "G"
            baseNo = 182
        Case "H"
            baseNo = 208
        Case "I"
            baseNo = 234
        Case "J"
            baseNo = 260
        Case "K"
            baseNo = 286
        Case "L"
            baseNo = 312
        Case "M"
            baseNo = 338
        Case "N"
            baseNo = 364
        Case "O"
            baseNo = 390
        Case "P"
            baseNo = 416
        Case "Q"
            baseNo = 442
        Case "R"
            baseNo = 468
        Case "S"
            baseNo = 494
        Case "T"
            baseNo = 520
        Case "U"
            baseNo = 546
        Case "V"
            baseNo = 572
        Case "W"
            baseNo = 598
        Case "X"
            baseNo = 624
        Case "Y"
            baseNo = 650
        Case "Z"
            baseNo = 676
        End Select
    End If
    '找出累加值
    Select Case addStr
        Case "A"
            addNo = 1
        Case "B"
            addNo = 2
        Case "C"
            addNo = 3
        Case "D"
            addNo = 4
        Case "E"
            addNo = 5
        Case "F"
            addNo = 6
        Case "G"
            addNo = 7
        Case "H"
            addNo = 8
        Case "I"
            addNo = 9
        Case "J"
            addNo = 10
        Case "K"
            addNo = 11
        Case "L"
            addNo = 12
        Case "M"
            addNo = 13
        Case "N"
            addNo = 14
        Case "O"
            addNo = 15
        Case "P"
            addNo = 16
        Case "Q"
            addNo = 17
        Case "R"
            addNo = 18
        Case "S"
            addNo = 19
        Case "T"
            addNo = 20
        Case "U"
            addNo = 21
        Case "V"
            addNo = 22
        Case "W"
            addNo = 23
        Case "X"
            addNo = 24
        Case "Y"
            addNo = 25
        Case "Z"
            addNo = 26
        Case "AA"
    End Select
    
    convertABCto123 = baseNo + addNo
End Function
