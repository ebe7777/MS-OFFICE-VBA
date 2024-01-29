Attribute VB_Name = "關於英文字與數字的處理"
Public Function convertABCto123_資料庫(ByVal inputVal As String)
Dim baseStr As String, addStr As String
Dim baseNo As Long, addNo As Long
Dim findBaseNo As Boolean, findAddNo As Boolean
'目前支援A~ZZ換算成數字
'如果字串當中包無法轉換的字，輸出結果一律為0
'如果字串大於等於3，輸出結果一律為0

    '將英文欄名拆為基本值與累加值
    '   基本值 = 該欄的前一欄位置,以數字表示
    '   累加值 = 基本值加上累加值等於該欄,以數字表示
    'A~Z：基本值為0,A~Z換算成1~26作為累加值
    'Ax~Zx，左邊的字是基本值，右邊的字是累加值
    '譬如:C欄 =>0 + 3;AC欄 =>26 + 3
    
    '防呆-只能處理A~ZZ
    If (Len(inputVal) > 2) Then
        convertABCto123 = 0
    Else
        '防呆設置-基本值或累加值任一個不是英文
        findBaseNo = False
        findAddNo = False
        '找出基本值
        If (Len(inputVal) = 1) Then
            baseNo = 0
            findBaseNo = True
            addStr = UCase(inputVal)
        ElseIf (Len(inputVal) = 2) Then
            baseNo = 0
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
            If (baseNo <> 0) Then
                findBaseNo = True
            End If
        End If
        '找出累加值
        addNo = 0
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
        End Select
        If (addNo <> 0) Then
            findAddNo = True
        End If
        
        If (findBaseNo = True And findAddNo = True) Then
            convertABCto123 = baseNo + addNo
        Else
            convertABCto123 = 0
        End If
    End If
    
End Function
Public Function convert123toABC_資料庫(ByVal inputVal As Long)
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
Public Function isABC_資料庫(ByVal inputVal As String)
    '檢查輸入值是否為英文
    isABC = False
    inputVal = UCase(inputVal)
    If (inputVal = "A" Or inputVal = "B" Or inputVal = "C" Or inputVal = "D" Or inputVal = "E" Or inputVal = "F" Or inputVal = "G" Or inputVal = "H" Or inputVal = "I" Or inputVal = "J" Or inputVal = "K" Or inputVal = "L" Or inputVal = "M" Or inputVal = "N" Or inputVal = "O" Or inputVal = "P" Or inputVal = "Q" Or inputVal = "R" Or inputVal = "S" Or inputVal = "T" Or inputVal = "U" Or inputVal = "V" Or inputVal = "W" Or inputVal = "X" Or inputVal = "Y" Or inputVal = "Z") Then
        isABC = True
    End If
End Function
Public Function is123_資料庫(ByVal inputVal As String)
    '檢查輸入值是否為數字
    is123 = IsNumeric(inputVal)
End Function
Private Function chkValIsNumetricAndValid_資料庫(ByVal chkThisValue As Variant) As Boolean
'檢查輸入值，以下狀況回報false
'   不是數值、不是整數、小於1
    chkValIsNumetricAndValid = False
    If (IsNumeric(chkThisValue) = False) Then
        GoTo 999
    End If
    If (chkThisValue <> CLng(chkThisValue)) Then
        GoTo 999
    End If
    If (chkThisValue < 1) Then
        GoTo 999
    End If
    chkValIsNumetricAndValid = True
999
End Function

Public Function isAbcBetweenAToAz_資料庫(ByVal inputVal As String)
    inputVal = UCase(Trim(inputVal))
    '檢查是否為空白
    If (inputVal = "") Then
        isAbcBetweenAToAz = False
    Else
        '檢查輸入值是否為英文，取介於A~AZ
        If (inputVal = "A" Or inputVal = "B" Or inputVal = "C" Or inputVal = "D" Or inputVal = "E" Or inputVal = "F" Or inputVal = "G" Or inputVal = "H" Or inputVal = "I" Or inputVal = "J" Or inputVal = "K" Or inputVal = "L" Or inputVal = "M" Or inputVal = "N" Or inputVal = "O" Or inputVal = "P" Or inputVal = "Q" Or inputVal = "R" Or inputVal = "S" Or inputVal = "T" Or inputVal = "U" Or inputVal = "V" Or inputVal = "W" Or inputVal = "X" Or inputVal = "Y" Or inputVal = "Z") Then
            isAbcBetweenAToAz = True
        ElseIf (inputVal = "AA" Or inputVal = "AB" Or inputVal = "AC" Or inputVal = "AD" Or inputVal = "AE" Or inputVal = "AF" Or inputVal = "AG" Or inputVal = "AH" Or inputVal = "AI" Or inputVal = "AJ" Or inputVal = "AK" Or inputVal = "AL" Or inputVal = "AM" Or inputVal = "AN" Or inputVal = "AO" Or inputVal = "AP" Or inputVal = "AQ" Or inputVal = "AR" Or inputVal = "AS" Or inputVal = "AT" Or inputVal = "AU" Or inputVal = "AV" Or inputVal = "AW" Or inputVal = "AX" Or inputVal = "AY" Or inputVal = "AZ") Then
            isAbcBetweenAToAz = True
        End If
    End If
End Function
Public Function is123NotEmptyNotZero_資料庫(ByVal inputVal As String)
    '預設為true
    is123NotEmptyNotZero = True
    '檢查是否為空白
    If (Trim(inputVal) = "") Then
        is123NotEmptyNotZero = False
    Else
        '檢查輸入值是否為數字
        If (IsNumeric(inputVal) = False) Then
            is123NotEmptyNotZero = False
        Else
            '檢查是否為0
            If (CLng(inputVal) = 0) Then
                is123NotEmptyNotZero = False
            '檢查是否為整數
            ElseIf (CDbl(inputVal) <> CInt(inputVal)) Then
                is123NotEmptyNotZero = False
            End If
        End If
    End If
End Function
Public Function is123NotEmpty_資料庫(ByVal inputVal As String)
    '預設為true
    is123NotEmpty = True
    '檢查是否為空白
    If (Trim(inputVal) = "") Then
        is123NotEmpty = False
    Else
        '檢查輸入值是否為數字
        If (IsNumeric(inputVal) = False) Then
            is123NotEmpty = False
        '檢查是否為整數
        ElseIf (CDbl(inputVal) <> CInt(inputVal)) Then
            is123NotEmpty = False
        End If
    End If
End Function
