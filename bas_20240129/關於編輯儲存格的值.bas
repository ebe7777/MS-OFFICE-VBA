Attribute VB_Name = "����s���x�s�檺��"
Private Sub Worksheet_Change(ByVal Target As Range)
'���o���ק諸�s�x�檺range�A�ûP�ۭq���@�ӽd�����o�� "�ק諸�O�_�b���S�w�d��"
'��sub�u��g�b�u�@��
'https://docs.microsoft.com/zh-tw/office/troubleshoot/excel/run-macro-cells-change
'https://docs.microsoft.com/zh-tw/office/vba/api/excel.application.intersect
Dim myRange1 As Range
Dim myRange2 As Range
    Set myRange1 = Range("C10:C11")
    Set myRange2 = Range(Target.address)
        
    If (Application.Intersect(myRange1, myRange2) Is Nothing) Then
        MsgBox "�ק諸�x�s�� ���b ���w�d��"
    Else
        MsgBox "�ק諸�x�s �b ���w�d��"
    End If
End Sub

Sub �ƻs�K�W_�зǻy�k_��Ʈw()
    Set mySht = ActiveSheet
    '�ϥ�range�f�tcell���w�ƻs�d��
    With mySht
        .Range(.Cells(1, 1), .Cells(1, 1)).copy
    End With
    
    '�K�W(�קK�K�W�ɦ����󪬪p-�p �W�٭��� �ɭP�{���Ȱ�)
    Application.DisplayAlerts = False
    mySht.Paste Destination:=mySht.Cells(2, 1)
    Application.DisplayAlerts = True
    
    '�@���g���ƻs�M�K�W
    Range("A1").copy Destination:=Range("A2:A3")
End Sub

Sub �ƻs��i�u�@��K�W����_��Ʈw()


    Cells.Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub �N�Ů�ƻs�W���()
PIVOT_ALL_ROWS = Sheets("�ϯä��R").Range("M1").End(xlDown).Row

Range("A1:M" & PIVOT_ALL_ROWS).SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
Sheets("�ϯä��R").Cells.Select
    Selection.copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Cells.Select
    Selection.Replace What:="(�ť�)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Public Function trimAndReplace2SpaceTo1_��Ʈw(myString As String)
'�N�Ҧ��۳s���h�ӪŮ������1�ӪŮ�
Dim iStr1 As String
    iStr1 = Trim(myString)
    iStr1 = Replace(iStr1, "  ", " ")
    If (InStr(1, iStr1, "  ") <> 0) Then
        Call trimAndReplace2SpaceTo1(iStr1)
    Else
        trimAndReplace2SpaceTo1 = iStr1
    End If
End Function

Public Function ��Ʈw_clearRange(myRange As Range)
'�M���x�s�� ���e�B����B�r��
    With myRange
        .Formula = ""
        '.ClearComments
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
End Function
Public Function ��Ʈw_varyRangeInSheetFunction1Main(cellOriginalFormula As Variant, copyToThisWB As Workbook, copyToThisWS As Worksheet, varyOnlySameWB As Boolean, varyOnlySameWS As Boolean, varyNoneFixedCol As Boolean, varyFixedCol As Boolean, varyNoneFixedRow As Boolean, varyFixedRow As Boolean, cellOriginalAdress As String, cellNewAdress As String, variedFormula As String, isVariedFormulaOutOfCell As Boolean, Optional allCellColVaryInfoArray As Variant, Optional allCellRowVaryInfoArray As Variant)
'Public Function varyRangeInSheetFunction1Main(cellOriginalFormula As String, copyToThisWB As Workbook, copyToThisWS As Worksheet, varyOnlySameWB As Boolean, varyOnlySameWS As Boolean, varyNoneFixedCol As Boolean, varyFixedCol As Boolean, varyNoneFixedRow As Boolean, varyFixedRow As Boolean, cellOriginalAdress As String, cellNewAdress As String, variedFormula As String, isVariedFormulaOutOfCell As Boolean, Optional allCellColVaryInfoArray As Variant, Optional allCellRowVaryInfoArray As Variant)
'�{���\��
'   �ҥ�excel���\�� , �ƻs�x�s�椺�e�ɦp���e�O���, ��Ƥ����x�s��range�|���۹諸�ܰ�
'algorithm
'   ���R��ƪ��զ��ñN��Ʒ���range�D�X > �H���ʫe�M���ʫ᪺�x�s��range�t�O���p���Ƥ���range�Ӧp����� > �N��Ƥ���range��s
'�ݭn�H�Ufunction
'   convertABCto123 ,convert123toABC ,varyRangeInSheetFunction2Combine
'�ǤJ�ܼƻ���
'   cellOriginalFormula�G
'       �n�ˬd���x�s��Formula(�p�D��Ʒ|�۰ʸ��L�ˬd)�A��func�N�ˬd����Ƥ����C��range�O�_�n�]�� "�����e��Ө�ƩҦb�x�s����C���P" �ӻݭn������
'   copyToThisWB / copyToThisWS�G
'       ��ƭn�K�b���Ӭ���ï / �u�@��
'   varyOnlySameWB / varyOnlySameWS�G
'       �O�_ "�ƻs�M�K�W������ï/�u�@��" �W�٬ۦP�~�B�z
'   varyNoneFixedCol / varyFixedCol As Boolean / varyNoneFixedRow / varyFixedRow As Boolean�G
'       �O�_�B�z "�T�w��(��$�Ÿ�)/���T�w��(�S$�Ÿ�) �� ��/�C"
'   cellOriginalAdress / cellNewAdress�G
'       �n�ˬd����ƩҦb���x�s��b�ƻs�� / �K�W�� ���s�x���}�A��ƫ��A�� "A1" �榡
'   variedFormula�G
'       ���ܫ᪺Formula�ȡA�I�s��func�̨��Φ��Ȭ��̲׵��G
'   isVariedFormulaOutOfCell�G
'       �аO�B�z�L�{�O�_�o�� "���ܫ᪺Col�p��A�BRow�p��1" �����p�F�o�ͦ����p��variedFormula�|����cellOriginalFormula�G�I�s��func�̻ݦۦ漶�g���~�T��
'   [��ܩ�]allCellColVaryInfoArray() / allCellRowVaryInfoArray()�G
'       �����e���i�u�@����C���ܪ���T�A�ΥH���ܨ�Ƥ���Range��
'           �p�G�Ӱ}�C�s�b�h��Ƥ�Range���ק�H�}�C������T�B�z�F�p�G�}�C���䤣���Range����T�h�H�x�s�檺���ʳB�zRange
'           �p�G�Ӱ}�C���s�b�A�h�H�x�s�檺���ʳB�zRange
'       [1,n]�ƻs��Ƥu�@����/�C���X(�H�Ʀr���) [2,n]�P�˪���/�C�b�K�W��Ƥu�@����/�C���X(�H�Ʀr���) [3,n]�P�w��/�C�O�_���ܪ���ƪ���(Ĵ�psn)
'           �M�w��C����ƭȥ����O����C���ߤ@�ȡF�p���O�h�@�ߥH"����"�B�z
'       [n,#]�ĴX�����
'           ��Ʀ��X��/�C�N���X����
'�Ƶ�
'   ���{������ "�ѷӪ��u�@���s�b" ���ˬd�A�]���n�Ҽ{�o�Ӧh

Dim funcSplitArray(), fsnArray(), rangeArray()
Dim cellAddressRowMoveValue As Long, cellAddressColMoveValue As Long
Dim i As Long, ii As Long, iii As Long, iv As Long, iStart As Long, iEnd As Long
Dim iQuotationMarkStart As Long, iQuotationMarkEnd As Long
Dim iWordInMid As Long, iNumInMid As Long
Dim iStr1 As String, iStr2 As String, iStr3 As String, iOriginalCN As String
Dim isNum As Boolean, isWord As Boolean, isSymble As Boolean
Dim iCount1 As Long, iCount2 As Long
Dim iOK As Boolean, iFound As Boolean
    '�P�_(1)�n�p�⪺formula���Ȭ���� (2)��D�T�w��/�D�T�w�C/�T�w��/�T�w�C �|�̨䤤�@�̬�true �~�~��B�z
    If ((Left(cellOriginalFormula, 1) <> "=") Or (varyNoneFixedCol = False And varyNoneFixedRow = False And varyFixedCol = False And varyFixedRow = False)) Then
        variedFormula = cellOriginalFormula
        Exit Function
    Else
        '�p��X row���ʭȡBcol���ʭ�
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
        '�Nformula�ѪR�é�J�}�C

        '   ��XFSN (File & Sheet Name)
        '       FSN�@�w�}�l��B��Ÿ��A�õ����� [!]
        '           �B��Ÿ�: = ( , + - * /
        iStart = 0
        iEnd = 0
        iCount1 = 0
        iCount2 = 0
        iQuotationMarkStart = 0
        iQuotationMarkEnd = 0
        ReDim fsnArray(4, 0)
        For i = 1 To Len(cellOriginalFormula)
            iStr1 = Mid(cellOriginalFormula, i, 1)
            'FSN���ȥ����]�i��]�t�B��Ÿ�,���J��o�Ǫ��p�|�Q��޸�(')���k�]�_
            If (iStr1 = "'") Then
                iCount1 = iCount1 + 1
                If (iCount1 Mod 2 = 1) Then
                    iQuotationMarkStart = i
                    'iQuotationMarkEnd�b�S���e�O���MStart�@��
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
            '   �Ĥ@�� [1,n]������Ʊq���@�Ӧr�}�l [2,n]������Ʊq���@�Ӧr���� [3,n]������ƪ��� [4,n]���FSN�W�h�ᦹ�ȬO�_�ק�
            '   �ĤG�� [n,#]�ĴX�����
            If (iStart <> 0 And iEnd <> 0) Then
                iCount2 = iCount2 + 1
                ReDim Preserve fsnArray(4, iCount2)
                fsnArray(1, iCount2) = iStart
                fsnArray(2, iCount2) = iEnd
                fsnArray(3, iCount2) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '�p�G���O�� ' �}�Y�����p�A�קK���h�l�Ů�ҥHtrim
                If (iQuotationMarkStart = 0) Then
                    fsnArray(3, iCount2) = Trim(fsnArray(3, iCount2))
                End If
                
                iStart = 0
                iEnd = 0
                iQuotationMarkStart = 0
                iQuotationMarkEnd = 0
            End If
        Next i

        '   ��XRANGE
        '       range���}�Y�@�w�O�Ÿ��B�����@�w�O�Ÿ���formula�̥��Frange���զ��@�w�� �^��(+�^��...)+�Ʀr(+�Ʀr....)
        iCount1 = 0
        iWordInMid = 0
        iNumInMid = 0
        iStart = 0
        iEnd = 0
        ReDim rangeArray(5, 0)
        For i = 1 To Len(cellOriginalFormula)
            '���bfsn�d�򤺪��~����
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
                'iEnd�S�Q���e���H���k�s
                iEnd = 0
    
                '�P�_�C�@�Ӧr�O�_���^��μƦr�A�p�����O�h���Ÿ�
                iStr2 = UCase(iStr1)
                If (iStr2 = "A" Or iStr2 = "B" Or iStr2 = "C" Or iStr2 = "D" Or iStr2 = "E" Or iStr2 = "F" Or iStr2 = "G" Or iStr2 = "H" Or iStr2 = "I" Or iStr2 = "J" Or iStr2 = "K" Or iStr2 = "L" Or iStr2 = "M" Or iStr2 = "N" Or iStr2 = "O" Or iStr2 = "P" Or iStr2 = "Q" Or iStr2 = "R" Or iStr2 = "S" Or iStr2 = "T" Or iStr2 = "U" Or iStr2 = "V" Or iStr2 = "W" Or iStr2 = "X" Or iStr2 = "Y" Or iStr2 = "Z") Then
                    isWord = True
                Else
                    isWord = False
                End If
                '   $�Ÿ������^���r
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
                    '�̫�@�Ӧr�O�Ÿ�
                    If (iNumInMid = 0) Then
                        iStart = i + 1
                        iWordInMid = 0
                    ElseIf (iWordInMid <> 0 And iNumInMid <> 0) Then
                        iEnd = i - 1
                        iWordInMid = 0
                        iNumInMid = 0
                    End If
                ElseIf (i = Len(cellOriginalFormula) And isNum = True And iWordInMid <> 0 And iNumInMid <> 0) Then
                    '�̫�@�Ӧr���O�Ÿ��B�O�Ʀr
                    iEnd = i
                    iWordInMid = 0
                    iNumInMid = 0
                End If
    
                
                '���range����Ƽg�i�}�C
                'rangeArray(3,n)
                '   �Ĥ@�� [1,n]������Ʊq���@�Ӧr�}�l [2,n]������Ʊq���@�Ӧr���� [3,n]������ƪ���(�­ȡF�ק���л\����) [4,n]���FSN�W�h�ᦹ�ȬO�_�ק� [5,n]��range�ݩ���@��FSN
                '   �ĤG�� [n,#]�ĴX�����
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
    '�p�G������range��Ƥ~�~��B�z
    If (UBound(rangeArray, 2) = 0) Then
        variedFormula = cellOriginalFormula
    Else
        '��FSN�W�h�P�wfsnArray���U�լO�_���ݭק�
        '   ��X�ѷӭȬO�_�]�t�ɦW�B�ɦW�O�_�PcopyTo�ۦP�F�u�@��W�٬O�_�PcopyTo�ۦP
        For i = 1 To UBound(fsnArray, 2)
            iOK = True
            '�ѷӭȤ������|�P�ɦW
            If (varyOnlySameWB = True) Then
                If (InStr(fsnArray(3, i), "\") <> 0) Then
                    '���|
                    iStr1 = Left(fsnArray(3, i), InStr(fsnArray(3, i), "\[") - 1)
                    If (iStr1 <> copyToThisWB.Path) Then
                        iOK = False
                    End If
                    '�ɦW
                    ii = InStr(fsnArray(3, i), "[") + 1
                    iii = InStr(fsnArray(3, i), "]") - 1
                    iStr2 = Mid(fsnArray(3, i), ii, iii - ii + 1)
                    If (iStr2 <> copyToThisWB.Name) Then
                        iOK = False
                    End If
                End If
            End If
            '�ѷӭȤ����u�@��W
            If (varyOnlySameWS = True) Then
                '�ѷӭȤ����u�@��W
                ii = InStr(fsnArray(3, i), "]")
                If (ii <> 0) Then
                    '����-�ѷӭȥ]�t���ɦW�b�`
                    iStr1 = Right(fsnArray(3, i), Len(fsnArray(3, i)) - ii)
                Else
                    '����-�ѷӭȤ��t�ɦW
                    iStr1 = fsnArray(3, i)
                End If
                If (iStr1 <> copyToThisWS.Name) Then
                    iOK = False
                End If
            End If
            
            fsnArray(4, i) = iOK
        Next i
        '�P�_�Urange��ƬO�_�n�ק�
        If (UBound(fsnArray, 2) = 0) Then
            'fsnArray�S�Ȯɥ����P�_���n�B�z
            For i = 1 To UBound(rangeArray, 2)
                rangeArray(4, i) = True
                rangeArray(5, i) = "NA"
            Next i
        Else
            For i = 1 To UBound(rangeArray, 2)
                '�N�Urange�O�ݩ����fsn����X��A�Nrange�]�w���Pfsn�ۦP����(�b�����Ҽ{varyXXXRow/varyXXXCol���v�T)
                
                iStr1 = Mid(cellOriginalFormula, rangeArray(1, i) - 1, 1)
                If (iStr1 = ",") Then
                    '�e�@�Ӧr�O�r��, �A�h���ݩ���@��fsn
                    rangeArray(4, i) = True
                    rangeArray(5, i) = "NA"
                ElseIf (iStr1 = ":") Then
                    '�e�@�Ӧr�O�_�� : �A�h��range�M�e�@��range�ݩ�P�@��fsn
                    rangeArray(4, i) = rangeArray(4, i - 1)
                    rangeArray(5, i) = rangeArray(5, i - 1)
                ElseIf (iStr1 = "!") Then
                    '�e�@�Ӧr�O "!"�A�h�h�ݩ��fsn
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
                    '�����ŦX�H�W�����p�A�N�� "�h��RANGE����L����fsn�B�o�ӨS��"�A�n�B�z
                    rangeArray(4, i) = True
                    rangeArray(5, i) = "NA"
                End If
            Next i
        End If
        
    
    
''======test��
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
    
        '�N�Orange��row/column�̷��x�s��Range�ܰʰ��p��
        isVariedFormulaOutOfCell = False
        For i = 1 To UBound(rangeArray, 2)
            '��range�ݭn�׭n�ק�~�i�@�B�P�_range���e�O�_�n�ץ�
            If (rangeArray(4, i) = True) Then
                iStr1 = ""
                '��ȧ�X
                If (Left(rangeArray(3, i), 1) = "$") Then
                    '�p�G��Ȧ��T�w(�Ĥ@�Ӧr��$)�A�q�ĤG�Ӧr�}�l�P�_�O�_�����
                    iii = 2
                Else
                    '�p�G��ȨS���T�w�A�q�Ĥ@�Ӧr�}�l�P�_
                    iii = 1
                End If
                For ii = iii To Len(rangeArray(3, i))
                    If (IsNumeric(Mid(rangeArray(3, i), ii, 1)) = True Or Mid(rangeArray(3, i), ii, 1) = "$") Then
                        Exit For
                    Else
                        iStr1 = iStr1 & Mid(rangeArray(3, i), ii, 1)
                    End If
                Next ii
                '�O���U��l��W�A�ѫ���ϥ�
                iOriginalCN = iStr1
                '   �p�G�ŦX�ݭn���s�p�⪺����A�h�p��
                '       varyNoneFixedCol/varyFixedCol ����O�_�ŦX
                If ((varyNoneFixedCol = True And iii = 1) Or (varyFixedCol = True And iii = 2)) Then
                    '�p�G��ܩʰ}�C����J�h�H�}�C���e�p��F�}�C���d�L��� �� �S��J�}�C�h�HcellAddressColMoveValue�p��
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
                        '�p�G���ܫ᪺��Ȥp��A�A�פ�Ҧ��ܰʭp��
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
                '   �p�G�ݭn�A�ɦ^"$"
                If (iii = 2) Then
                    iStr2 = "$" & iStr2
                    iOriginalCN = "$" & iOriginalCN
                End If
        
        
                '�C�ȧ�X
                iStr1 = Right(rangeArray(3, i), Len(rangeArray(3, i)) - Len(iOriginalCN))
                If (Left(iStr1, 1) = "$") Then
                    '�p�G�C�Ȧ��T�w(�Ĥ@�Ӧr��$)
                    ii = 1
                Else
                    '�p�G�C�ȨS���T�w
                    ii = 0
                End If
                iii = CLng(Right(rangeArray(3, i), Len(rangeArray(3, i)) - Len(iOriginalCN) - ii))
                '   �p�G�ŦX�ݭn���s�p�⪺����A�h�p��
                '       varyNoneFixedRow/varyFixedRow ����O�_�ŦX
                If ((varyNoneFixedRow = True And ii = 0) Or (varyFixedRow = True And ii = 1)) Then
                    '�p�G��ܩʰ}�C����J�h�H�}�C���e�p��F�}�C���d�L��� �� �S��J�}�C�h�HcellAddressRowMoveValue�p��
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
                    '�p�G���ܫ᪺�C�Ȥp��1�A�פ�Ҧ��ܰʭp��
                        iii = iii + cellAddressRowMoveValue
                        If (iii < 1) Then
                            isVariedFormulaOutOfCell = True
                            GoTo 881
                        End If
                    End If
                End If
                '   �p�G�ݭn�A�ɦ^"$"
                If (ii = 1) Then
                    iStr3 = "$" & iii
                Else
                    iStr3 = iii
                End If
                rangeArray(3, i) = iStr2 & iStr3
            End If
        Next i
881
        '   �N�}�C���e���s��b�@�_�^��
        Call varyRangeInSheetFunction2Combine(0, 0, cellOriginalFormula, variedFormula, fsnArray, rangeArray, 1, 1)
        
         
    End If
    
''======test��
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

Public Function ��Ʈw_varyRangeInSheetFunction2Combine(ByVal lastArrayWordCounter As Long, ByVal currentWordCounter As Long, ByVal originalFormula As String, ByRef newFormula As Variant, fsnArray, rangeArray, ByVal array1Counter As Long, ByVal array2Counter As Long)
'�Nsheet function���ȫ��^�h���@��
    'main�I�s��func�� lastArrayWordCounter,currentWordCounter ����0�AfsnArray,rangeArray ����1
     
    
    'fsnArray/array2���ŦX�H�U�榡
    '   �Ĥ@�� [1,n]������Ʊq���@�Ӧr�}�l [2,n]������Ʊq���@�Ӧr���� [3,n]������ƪ���
    '   �ĤG�� [n,#]�ĴX�����
    
    'algorithm:�ƼƱq1�}�l�A�C�Ƥ@�ӼƴN�̧��ˬdarray1/2��[1,n1]/[1,n2]�O�_�ŦX�Ӽƭ�
    '   �p�O
    '   (1)�N�������ť�([�W�@���ŦX���}�C��[2,n]����+1]��[�ثe���ƼƼƭ�-1])�H��Ӫ�formula�r���J�sformula�r�ꤤ
    '   (2)�N�}�C����[3,n]��J�sformula�r��
    '   �p�_(arry1/2�����ŦX)�h�~���
Dim skipRest As Boolean
    currentWordCounter = currentWordCounter + 1
    skipRest = False
    '�}�C������Ƥ~�~�����
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

Public Function ��Ʈw_getAllRangesInfoInFormula(ByVal cellOriginalFormula As String, nowSheettName As String) As Variant
'   getAllRangesInfoInFormula�^�Ǥ@��2���}�C
'       [1,0]   [2,0]   [3,n]   [4,0]�O�_�Ҧ���쳣�ݩ�ثe�ɮ� [5,0]�O�_�Ҧ����ثe�ɮפΥثe�u�@��(�`�N�A��!�]�i��O�P�@�i�u�@��)
'       [1,n]��W(A,B,C...) [2,n]�渹�X(1,2,3...) [3,n]�C���X [4,n]�O�_�ݩ�ثe�ɮ� [5,n]�P�ɮסA���O�_�ݩ�ثe�u�@��(�`�N�A��!�]�i��O�P�@�i�u�@��)
'       [n,#]�ĴX�����
'   �p�G�ѷӽd��O��C(1:1)�A[1,n]�^��"ALL"  [2,n]�^�� "1234567890" [3,n]�^�ǸӦC���X
'   �p�G�ѷӽd��O����(A:A)�A[1,n]�^�Ǹ���W [2,n]�^�Ǹ��渹�X      [3,n]�^�� "1234567890"
    
'*�`�N�A�ϥΦ�fun�ɻݥ����n���b�A�T�w�ϥΪ̿�J���O�X�k����ơB���Oexcel�t�ο��~��(#REF!������)�A�_�h�i��|�o�͵L�k�w�������~
'   iVar = IsError(chkThisWS.Cells(i, ii).Value)


    '�N�@���x�s�椺����Ʀ��ѷӨ쪺��쳣�C�X
    '   ��J���ܼ�
    '       cellOriginalFormula �n�B�z���x�s�椺��formula(���O��Ƥ]�S���Y)
    '       nowSheettName �u�@��W�١A�|�P�_�ѷӽd��O�_�O�b���u�@��
    
    '   getAllRangesInfoInFormula�^�Ǥ@��2���}�C
    '       [1,0]   [2,0]   [3,n]   [4,0]�O�_�Ҧ���쳣�ݩ�ثe�ɮ� [5,0]�O�_�Ҧ����ثe�ɮפΥثe�u�@��(�`�N�A��!�]�i��O�P�@�i�u�@��)
    '       [1,n]��W(A,B,C...) [2,n]�渹�X(1,2,3...) [3,n]�C���X [4,n]�O�_�ݩ�ثe�ɮ� [5,n]�P�ɮסA���O�_�ݩ�ثe�u�@��(�`�N�A��!�]�i��O�P�@�i�u�@��)
    '       [n,#]�ĴX�����
    
    '   �p�G�ѷӽd��O��C(1:1)�A�N��W�g�W"ALL",�渹�X"1234567890"
    '   �p�G�ѷӽd��O����(A:A)�A�N�C���X�g�W"1234567890"
    '   �p�G�x�s�檺�Ȥ��]�t����T�AgetAllRangesInfoInFormula�^��1���O5�A�G���u��0���}�C�A�B(4,0) = True�B(5,0) = True
    'algorithm
    '   ���R��Ʈɩ�X�ɮ׸��|�B�ɮצW�١B�ѷӽd�򪺤覡
    '       �N�O "�ɮ׸��| & �ɮצW��" ���r��D�X�A�����_�l��m�B������m
    '       �N�O "�ѷӽd��" ���r��D�X�A�����_�l��m�B������m
    '       ��� "�ѷӽd��" ���_�l��m �P "�ɮ׸��| & �ɮצW��" ��������m�A�o���Ӱѷӽd���ݩ�����ɮסB���Ӥu�@��
    '       �o���ӵ���ƬO�_���@�ӽd���Y/���A�Ϊ̥u�O�@�ӳ�W���x�s��
    '   �N���R���G���x�s���Ƥ@�Ӥ@�Ӯi�}�g�J�}�C�A�����=�Ӱ}�C
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
    
   
    '���b-�n�p�⪺formula���Ȭ���Ƥ~�~��B�z
    If (Left(cellOriginalFormula, 1) <> "=") Then
        ReDim iArray(5, 0)
        iArray(4, 0) = True
        iArray(5, 0) = True
        GoTo 999
    Else
        '�Nformula�ѪR�é�J�}�C

        '   ��XFSN (File & Sheet Name)
        '       FSN�@�w�}�l��B��Ÿ��A�õ����� [!]
        '           �B��Ÿ�: = ( , + - * /
        iStart = 0
        iEnd = 0
        iCount1 = 0
        iCount2 = 0
        iQuotationMarkStart = 0
        iQuotationMarkEnd = 0
        ReDim fsnArray(5, 0)
        For i = 1 To Len(cellOriginalFormula)
            iStr1 = Mid(cellOriginalFormula, i, 1)
            'FSN���ȥ����]�i��]�t�B��Ÿ�,���J��o�Ǫ��p�|�Q��޸�(')���k�]�_
            If (iStr1 = "'") Then
                iCount1 = iCount1 + 1
                If (iCount1 Mod 2 = 1) Then
                    iQuotationMarkStart = i
                    'iQuotationMarkEnd�b�S���e�O���MStart�@��
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
            '   �Ĥ@�� [1,n]������Ʊq���@�Ӧr�}�l [2,n]������Ʊq���@�Ӧr���� [3,n]������ƪ��� [4,n]�O�_�ݩ�P�@���ɮ� [5,n]�O�_�ݩ�P�@�Ӥu�@��
            '   �ĤG�� [n,#]�ĴX�����
            If (iStart <> 0 And iEnd <> 0) Then
                iCount2 = iCount2 + 1
                ReDim Preserve fsnArray(5, iCount2)
                fsnArray(1, iCount2) = iStart
                fsnArray(2, iCount2) = iEnd
                fsnArray(3, iCount2) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '�p�G���O�� ' �}�Y�����p�A�קK���h�l�Ů�ҥHtrim
                If (iQuotationMarkStart = 0) Then
                    fsnArray(3, iCount2) = Trim(fsnArray(3, iCount2))
                End If
                
                iStart = 0
                iEnd = 0
                iQuotationMarkStart = 0
                iQuotationMarkEnd = 0
            End If
        Next i
        
        '   ��X�C�@��cell/����/��C����ƦbcellOriginalFormula���r����m
        '       ��Ʋզ��覡
        '           cell ���զ�
        '               �^��(+�^��...)+�Ʀr(+�Ʀr....)
        '           ���� ���զ�
        '               �^��:�^��
        '           ��C ���զ�
        '               �Ʀr:�Ʀr
        '           �t��k
        '               �ǳ��ܼ�...
        '                   iDict�Ӽg�J�U�r������ƺ���
        '                       [key]�r�����X [value]
        '                                       A(�^��r��)
        '                                       1(�Ʀr�r��)
        '                                       $($�r��)
        '                                       :(:�r��)
        '                                       =(=�r��)
        '                                       other(��L�Ÿ��r��[�]�t�Ů�])
        '                   iStart�Ӽg�J�@�� cell/����/��C ����Ʀbformula�̪��_�l�r����m
        '                   iEnd�Ӽg�J�@�� cell/����/��C ����Ʀbformula�̪������r����m
        '                   iColonAhead�Ӽg�J "�e���� :�r�� �A�B�|�����r��" �ɬ�True
        '                   iFound�ӼХܾ�Ӧr�����_�l������m���w���
        '                   iPureNumStart�Ӽg�J �e�@�ӬO�Ÿ��r�� ���O�^��B���O�Ʀr�B���O$�r���B���O:�r�� ���Ʀr��Ʀbformula�̪��_�l�r����m
        '
        '                   !! ����ɭԡAiFound��True�ɡA������Cell / ���� / ��C ���ȡF���������A�M��NiStart & iEnd & iPureNumStart �]��0�BiColonAhead ��false
        '
        '               �N�r����������6��
        '                   (1)�^��r�� (2)�Ʀr�r�� (3)$�r�� (4):�r�� (5)=�r�� (6)��L�Ÿ��r��[�]�t�Ů�]
        '               �N��Ӧr��C�@�Ӧr������������X
        '               �N�r��H�v�@�r���˵����覡��Xcell��
        '                   �e�m�ˬd
        '                       ��1�Ӧr���O =�r�� �~�B�z
        '                       �r������m���bfsn���d�򤺤~�B�z
        '                   �J�� ��L�Ÿ��r��[�]�t�Ů�]�A�p�G...
        '                       a.iStart & iEnd ������0  ---> �N iFound �]��True
        '                       b.iStart ��0 �B iEnd ����0�A�N iEnd �]��0�BiPureNumStart �]��0�BiColonAhead ��false
        '                       c.iStart & iEnd ����0�A�N iPureNumStart �]��0�BiColonAhead ��false
        '                       d.iStart ����0 �B iEnd ��0�A�N iStart�]��0�BiEnd �]��0�BiPureNumStart �]��0�BiColonAhead ��false
        '                   �J�� $�r���A�p�G...
        '                       a.iStart �� 0�A�p�G...
        '                           a.�e�@�Ӧr���O :�r�� �A�N iColonAhead �]��True�A�N iStart �]�� �ثe�r����m�A�N iEnd �]�� �ثe�r����m
        '                           b.�e�@�Ӧr�����O :�r���A iStart �]�� �ثe�r����m
        '                   �J�� :�r���A�p�G...
        '                       a.iPureNumStart��0�A�NiEnd�]���e�@�Ӧr�� ---> �N iFound �]��True
        '                       b.iPureNumStart����0�A�NiStart�]��iPureNumStart���ȡA�NiEnd�]���e�@�Ӧr�� ---> �N iFound �]��True
        '                   �J�� �^��r���A�p�G...
        '                       iStart �� 0�A�p�G...
        '                               a.�e�@�Ӧr���O :�r�� �A�N iColonAhead �]��True�A�N iStart �]�� �ثe�r����m�A�N iEnd �]�� �ثe�r����m
        '                               b.�e�@�Ӧr�����O :�r���A iStart �]�� �ثe�r����m
        '                       iStart ���� 0�A�BiColonAhead ��True �A iEnd �]�� �ثe�r����m
        '                   �J�� �Ʀr�r���A�p�G...
        '                       a.iStart �� 0 �B...
        '                           a-1.iPureNumStart�� 0...
        '                               a-1-1.�e�@�Ӧr���O :�r�� �A�N iStart �]�� �ثe�r����m�A�N iEnd �]�� �ثe�r����m
        '                               a-1-2.�e�@�Ӧr���O ��L�Ÿ��r��[�]�t�Ů�] �A�N iPureNumStart �]�� �ثe�r����m
        '                           a-2.iPureNumStart���� 0 �B �U�@�Ӧr���O ��L�Ÿ��r��[�]�t�Ů�]�A�NiEnd�]���ثe�r���B�N iStart �]��iPureNumStart����
        '                       b.iStart ���� 0�A�NiEnd�]���ثe�r��
        '                   �̫��ˬd�A���ץثe���r���O����A�u�n�ثe�r���Oformula�̫�@�Ӧr���A�h�p�G...
        '                       �ثe�r���O �^��r�� �A�B iStart ���ȡAiEnd�S���ȡA�NiEnd�]���ثe�r��  ---> �NiFound �]��True
        '                       �ثe�r���O �Ʀr�r�� �A�B iStart & iEnd �����ȡA�NiEnd�]���ثe�r��  ---> �NiFound �]��True
         
        
        
        
        '   �U�r������ƺ����g�Jdict
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
        '   fsn����m�g�Jdict
        For i = 1 To UBound(fsnArray, 2)
            For ii = fsnArray(1, i) To fsnArray(2, i)
                iFsnCharPosDict.Add ii, ii
            Next ii
        Next i
        '   �N�r��H�v�@�r���˵����覡��Xcell��
        iCount1 = 0
        iStart = 0
        iEnd = 0
        iPureNumStart = 0
        iFound1 = False
        ReDim cellArray(6, 0)
        For i = 1 To Len(cellOriginalFormula)
            '�r������m���bfsn���d�򤺤~�B�z
            iNotFsn = False
            If (UBound(fsnArray, 2) = 0) Then
                iNotFsn = True
            Else
                If (iFsnCharPosDict.Exists(i) = False) Then
                    iNotFsn = True
                End If
            End If
            If (iNotFsn = True) Then
                '�}�l�ѪR
                If (iCharTypeDict.GetValue(i) <> "=") Then
                    If (iCharTypeDict.GetValue(i) = "OTHER") Then
                        '�J�� ��L�Ÿ��r��[�]�t�Ů�]�A�p�G...
                        If (iStart <> 0 And iEnd <> 0) Then
                            'a.iStart & iEnd ������0  ---> �N iFound1 �]��True
                            iFound1 = True
                        ElseIf (iStart = 0) Then
                            'b.iStart ��0 �B iEnd ����0�A�N iEnd �]��0�BiPureNumStart �]��0�BiColonAhead ��false
                            'iStart & iEnd ����0�A�N iPureNumStart �]��0�BiColonAhead ��false
                            iEnd = 0
                            iPureNumStart = 0
                            iColonAhead = False
                        ElseIf (iStart <> 0 And iEnd = 0) Then
                            'd.iStart ����0 �B iEnd ��0�A�N iStart�]��0�BiEnd �]��0�BiPureNumStart �]��0�BiColonAhead ��false
                            iStart = 0
                            iEnd = 0
                            iPureNumStart = 0
                            iColonAhead = False
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "$") Then
                        '�J�� $�r���A�p�GiStart �� 0
                        If (iStart = 0) Then
                            If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                'a.�e�@�Ӧr���O :�r�� �A�N iColonAhead �]��True�A�N iStart �]�� �ثe�r����m�A�N iEnd �]�� �ثe�r����m
                                iColonAhead = True
                                iStart = i
                                iEnd = i
                            ElseIf (iCharTypeDict.GetValue(i - 1) <> ":") Then
                                'b.�e�@�Ӧr�����O :�r���A iStart �]�� �ثe�r����m
                                iStart = i
                            End If
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = ":") Then
                        '�J�� :�r���A�p�G...
                        If (iPureNumStart = 0) Then
                            'a.iPureNumStart��0 > �NiEnd�]���e�@�Ӧr�� ---> �N iFound1 �]��True
                            iEnd = i - 1
                            iFound1 = True
                        ElseIf (iPureNumStart <> 0) Then
                            'b.iPureNumStart����0 > �NiStart�]��iPureNumStart���ȡA�NiEnd�]���e�@�Ӧr�� ---> �N iFound1 �]��True
                            iStart = iPureNumStart
                            iEnd = i - 1
                            iFound1 = True
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "A") Then
                        '�J�� �^��r���A�p�G...
                        If (iStart = 0) Then
                            'iStart �� 0 �A�p�G...
                            If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                'a.�e�@�Ӧr���O :�r�� �A�N iColonAhead �]��True�A�N iStart �]�� �ثe�r����m�A�N iEnd �]�� �ثe�r����m
                                iColonAhead = True
                                iStart = i
                                iEnd = i
                            ElseIf (iCharTypeDict.GetValue(i - 1) <> ":") Then
                                'b.�e�@�Ӧr�����O :�r���A iStart �]�� �ثe�r����m
                                iStart = i
                            End If
                        ElseIf (iStart <> 0 And iColonAhead = True) Then
                            'iStart ���� 0�A�BiColonAhead ��True �A iEnd �]�� �ثe�r����m
                            iEnd = i
                        End If
                    ElseIf (iCharTypeDict.GetValue(i) = "1") Then
                        '�J�� �Ʀr�r���A�p�G...
                        If (iStart = 0) Then
                            'a.iStart �� 0
                            If (iPureNumStart = 0) Then
                                'a-1.iPureNumStart�� 0...
                                If (iCharTypeDict.GetValue(i - 1) = ":") Then
                                    'a-1-1.�e�@�Ӧr���O :�r�� > �N iStart �]�� �ثe�r����m
                                    iStart = i
                                    iEnd = i
                                ElseIf (iCharTypeDict.GetValue(i - 1) = "OTHER") Then
                                    'a-1-2.�e�@�Ӧr���O ��L�Ÿ��r��[�]�t�Ů�] �� =�r�� > �N iPureNumStart �]�� �ثe�r����m
                                    iPureNumStart = i
                                End If
                            ElseIf (iPureNumStart <> 0) Then
                                'a-2.iPureNumStart���� 0 �B �U�@�Ӧr���O ��L�Ÿ��r��[�]�t�Ů�] > �NiEnd�]���ثe�r���B�N iStart �]��iPureNumStart����
                                If (iCharTypeDict.GetValue(i + 1) = "OTHER") Then
                                    iEnd = i
                                    iStart = iPureNumStart
                                End If
                            End If
                        ElseIf (iStart <> 0) Then
                            'b.iStart ���� 0 > �NiEnd�]���ثe�r��
                            iEnd = i
                        End If
                    End If
                End If
                
                '���ץثe���r���O����A�u�n�ثe�r���Oformula�̫�@�Ӧr���A�h�p�G...
                If (i = iCharTypeDict.Count) Then
                    If (iCharTypeDict.GetValue(i) = "A" And iStart <> 0 And iEnd = 0) Then
                        '�ثe�r���O �^��r�� �A�B iStart ���ȡAiEnd�S���� > �NiEnd�]���ثe�r��  ---> �NiFound �]��True
                        iEnd = i
                        iFound1 = True
                    ElseIf (iCharTypeDict.GetValue(i) = "1" And iStart <> 0 And iEnd <> 0) Then
                        '�ثe�r���O �Ʀr�r�� �A�B iStart & iEnd ������ > �NiEnd�]���ثe�r��  ---> �NiFound �]��True
                        iEnd = i
                        iFound1 = True
                    End If
                End If
            End If
            
            'iFound��True�ɡA������Cell / ���� / ��C ���ȡF���������A�M��NiStart & iEnd & iPureNumStart �]��0�BiFound & iColonAhead �]�� false
            '   cellArray(6,n)
            '       �Ĥ@�� [1,n]������Ʊq���@�Ӧr�}�l [2,n]������Ʊq���@�Ӧr���� [3,n]������ƪ��� [4,n]�O�_�ݩ�P�@���ɮ� [5,n]�O�_�ݩ�P�@�Ӥu�@�� [6,n]�O:���Y���γ����O(��J H/T/NA)
            '       �ĤG�� [n,#]�ĴX�����
            If (iFound1 = True) Then
                iCount1 = iCount1 + 1
                ReDim Preserve cellArray(6, iCount1)
                cellArray(1, iCount1) = iStart
                cellArray(2, iCount1) = iEnd
                cellArray(3, iCount1) = Mid(cellOriginalFormula, iStart, (iEnd - iStart + 1))
                '���ƬO�e��O�_��:
                '   �w�]���S��
                cellArray(6, iCount1) = "NA"
                If (iCharTypeDict.GetValue(iStart - 1) = ":") Then
                    '���qcell�e����:
                    cellArray(6, iCount1) = "T"
                Else
                    If (iEnd <> iCharTypeDict.Count) Then
                        '���O�̫�@�Ӧr���A�~�䦹�qcell�᭱�O�_��:
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

    
    
    '�p�G������range��Ƥ~�~��B�z
    If (UBound(cellArray, 2) = 0) Then
        ReDim iArray(5, 0)
        iArray(4, 0) = True
        iArray(5, 0) = True
        GoTo 999
    Else
        '���R�C��FSN�O�_�O��L�ɮסB�Ϊ̦P�ɮצ��O��L�u�@��
        If (UBound(fsnArray, 2) <> 0) Then
            For i = 1 To UBound(fsnArray, 2)
                iNotFsn = False
                '�ѷӭȤ��O�_���ɦW�A�����ܦ�fsn�@�w�O��L�ɮ�
                ii = InStr(fsnArray(3, i), "[")
                If (ii <> 0) Then
                    '�ѷӭȤ����ɦW
                    iii = ii
                    ii = InStr(iii + 1, fsnArray(3, i), "]")
                    If (ii <> 0) Then
                        fsnArray(4, i) = False
                        fsnArray(5, i) = False
                    End If
                Else
                    '�ѷӭȤ����t�ɦW�A�N��u���u�@��W
                        '����-�ѷӭȤ��t�ɦW
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
        '��z�N�Urange�ݩ����fsn��X�ӡA�g�W[4,n]�O�_�ݩ�P�@���ɮ� [5,n]�O�_�ݩ�P�@�Ӥu�@��
        If (UBound(fsnArray, 2) = 0) Then
            'fsnArray�S�Ȯɥ����P�_���n�B�z
            For i = 1 To UBound(cellArray, 2)
                cellArray(4, i) = True
                cellArray(5, i) = True
            Next i
        Else
            For i = 1 To UBound(cellArray, 2)
                iStr1 = Mid(cellOriginalFormula, cellArray(1, i) - 1, 1)
                If (iStr1 = ",") Then
                    '�e�@�Ӧr�O�r��, �A�h���ݩ���@��fsn
                    cellArray(4, i) = True
                    cellArray(5, i) = True
                ElseIf (iStr1 = ":") Then
                    '�e�@�Ӧr�O�_�� : �A�h��range�M�e�@��range�ݩ�P�@��fsn
                    cellArray(4, i) = cellArray(4, i - 1)
                    cellArray(5, i) = cellArray(5, i - 1)
                ElseIf (iStr1 = "!") Then
                    '�e�@�Ӧr�O "!"�A�h�ݩ�Yfsn
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
                    '�����ŦX�H�W�����p�A�N�� "�h��RANGE����L����fsn�B�o�ӨS��"�A�n�B�z
                    cellArray(4, i) = True
                    cellArray(5, i) = True
                End If
            Next i
        End If
    End If
    
    '�qrangeArray�`���Ҧ���ƨϥΨ쪺����T
    '   �p�G�ѷӽd��O��C(1:1)�A�N��W�g�W"ALL",�渹�X"1234567890"
    '   �p�G�ѷӽd��O����(A:A)�A�N�C���X�g�W"1234567890"
    iCount1 = 0
    ReDim iArray(5, 0)
    iArray(4, 0) = True
    iArray(5, 0) = True
    '�w�]��
    '   ������Ƴ��P�ɮסB�P�u�@��
    iArray(4, 0) = True
    iArray(5, 0) = True
    '   iFound1�N����Ʀr��ơAiFound2�N�����r���
    iFound1 = False
    iFound2 = False
    For i = 1 To UBound(cellArray, 2)
        iStr2 = ""
        '�N$�Ÿ�����
        iStr1 = WorksheetFunction.Substitute(cellArray(3, i), "$", "")
        '�T�{�d��ȬO�_����C�ξ���
        '   �յ۴M��Ʀr���
        For ii = 1 To Len(iStr1)
            If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                iFound1 = True
                Exit For
            End If
        Next ii
        '   �յ۴M���r���
        For ii = 1 To Len(iStr1)
            If (IsNumeric(Mid(iStr1, ii, 1)) = False) Then
                iFound2 = True
                Exit For
            End If
        Next ii
        If (iFound1 = True And iFound2 = True) Then
            '��ƭȤ��O��C�ξ���
'            For ii = 1 To Len(iStr1)
'                If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
'                    '�J��O�Ʀr�A�N���Ȫ���W�w�j������
'                    Exit For
'                Else
'                    iStr2 = iStr2 & Mid(iStr1, ii, 1)
'                End If
'            Next ii
            '   �N��C�ȧ�X
            For ii = 1 To Len(iStr1)
                If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                    '�J��O�Ʀr�A�N���Ȫ���W�w�j������
                    Exit For
                Else
                    iStr2 = iStr2 & Mid(iStr1, ii, 1)
                End If
            Next ii
            colName = iStr2
            colNo = convertABCto123(iStr2)
            '   �N�C�ȧ�X
            rowNo = Right(iStr1, Len(iStr1) - Len(colName))
        ElseIf (iFound1 = True And iFound2 = False) Then
            '��ƭȬO��C
            colName = "ALL"
            colNo = 1234567890
            rowNo = CLng(iStr1)
        ElseIf (iFound1 = False And iFound2 = True) Then
            '��ƭȬO����
            colName = iStr1
            colNo = convertABCto123(iStr1)
            rowNo = 1234567890
        End If
        
        '��Ƽg�J�}�C
        '   ���@����Ƥ��ݩ��ɮ�/���u�@��A������
        If (cellArray(4, i) = False) Then
            iArray(4, 0) = False
        End If
        If (cellArray(5, i) = False) Then
            iArray(5, 0) = False
        End If
        '   �ݬO�_�� : �e�����
        ibool1 = cellArray(4, i)
        ibool2 = cellArray(5, i)
        If (cellArray(6, i) = "NA") Then
            '�_�A��range���O�@�ӽd��
            iCount1 = iCount1 + 1
            ReDim Preserve iArray(5, iCount1)
            iArray(1, iCount1) = colName
            iArray(2, iCount1) = colNo
            iArray(3, iCount1) = rowNo
            iArray(4, iCount1) = ibool1
            iArray(5, iCount1) = ibool2
        ElseIf (cellArray(6, i) = "H") Then
            '�O�A��range�O�d���Y�A�N�d�򤺪���Ƴ��g�J
            If (colNo = 1234567890 Or rowNo = 1234567890) Then
                '�B�z�O��C�ξ��檺
                iStr1 = WorksheetFunction.Substitute(cellArray(3, i + 1), "$", "")
                If (rowNo = 1234567890) Then
                    '�p�G���d��ȬO����A�U�@��T��Ƥ@�w�]�O����
                    endColName = iStr1
                    endColNo = convertABCto123(iStr1)
                    endRowNo = 1234567890
                ElseIf (colNo = 1234567890) Then
                    '�p�G���d��ȬO��C�A�U�@��T��Ƥ@�w�]�O��C
                    endColName = "ALL"
                    endColNo = 1234567890
                    endRowNo = CLng(iStr1)
                End If
                '�N�d�򤺪��g�J�}�C
                '   :�e
                iCount1 = iCount1 + 1
                ReDim Preserve iArray(5, iCount1)
                iArray(1, iCount1) = colName
                iArray(2, iCount1) = colNo
                iArray(3, iCount1) = rowNo
                iArray(4, iCount1) = ibool1
                iArray(5, iCount1) = ibool2
                '   :��
                iCount1 = iCount1 + 1
                ReDim Preserve iArray(5, iCount1)
                iArray(1, iCount1) = endColName
                iArray(2, iCount1) = endColNo
                iArray(3, iCount1) = endRowNo
                iArray(4, iCount1) = ibool1
                iArray(5, iCount1) = ibool2
            Else
                '�B�z���O��C�]���O���檺
                '   �U�@�Ӥ@�w�OT�A�����
                iStr2 = ""
                '   �N$�Ÿ�����
                iStr1 = WorksheetFunction.Substitute(cellArray(3, i + 1), "$", "")
                '   �N��C�ȧ�X
                For ii = 1 To Len(iStr1)
                    If (IsNumeric(Mid(iStr1, ii, 1)) = True) Then
                        '�J��O�Ʀr�A�N���Ȫ���W�w�j������
                        Exit For
                    Else
                        iStr2 = iStr2 & Mid(iStr1, ii, 1)
                    End If
                Next ii
                endColNo = convertABCto123(iStr2)
                '   �C�ȧ�X
                endRowNo = Right(iStr1, Len(iStr1) - Len(colName))
                '�N�d�򤺪��g�J�}�C
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
            '�U�@�Ӹ��L
            i = i + 1
        End If
    Next i


999
getAllRangesInfoInFormula = iArray

''======�ˬd��=====
'Dim mySht As Worksheet
'    Set mySht = ThisWorkbook.Worksheets("TESTX")
'    mySht.Cells.ClearContents
'    '���D
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
