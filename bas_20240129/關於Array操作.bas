Attribute VB_Name = "����Array�ާ@"
Sub ����Array�����ާ@()
'�N�}�C�����r��b�@�_,�`�N,�}�C�O�q0�}�l��
Dim myA(2)
myA(1) = 1
myA(2) = 2
myStr = Join(myA, ",")

'�N�}�C�����ȼg�i�x�s��
'   �}�C�Ҧ����׶��q0�}�l��J
'   �i�H�O�h���ת��}�C�AĴ�p�U�C
    Dim iArray1(1, 1)
    iArray1(0, 0) = "00"
    iArray1(0, 1) = "01"
    iArray1(1, 0) = "10"
    iArray1(1, 1) = "11"
    Range("AB1:AC2").Value = iArray1
'   1���}�C�����p�U
    '�p�Grange�W�L1�ӡA�}�C�u���@����ơA�|�N�Ҧ�range��J�Ӱ}�C��
    Range("A1:C1").Value = Array("1")
    '�p�Grange�W�L1�ӡA�}�C�]�W�L1�ӡA�h�|�N�}�C������ƨ̧Ǽg�Jrange�F�}�C���h�l���|�Q�˥h
    '   A1~C1���Q��J�۹�������1~3�A�̫᪺4���|�Q��J
    Range("A1:C1").Value = Array("1", "3", "3", "4")
    '�ϥ�Range.value = array�A�u�O�Φb������Range
    '   Ĵ�pA1~C1
    '   �p�GRange�O�������A���[�WApplication.Transpose
    Range("A1:A3").Value = Application.Transpose(Array("1", "2", "3", "4"))
End Sub

'======�ۼg�\��(1-1���}�C-1) ��@���}�C�ƧǨò������ƪ����
Public ��Ʈw_transpose2DimensionArray(myArray)
'�N�ǤJ���}�C��m�AĴ�p arr(5,100) �ܦ� arr(100,5)�A�̤����Ȥ]��m
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


Function ��Ʈw_myUniqueSort1DimArray(ByVal inputArray As Variant, ascendOrDescend As String)
'�`�N�A�P�_�O�_���ƪ��Ȧp�G�O�¼Ʀr�ɡA�n�ϥξ��(�pLONG)���n�ϥΤp�Ʀ쫬�A(�pDOUBLE)�A�]��EXCEL�b�p��DOUBLE�ɷ|���ͫD�`�p���t���A�ɭP�ݰ_�Ӥ@�˪��ȹ�EXCEL�ӻ��o���@��

'��J���:inputArray ��l��Ʈư}�C / ascendOrDescend ��Jascend(�C�ܰ��Ƨ�)��descend(���ܧC�Ƨ�)
'inputArray������Ʈ榡: �@���}�C�B�Ʀr or �^�� or �V�M (��r���A���Ʀr�AĴ�p"10"�A�|�Q��@��r�B�z)�B�i�����Ƹ�ơB�i���Ů�
'�ƧǤ覡: �̨ϥΪ̨M�w�n�C>��(�¼Ʀr�b�e�A1 to 9 then A to Z) �� ��<�C (�D�¼Ʀr�b�e�AZ to A then 9 to 1)
'�Ƨǫ�|�Q�簣��: trim��O�Ů�("")�B���ƪ����
Dim i As Long, ii As Long, iii As Long
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()

    '�Ƨ�-�P�ˤj�p���Ʀb�@�_�A�ϱo���X�ӭ��ƴN���X�ӪŮ沣��
    ReDim sortWithDuplicateArray(UBound(inputArray, 1))
    For i = 1 To UBound(inputArray, 1)
        iii = 1
        For ii = 1 To UBound(inputArray, 1)
            If (i <> ii) Then
                '�Ѱ��ܧC �� �ѧC�ܰ�
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
    '�����]���ƾɭP���Ů�
    '   �p�⦳�X�ӪŮ�
    emptyCount = 0
    For i = 1 To UBound(sortWithDuplicateArray, 1)
        If (Trim(sortWithDuplicateArray(i)) = "") Then
            emptyCount = emptyCount + 1
        End If
    Next i
    '   �����Ů�
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

'======�ۼg�\��(1-2���}�C-1) ��G���}�C�ƧǨò������ƪ����
Function ��Ʈw_myUniqueSort2DimArray(ByVal inputArray As Variant, ascendOrDescend As String, dataInWhichDim As Long, sortByWhichAttribute As Long)
'�`�N�A�P�_�O�_���ƪ��Ȧp�G�O�¼Ʀr�ɡA�n�ϥξ��(�pLONG)���n�ϥΤp�Ʀ쫬�A(�pDOUBLE)�A�]��EXCEL�b�p��DOUBLE�ɷ|���ͫD�`�p���t���A�ɭP�ݰ_�Ӥ@�˪��ȹ�EXCEL�ӻ��o���@��

'��J���:inputArray ��l��Ʈư}�C / ascendOrDescend ��Jascend(�C�ܰ��Ƨ�)��descend(���ܧC�Ƨ�)
'inputArray������Ʈ榡: �G���}�C�B�Ʀr or �^�� or �V�M (��r���A���Ʀr�AĴ�p"10"�A�|�Q��@��r�B�z)�B�i�����Ƹ�ơB�i���Ů�
'�ƧǤ覡: �̨ϥΪ̨M�w�n�C>��(�¼Ʀr�b�e�A1 to 9 then A to Z) �� ��<�C (�D�¼Ʀr�b�e�AZ to A then 9 to 1)
'�Ƨǫ�|�Q�簣��: trim��O�Ů�("")�B���ƪ����
Dim attributeInThisDim
Dim i As Long, ii As Long, iii As Long, iv As Long
Dim iIsEmpty As Boolean
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()
    '���o�ݩʩҦb���פ�K����ϥ�
    If (dataInWhichDim = 1) Then
        attributeInThisDim = 2
    ElseIf (dataInWhichDim = 2) Then
        attributeInThisDim = 1
    End If
    '�Ƨ�-�P�ˤj�p���Ʀb�@�_�ƶi�}�C�� > ���X�ӭ��ư}�C���N���X��O�Ů�
    ReDim sortWithDuplicateArray(UBound(inputArray, 1), UBound(inputArray, 2))
    For i = 1 To UBound(inputArray, dataInWhichDim)
        iii = 1

        For ii = 1 To UBound(inputArray, dataInWhichDim)

            If (i <> ii) Then
                '�Ѱ��ܧC �� �ѧC�ܰ�
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
    '�����]���ƾɭP�}�C���s�b���Ů� (�C���ݩʳ��O�Ů�~��O)
    '   �p�⦳�X�ӪŮ�
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
    '   �����Ů�
    If (dataInWhichDim = 1) Then
        ReDim sortWithoutDuplicateArray(UBound(sortWithDuplicateArray, 1) - emptyCount, UBound(sortWithDuplicateArray, 2))
    ElseIf (dataInWhichDim = 2) Then
        ReDim sortWithoutDuplicateArray(UBound(sortWithDuplicateArray, 1), UBound(sortWithDuplicateArray, 2) - emptyCount)
    End If

    iii = 0
    For i = 1 To UBound(sortWithDuplicateArray, dataInWhichDim)
        '��X�Ů�� - �ѩ�ΨӱƧǪ��ݩʤ]�i��O�ŭȡA�G�]�p���Ҧ��ݩʳ� ���O �ŭȤ~�P�w���n�O�d
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

'======�ۼg�\��(1-2���}�C-2) ��G���}�C�Ƨ�-�p�ƧǪ��ƩʭȬۦP�ɨ̭�l�}�C���e�ᶶ�ǱƦb�@�_
Function ��Ʈw_mySort2DimArray(ByVal inputArray As Variant, ascendOrDescend As String, dataInWhichDim As Long, sortByWhichAttribute As Long)
'�`�N�A�P�_�O�_���ƪ��Ȧp�G�O�¼Ʀr�ɡA�n�ϥξ��(�pLONG)���n�ϥΤp�Ʀ쫬�A(�pDOUBLE)�A�]��EXCEL�b�p��DOUBLE�ɷ|���ͫD�`�p���t���A�ɭP�ݰ_�Ӥ@�˪��ȹ�EXCEL�ӻ��o���@��

'��J���:inputArray ��l��Ʈư}�C / ascendOrDescend ��Jascend(�C�ܰ��Ƨ�)��descend(���ܧC�Ƨ�)
'inputArray������Ʈ榡: �G���}�C�B�Ʀr or �^�� or �V�M (��r���A���Ʀr�AĴ�p"10"�A�|�Q��@��r�B�z)�B�i�����Ƹ�ơB�i���Ů�
'�ƧǤ覡: �̨ϥΪ̨M�w�n�C>��(�¼Ʀr�b�e�A1 to 9 then A to Z) �� ��<�C (�D�¼Ʀr�b�e�AZ to A then 9 to 1)
'�Ƨǫ�|�Q�簣��: trim��O�Ů�("")�B���ƪ����
Dim attributeInThisDim
Dim i As Long, ii As Long, iii As Long, iv As Long
Dim iIsEmpty As Boolean
Dim emptyCount As Long
Dim sortWithDuplicateArray(), sortWithoutDuplicateArray()
    '���o�ݩʩҦb���פ�K����ϥ�
    If (dataInWhichDim = 1) Then
        attributeInThisDim = 2
    ElseIf (dataInWhichDim = 2) Then
        attributeInThisDim = 1
    End If
    '�Ƨ�-�P�ˤj�p���̷ӭ���e�ᶶ�ǱƦb�@�_�ƶi�}�C��
    '   *�p�G�n�N���ǬۤϡA�h�U�� If (i > ii) Then �令 If (i < ii) Then
    ReDim sortWithDuplicateArray(UBound(inputArray, 1), UBound(inputArray, 2))
    For i = 1 To UBound(inputArray, dataInWhichDim)
        iii = 1
'If (i = 1022) Then
'iIsEmpty = False
'End If
        '���Ƨǫ᪺�s��m
        For ii = 1 To UBound(inputArray, dataInWhichDim)
'If (ii = 1021) Then
'iIsEmpty = False
'End If
            If (i <> ii) Then
                '�Ѱ��ܧC �� �ѧC�ܰ�
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
        '�N��Ƽg�J�s�}�C�����s��m
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

'======�ۼg�\��(2) ��@���}�C�������ƿz��(���ӭ�l����Ƨ�)
Function ��Ʈw_make1DimArrayUnique(inputArray, arrayStartNum As Long)
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

'======�ۼg�\��(3) ��G���}�C�������ƿz��(���ӭ�l����Ƨ�)
Function ��Ʈw_make2DimArrayUnique(inputArray, dataInWhichDim As Long, filterInWhichAttribute As Long, arrayStartNum As Long)
Dim attributesTotalNum As Long
Dim tempArray()
Dim i As Long, ii As Long, iCount As Long
Dim iFound As Boolean
    'dataInWhichDim: 1 or 2 ,��n����ƪ�n�\�b��1���٬O��2��
    '   1: myArray(n,1)
    '   2: myArray(1,n)
    'filterInWhichAttribute: 1 to n ,�C����ƲĴX���ݩʬO�����ƿz�諸�̾�
    'arrayStartNum: 0 to m ,��ƭ��Ӻ��ת��ĴX�Ӷ}�l��m
    ' <�|��>
    '   ��1���O�ݩ� (�@��2���ݩ� �W�l,�q��)�A��2���O��
    '   myArray(1,0) ��0����ƪ���1���ݩʬO�Ĥ@�ӤH���W�l
    '   myArray(2,0) ��0����ƪ���2���ݩʬO�Ĥ@�ӤH���q��
    '   myArray(1,1) ��1����ƪ���1���ݩʬO�ĤG�ӤH���W�l
    '   myArray(2,1) ��1����ƪ���2���ݩʬO�ĤG�ӤH���q��
    '   �H�H�W�������ƿz��A�h
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
        '��1�����
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
        '��L�����
        Else
            iFound = False
            '��惡����ƩM�e������ƬO�_����
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
            '�S���ƫh����
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
'======�ۼg�\��(4) �b�G���}�C���H��ӱ���ȧ��S�w���h�����[datas]�F����Ҧ�[datas]���Y�Ʀr�ݩʪ��ȡA�ݽ֪��ȳ̤j�F�H�ӳ̤j�ȼg�^�Ҧ�[datas]�����ݩ�
Function ��Ʈw_getMaxNumberThenOverwriteInArrayWith2Filter(inputArray, dataInWhichDim As Long, dataInWhichAttribute, ByVal filterInWhichAttribute1 As Long, ByVal filterString1 As String, ByVal filterInWhichAttribute2 As Long, ByVal filterString2 As String)
Dim iVal As Double, currentMaxVal As Double
Dim i As Long, ii As Long, iCount As Long
Dim iStr1 As String, iStr2 As String
Dim tempArray()
    'dataInWhichDim: 1 or 2 ,��n����ƪ�n�\�b��1���٬O��2��
    '   1: myArray(n,1)
    '   2: myArray(1,n)
    'dataInWhichAttribute: 1 to n ,�C����ƲĴX���ݩʬO�s��n��̤j�Ȫ��Ʀr
    'filterInWhichAttribute: 1 to n ,�C����ƲĴX���ݩʬO�����ƿz�諸�̾�
    'filterValue:�z���
    
    '�z��X�}�C���ŦX����̨ç��̤j��
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
    '�γ̤j�ȧ�ginputArray
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

Public Function make1DArrayToCsvString_��Ʈw(myArray)
'�N�}�C (�@���A�q1�}�l)���ȥH,��b�@�_��CSV�榡
Dim i As Long
    For i = 1 To UBound(myArray, 1)
        If (i = 1) Then
            make1DArrayToCsvString = myArray(i)
        Else
            make1DArrayToCsvString = make1DArrayToCsvString & "," & myArray(i)
        End If
    Next i
End Function
