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
'�}�o��     brucechen1@micb2b.com
'�}�o���   2023-03-10
'�ק���   2023-12-22
'=========

'''Sub ��Ʈw_�Ҧ��\��ϥνd��()
'''Dim iDict1 As New ebeDictionary, iDict2 As New ebeDictionary
'''Dim iCount As Long
'''Dim iValue As String
'''Dim iBool As Boolean
'''Dim iArray()
'''    '�s�W��ƨ�dict�� key,value (�Ҭ�string�F���M���\��J�¼Ʀr�A���[�Jdict��|�Q�j���ରstring)
'''    ' **key���i�H����**
'''    iDict1.Add 100.1, 1.5
'''    iDict1.Add "20", "B"
'''    iDict1.Add "AA", "C"
'''    iDict1.Add "��", "D"
'''    iDict1.Add "B", "E"
'''    '�N�@��dict���Ƚƻs��t�@��dict
'''    iDict2.copy iDict1
'''    '���o�ݩ�Count - ��Ƶ���
'''    iCount = iDict1.Count
'''    '��key���ovalue
'''    iValue = iDict1.GetValue("AA")
'''    '�ˬd�Ykey�O�_�s�b
'''    iBool = iDict1.Exists("B")
'''    '���o�Ҧ���Value�A�H1���}�C�����A�F�ݫ��w1���}�C���_�l���X
'''    Erase iArray
'''    iArray = iDict1.GetValues(3)
'''    '���o�Ҧ���Key�A�H1���}�C�����A�F�ݫ��w1���}�C���_�l���X
'''    Erase iArray
'''    iArray = iDict1.GetKeys(3)
'''    '���o�Ҧ���Key�MValue�A�H2���}�C�����A
'''    Erase iArray
'''    iArray = iDict1.GetKeysAndValues
'''    '���o�Ykey�bItemsArray����2�����X
'''    iCount = iDict1.GetPageNo("20")
'''    '�ǥ�PageNo���okey��
'''    iValue = iDict1.GetKeyByPageNo(iCount)
'''    '�ǥ�PageNo���oValue��
'''    iValue = iDict1.GetValueByPageNo(iCount)
'''    '�Ƨ�key - �n��Jascend�ӥѤp�ܤj�ƧǡA��descend�Ӧ��j�ܤp�Ƨ�(ascend�BDescend�j�p�g����)
'''    '   �ƧǤj�p���W�h�� ����>�^��>�¼Ʀr (�¤���/�­^��/���^�ƥ����h�O�@�Ӥ@�Ӧr���)
'''    '       �p�G�O����descend(�V�ӶV�p)�A�ƦC���G���O ����=>�^��=>�¼Ʀr (�¤���/�­^��/���^�ƥ����h�O�@�Ӥ@�Ӧr���)
'''    '       �p�G�O�@��ascend(�V�ӶV�j) �A�ƦC���G�� �¼Ʀr=>�^��=>����   (�¤���/�­^��/���^�ƥ����h�O�@�Ӥ@�Ӧr���)
'''    iDict1.SortKey ("ascend")
'''    '�Hkey�R���S�w���
'''    iDict1.Remove ("A")
'''    '�R���Ҧ���J��
'''    iDict1.RemoveAll
'''
'''End Sub

'===�]�w�ݩʭȤ覡
'Public Property Let iname(ByVal strName As String)
'    mystrName = strName
'End Property

'===���o�ݩʭȤ覡
'Public Property Get iname() As String
'    iname = mystrName
'End Property


Private ItemsCollection As New Collection
Private ItemsArray()
'   �ϥΪ̨��ήɻݦۭq�@��2�����P�Ӧs��myDict���
'   [1,n]key [2,n]value
'   [n,#]�ĴX�����
Private xxxCount As Long

'20231222 ���յo�{�Acollection Add�ɡA�j�p�gM�|�Q�����P�@��;�����S���r��excel�{�����P��collection�{���ۦP
'   �o�Ǫ��p��ebeDictionary����_�Ӥ���O���D�A���i���ϥ�ebeDictionary���{���|�y���v�T
'   ==>�g�@���ܼơA�N�{���J�쪺���p�g�J�A�u�n�ϥ�ebeDictionary���n�b�{���ݬd�ݦ��ܼƶǦ^�����p�O�_�v�T����

'   �٦���L���p

'=====�bsub��function���ϥΨ즹class�ɪ��_�l�]�m
Private Sub Class_Initialize()
    xxxCount = 0
    ReDim ItemsArray(2, xxxCount)
End Sub
'=====�ϥΨ즹class�ɪ�sub��function�����ɪ������]�m
Private Sub Class_Terminate()

End Sub
Public Property Get Count() As Long
'�`�N�AxxxCounts���H��Remove��s
    Count = xxxCount
End Property

'======��k
Friend Function Add(ByVal Key As String, ByVal Value As String)
'�s�W��ơAkey���i���ơAvalue�i����
'�|�Nkey & value�[�J�ݩ�
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
'�N�t�@dict���Ƚƻs�즹dict
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
'�����YKey�O�_�s�b
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
'���o�Ҧ���Key�A�H�@���}�C���A�^�ǵ�GegKeys�F�@���}�C�qArrayIndexStartNumber�}�l
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
'���o�Ykey�bItemsArray����2�����X
Dim i As Long
Dim myKey
'20231222:�쥻�ϥ��ݩ�ItemsArray�ӭp��,���o�{�����S���r��excel�{�����P��collection�{���ۦP�A�G�N�ϥ�ItemsArray���P�_���令�ϥ�collection�ۤv���覡
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
'   �~��ק��L���A�N�ϥ�ItemsArray���P�_���令�ϥ�collection�ۤv���覡
End Function
Friend Function GetKeyByPageNo(ByVal pageNo As Long) As Variant
'�ǥ�PageNo(ItemArray�ĤG�������X)���okey��
Dim i As Long
    GetKeyByPageNo = ItemsArray(1, pageNo)
End Function
Friend Function GetValueByPageNo(ByVal pageNo As Long) As Variant
'�ǥ�PageNo(ItemArray�ĤG�������X)���oValue��
Dim i As Long
    GetValueByPageNo = ItemsArray(2, pageNo)
End Function
Friend Function GetValue(ByVal Key As String)
'��key�Ȩ��o��value
'�p�G��key���s�bcollection���A�h�^�Ǫŭ�""
    On Error Resume Next
        GetValue = ItemsCollection.Item(Key)
    On Error GoTo 0
End Function
Friend Function GetValues(ArrayIndexStartNumber As Integer) As Variant
'���o�Ҧ���Value�A�H�@���}�C���A�^�ǵ�GetValues�F�@���}�C�qArrayIndexStartNumber�}�l
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
'���o�Ҧ���Key�MValue�A�H�G���}�C���A�^�ǵ�GetKeysAndValues�A��1���q1�}�l
'   [1,n]key [2,n]Value
'   [#,n]�@���X�ո��
    GetKeysAndValues = ItemsArray
End Function
Friend Function RemoveAll()
'�N�Ҧ���key�PValue�R��
    'ItemsArray
    Erase ItemsArray
    'ItemsCollection
    Set ItemsCollection = Nothing
    'xxxCount
    xxxCount = 0
End Function
Friend Function Remove(Key As String)
'�Hkey�N�S�w����ƧR��
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
'�Nkey���ƧǡA�ϥΪ̶���J"ascend"��"descend"(�j�p�g����)�ӫ��w�ƧǤ覡
    '���b-�ϥΦ�function�A���ܼ� ascendOrDescend ���e�g��
    If (UCase(ascendOrDescend) <> "ASCEND" And UCase(ascendOrDescend) <> "DESCEND") Then
        MsgBox "ebeDictionary��function [SortKey]����J�ܼƦ��~"
    End If
    '���b-dict���S����ơA�������}function
    If (xxxCount = 0) Then
        Exit Function
    End If
    
    '�NItemsArray��key�P��}�C���X�g�iiCollection1�ѫ����value�� (value = �}�C���X,key = key)
    For i = 1 To UBound(ItemsArray, 2)
        iColllection1.Add CStr(i), ItemsArray(1, i)
    Next i

    '�C2����ƴN�biArray���]���@��(�p�h�X�Ӫ��ۦ�1��)�A��2���J�}�C��
    '   [1,n]�o�ո�Ƥ����j�� [2,n]�o�ո�Ƥ����p�� [3,n]�ثe�������H
    '   [n,#]�ĴX�����
    iCount2 = 0
    For i = 1 To UBound(ItemsArray, 2)
        If (i = 1 Or iCount1 = 2) Then
            iCount1 = 1
        Else
            iCount1 = iCount1 + 1
        End If
        '�����Xkey���A�p�G�O�Ʀr�N�ରDbl
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
        '�N�C�ռƦr��j�p�A�j����[1,n]�B�p����[2,n]
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
                '�p�G�̫�@�եu���@�ӭȡA���[1,n]
                iArray1(1, iCount2) = iVar1
            End If
            '�N�n����o�����[3,n]
            If (UCase(ascendOrDescend) = "ASCEND") Then
                '�@��,��p��
                If (iArray1(2, iCount2) <> "") Then
                    iArray1(3, iCount2) = iArray1(2, iCount2)
                Else
                    iArray1(3, iCount2) = iArray1(1, iCount2)
                End If
            ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                '���ǡA��j��
                iArray1(3, iCount2) = iArray1(1, iCount2)
            End If
        End If
    Next i

    '�]�w�@��Collection�AValue��J�H�Ϥ����@�ժ��}�C����2�����X,key��J�}�C[1,n](key��)
    '   ������j�p�ɡA�w�T�w���Ǫ̱q��Collection�����A�H�ѫ���{�����D�}�C�����ӵ���ƬO�_�w��粒��
    For i = 1 To UBound(iArray1, 2)
        iColllection2.Add CStr(i), CStr(iArray1(3, i))
    Next i
    
    '�NiArray1������ƥ�����X�j�p
    '   �C����X�@�ӭ�(�ثe�̤j/�̤p)�A�ҥH�`���榸�Ƶ����Ƶ���
    iCount1 = 0
    For i = 1 To UBound(ItemsArray, 2)
        iVar1 = ""
        For ii = 1 To UBound(iArray1, 2)
            '�ˬd�Ӱ}�C�O�_������
            '�O�_�s�b
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
                        '�@��,��p��
                        If (iArray1(3, ii) < iVar1) Then
                            iVar1 = iArray1(3, ii)
                        End If
                    ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                        '���ǡA��j��
                        If (iArray1(3, ii) > iVar1) Then
                            iVar1 = iArray1(3, ii)
                        End If
                    End If
                    
                End If
            End If
701
        Next ii
        '�N�ثe��쪺�̤j/�̤p�Ȫ�key�Mvalue�g�i�}�C
        iCount1 = iCount1 + 1
        ReDim Preserve iArray2(2, iCount1)
        iArray2(1, iCount1) = CStr(iVar1)
        iArray2(2, iCount1) = ItemsArray(2, iColllection1.Item(CStr(iVar1)))
        '���s��zcollection �M ���s�w�qiArray(3,n)
        ii = CLng(iColllection2.Item(CStr(iVar1)))
        iDoIt = False
        If (iArray1(2, ii) <> "") Then
            If (UCase(ascendOrDescend) = "ASCEND") Then
                '�@��,�N�j����J[3,n]
                If (iArray1(3, ii) <> iArray1(1, ii)) Then
                    iArray1(3, ii) = iArray1(1, ii)
                    iDoIt = True
                End If
            ElseIf (UCase(ascendOrDescend) = "DESCEND") Then
                '����,�N�p����J[3,n]
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
    '�N��z�᪺iArray2��Ƽg�^dict
    '   ItemsArray
    ItemsArray = iArray2
    'xxxCount����
    'ItemsCollection
    Set ItemsCollection = Nothing
    For i = 1 To UBound(iArray2, 2)
        ItemsCollection.Add CStr(iArray2(2, i)), CStr(iArray2(1, i))
    Next i
End Function