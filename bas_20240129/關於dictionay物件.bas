Attribute VB_Name = "����dictionay����"
Sub Dict���ާ@_��Ʈw()
Attribute Dict���ާ@_��Ʈw.VB_ProcData.VB_Invoke_Func = " \n14"
'�`�N!!
'�ϥ� �ʬݦ� �b�ˬddict���e�ɡA�u��ݨ�256�����

'https://excelmacromastery.com/vba-dictionary/
'Dict�ΨӰO�����
'ArrayList�ΨӱƧ�
Dim iVar As Variant
Dim iDict As Object, iDictNew As Object

    '�Ѧ�   https://excelmacromastery.com/vba-dictionary/
    '       https://excelmacromastery.com/vba-arraylist/
'===late binding
    Set iDict = CreateObject("Scripting.Dictionary")
    Set iDictNew = CreateObject("Scripting.Dictionary")

'===dicionary�O�_�Ϥ��j�p�g
    '�]�wiDict�O�_case sensitive
    '   �O(�w�])
    iDict.CompareMode = vbBinaryCompare
    '   �_
    iDict.CompareMode = vbTextCompare

'***�s�y�@�Ǹ�ƴ��ե�
    iDict.Add "key1", "Value1"
    iDict.Add "key2", "Value2"
    iDict.Add "key3", "Value3"
    
'===dicitonary����Ƶ���
'   �S���(�Q.RemoveAll)=0,add�L1��=1
    iVar = dict.Count
    
'===�d�߬Y��key�O�_�s�b
    If (iDict.Exists("key1") = True) Then
        '...
    End If
'===�NiDict�����e�ݤ@�M
    For Each iVar In iDict
        '���okey
        mykey = iVar
        '���oValue
        myVal = iDict(iVar)
    Next iVar
    
    '����dict���e
    i = 0
    For Each iVar In iDict
        i = i + 1
        With Worksheets("test")
            '���okey
            .Cells(i, 1) = iVar
            '���oValue
            .Cells(i, 2) = iDict(iVar)
        End With
    Next iVar

'===��L
    '���odict���Ykey��value
    myVal = iDict("key1")
    'dict������Ƽƶq(�S�������ƮɬO0)
    myVal = iDict.Count
    '����dict���S�w���(dict��Ƽƶq�|����)
    iDict.Remove "key1"
    '����dict���Ҧ����(dict��Ƽƶq�|����)
    iDict.RemoveAll
    
    '�ק�S�wKey�̪�Value
    '   �S��k�A�u��N��key�����A���s�[�J
    
    '���odict���Ҧ�key��Value���̤j��/�̤p��
    MsgBox Application.max(iDict.items)
    MsgBox Application.min(iDict.items)
    '�Ndict����]��nothing - ���O�N��Ʋ����A�ӬO�N���ܼƳ]�����A�O�r�媫��
    '   �����s�w�q�o�Ӫ���A�]��nothing�i��ְ���ɶ�
    Set iDict = CreateObject("Scripting.Dictionary")
    Set iDict = Nothing
    If (iDict Is Nothing = True) Then
       MsgBox 1
    End If
    
'===�ϥ�ArrayList�Ndictionay�Ƨ�
'**�`�N** �p�G�q���S�w�� .NET Framwork3.5 �|�L�k�ϥ�arrlist
'https://stackoverflow.com/questions/40625618/automation-error-2146232576-80131700-on-creating-an-array
Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")

    For Each iVar In iDict
        arrList.Add iVar
    Next iVar
    '�Ƨ�-1~9,A~B
    arrList.Sort
    '�N�ثe���G�A�˱ƦC
    arrList.Reverse
    '�N�Ƨǵ��G���s��dictionary����
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
