Attribute VB_Name = "����z��P�Ƨ�"
Sub �Ұʿz��_��Ʈw()
    ROWS("1:1").Select
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
        Selection.AutoFilter
        Else
        Selection.AutoFilter
    End If
End Sub

Sub ���Ѱ��z�窱�A�U�����z��_��Ʈw()
On Error Resume Next
Sheets("123").Select
ROWS("1:1").Select
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
End If
On Error GoTo 0
End Sub

Sub �ۭq�Ƨǥ\��_��Ʈw()
Dim mySht As Worksheet
Dim customListOriginalCount As Long
Dim i As Long
'�覡(1)�ϥΤ@�өΦh�Ӽg������ ���ƧǨ̾�
    mySht.Sort.SortFields.Add Key:=Range("A2:A10"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="�T��,����", DataOption:=xlSortNormal
    '   ����Ƨ�
    With mySht.Sort
        .SetRange Range("A2:C10")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'�覡(2)�ϥ��ܰʪ��� ���ƧǨ̾�
'   ��z (1)�N�n�ƧǪ��� �[�J �ۭq�M�� (�P �ɮ�>�ﶵ>�i��>�@��>�s��ۭq�Ƨ�)
'        (2)�b�ƧǤ�k���ϥΦۭq�M�氵�ƧǨ̾�
'        (3)�b�ۭq�M�椺�R���ӭ�
    '�p���l�ۭq�M�椺���h�ֵ����
    customListOriginalCount = Application.CustomListCount
    '�N�n�[�J �ۭq�M�� ���ȼg�JArray
    sortOrderArray(1) = "�T��"
    sortOrderArray(2) = "����"
    '�s�W �ۭq�M��
    Application.AddCustomList ListArray:=sortOrderArray
    '�ϥ� �ۭq�M�� ���̫�@�����(�]�N�O�W�z�s�W��)���ƧǨ̾�
    mySht.Sort.SortFields.Add Key:=Range("A2:A10"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=Application.CustomListCount, DataOption:=xlSortNormal
    '   ����Ƨ�
    With mySht.Sort
        .SetRange Range("A2:C10")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '�R���s�W��CustomList
    '   ��ϥΦۭq�M��\���i��@�s��excel�N����A�����W���[�W����N���|(�w�g�L�����ҹ�)
    mySht.Sort.SortFields.Clear
    For i = Application.CustomListCount To customListOriginalCount + 1 Step -1
        Application.DeleteCustomList i
    Next i

End Sub
