VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormShowCellInfoSimple 
   Caption         =   "���~�T���C��"
   ClientHeight    =   3555
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4425
   OleObjectBlob   =   "UserFormShowCellInfoSimple.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserFormShowCellInfoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========
'�}�o��     brucechen1@micb2b.com
'�}�o�_�l�� 2023-03-09
'�ק���   2023-11-10
'=========


''�γ~�G
''   �b�{���B�@�L�{�A�N�S�w���A���x�s����ܦbform�W�A�I��form�W����ƥi����ӳB

''�ϥΤ覡�G
''   �]�w�H�U�����ܼ�
''��ܪ��p�u��
'Public showCellInfoCount As Long
'Public showCellInfoArray()
'    '[1,n] �u�@��W�� [2,n]��W�� [3,n]�C���X [4,n]���p���� [5,n]���p�N�X
'    '[n,#]�ĴX�����
'Public showCellInfoListArray()
'    '[#,0]�ĴX����ƭn��ܪ���T�B�q0�}�l
'    '[n,0]�S�N�q�A�t�XListBox�i�HŪ�����}�C���A�ӳ]
''   �O�_���ˬd�쪬�p
'Public showCellInfoHaveData As Boolean
''   �ˬd���覡
'Public userFormShowCellInfoCheckMode As String
''   �קKform����L�{�����а��椣���n���ʧ@
'Public skipUserFormShowCellInfo As Boolean
''   �bUserFormShowCellInfoSimple�����U���i�D�ϥΪ̳B�z�覡
'Public showCellInfoMsg
''   �ˬd�Ҧ��N�X(�̾ڱM�פ��P�ק�)
'Public checkModePriceAdjustCoefEmpty As String, checkModePickPriceEmpty As String

'===��b�I�s��form���{����===
''  ��b�{������̤W��
''  ���b-�ϥΪ̨S�������~�T������==>����
'    If (IsFormInitialized("UserFormShowCellInfoSimple") = True) Then
'        Unload UserFormShowCellInfoSimple
'    End If

''  �I�s�ˬd�{��
''  �]�w�ˬd�Ҧ�
'   userFormShowCellInfoCheckMode = checkMode1
'   Call userFormShowCellInfoExcuteCheck(userFormShowCellInfoCheckMode)

''   ����L�{�N���p�����g�JshowCellInfoArray
''   showCellInfoCount ��l�Ȭ�0
'   showCellInfoCount = showCellInfoCount + 1
'   ReDim Preserve showCellInfoArray(5,iCount)
'   showCellInfoArray(1, showCellInfoCount) = mySht.Name
'   showCellInfoArray(2, showCellInfoCount) = myCN
'   showCellInfoArray(3, showCellInfoCount) = i
'   showCellInfoArray(4, showCellInfoCount) = "���p����"
'   showCellInfoArray(5, showCellInfoCount) = "���p�N�X"


        
''   �p�GshowCellInfoArray���g�J��T�A�N����Ƭ�ListBox�i�HŪ�����}�C���A ==> ��b�Ӭq�{���̫᭱
'    If (UBound(showCellInfoArray, 2) <> 0) Then
'        ReDim showCellInfoListArray(UBound(showCellInfoArray, 2) - 1, 0)
'        Call transformArryToList(showCellInfoArray, showCellInfoListArray)
'        showCellInfoHaveData = True
'    End If

''  �i�D�ϥΪ��ˬd��F����T��
'    If (showCellInfoHaveData = True) Then
'        msgTitle = "�`�N            "    ' �w�q���D�C
'
'        msgText = "�o�{���p����" + vbLf   ' �w�q�T���C
'        msgText = msgText + "====================================" + vbLf ' �w�q�T���C
'        msgText = msgText + "�а���H�U�ʧ@�G" + vbLf   ' �w�q�T��
'        msgText = msgText + "(1)�ק������ܰ��D" + vbLf   ' �w�q�T��
'        msgText = msgText + "(2)�A�����榹�{��"    ' �w�q�T��
''        msgStyle = vbCritical '���"X"�Ϯ�
''        msgStyle = vbExclamation '���"!"�Ϯ�
''        msgStyle = vbInformation '���"i"�Ϯ�
'        MsgBox msgText, msgStyle, msgTitle
'        UserFormShowCellInfoSimple.Show False
'        GoTo 999
'    End If

''===�H�Ufunciton������ҥ����A���s��Func���Ҳդ�===
'Public Function transformArryToList(originalArray(), listArray())
''�N���w�}�C�������Ƭ�ListBox�i�HŪ�����}�C���A
'Dim i As Long
'Dim iStr1 As String
'    For i = 1 To UBound(originalArray, 2)
'        '�N��T�ꦨ�@�Ӧr��
'        '   [1,n] �u�@��W�� [2,n]��W�� [3,n]�C���X [4,n]���p���� [5,n]���p�N�X
'        '   [n,#]�ĴX�����
'        iStr1 = originalArray(1, i) & "-" & originalArray(2, i) & originalArray(3, i) & " : " & originalArray(4, i)
'        listArray(i - 1, 0) = iStr1
'    Next i
'
'End Function

'Public Function IsFormInitialized(FormName As String) As Boolean
'    '�ˬdUserFormShowCellInfoSimple�O�_�Qinitialized
'    'Does not have the side effect of needing to load the form just to see if it's loaded.
'    Dim myForm As Variant
'    For Each myForm In UserForms
'        If myForm.Name = FormName Then
'            IsFormInitialized = True
'            Exit Function
'        End If
'    Next
'End Function


'Public Function userFormShowCellInfoExcuteCheck(checkMode)
''   �ˬd��ƪ��B�@�g�b��function��
''   �p�G���h���ˬd�覡�A��checkMode(�����ܼƬ�userFormShowCellInfoCheckMode)�����
'Dim totalRow As Long
'Dim iStr1 As String
'Dim i As Long
''====�H�U���d�Ҥ��ħ�n���F��====
''myWorkSheet �Q�ˬd���u�@��
''myCheckThisColumn �Q�ˬd����
''checkMode1/checkMode2 �ˬd�Ҧ��W��-�U���ˬd�\��n�W�ߤ@���ˬd�Ҧ��W��
''dataStartRow �n�ˬd���d�򪺰_�l�C
''================================
'    Erase showCellInfoArray
'    showCellInfoCount = 0
'    '���b-�ϥΪ̨S�������~�T������==>����
'    If (IsFormInitialized("UserFormShowCellInfoSimple") = True) Then
'        Unload UserFormShowCellInfoSimple
'    End If
'    With myWorkSheet
'        totalRow = myDataRows(ThisWorkbook.Name, .Name, myCheckThisColumn, 65536)
'        If (checkMode = checkMode1) Then
'            '�ˬd1 <���򤰻򪬪p...>
'            For i = dataStartRow To totalRow
'                iStr1 = .Range(myCheckThisColumn & i).Value
'                If (Trim(iStr1) = "") Then
'                    showCellInfoCount = showCellInfoCount + 1
'                    ReDim Preserve showCellInfoArray(5, showCellInfoCount)
'                    '[1,n] �u�@��W�� [2,n]��W�� [3,n]�C���X [4,n]���p���� [5,n]���p�N�X
'                    '[n,#]�ĴX�����
'                    showCellInfoArray(1, showCellInfoCount) = myWorkSheet
'                    showCellInfoArray(2, showCellInfoCount) = myCheckThisColumn
'                    showCellInfoArray(3, showCellInfoCount) = i
'                    showCellInfoArray(4, showCellInfoCount) = "�i�D�ϥΪ̤��򪬪p"
'                    showCellInfoArray(5, showCellInfoCount) = checkMode
'                End If
'            Next i
'        ElseIf (userFormShowCellInfoCheckMode = checkMode2) Then
'            '�ˬd1 <���򤰻򪬪p...>
'            '   ���Ƽ��g�ˬd���p�A�H�Φ����D�ɭn��JshowCellInfoArray����
'            Next i
'        End If
'    End With
'End Function


'^^^^��������^^^^


Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
Private Sub UserForm_Initialize()
    ListBox1.Clear
    ListBox1.List = showCellInfoListArray
    skipUserFormShowCellInfo = False
    ListBox1.ListIndex = 0
End Sub

Private Sub ListBox1_Click()
'''�d��
'''�I��C��ɪ��ʧ@
''Dim i As Long
''Dim myMode As String, showCellInfoMsg As String
''    If (skipUserFormShowCellInfo = False) Then
''        '�������ܤ�
''        i = ListBox1.ListIndex + 1
''        '   showCellInfoArray
''        '       [1,n] �u�@��W�� [2,n]��W�� [3,n]�C���X [4,n]���~���p [5,n]���~�N�X
''        '       [n,#]�ĴX�����
''        '   listbox index���X�����l�ĤG��-1
''        myMode = showCellInfoArray(5, i)
''        Select Case myMode
'''�̾ڪ��p��U���ˬd�Ҧ���w�@�q�T���A��bUserFormShowCellInfoSimple���U���i�D�ϥΪ̳B�z�覡
''            Case checkModePriceAdjustCoefEmpty
''                showCellInfoMsg = "�L�իY�Ƥ��i�O�ťաA�Эק�"
''            Case checkModePickPriceEmpty
''                showCellInfoMsg = "����ݭn�H�u�B�z"
''        End Select
''        Label2.Caption = showCellInfoMsg
''        '���ʨ�ӳB
''        Worksheets(showCellInfoArray(1, i)).Select
''        Range(showCellInfoArray(2, i) & showCellInfoArray(3, i)).Select
''        With Selection
''            .Borders(xlDiagonalDown).LineStyle = xlContinuous
''            .Borders(xlDiagonalDown).Weight = xlThick
''            .Borders(xlDiagonalUp).LineStyle = xlContinuous
''            .Borders(xlDiagonalUp).Weight = xlThick
''            Application.Wait Now + 1 / (24 * 60 * 60# * 1)
''            .Borders(xlDiagonalDown).LineStyle = xlNone
''            .Borders(xlDiagonalUp).LineStyle = xlNone
''        End With
''    End If

'�I��C��ɪ��ʧ@
Dim i As Long
Dim myMode As String, showCellInfoMsg As String
    If (skipUserFormShowCellInfo = False) Then
        '�������ܤ�
        i = ListBox1.ListIndex + 1
        '   showCellInfoArray
        '       [1,n] �u�@��W�� [2,n]��W�� [3,n]�C���X [4,n]���~���p [5,n]���~�N�X
        '       [n,#]�ĴX�����
        '   listbox index���X�����l�ĤG��-1
        myMode = showCellInfoArray(5, i)
        Select Case myMode
'�̾ڪ��p��U���ˬd�Ҧ���w�@�q�T���A��bUserFormShowCellInfoSimple���U���i�D�ϥΪ̳B�z�覡
            Case checkModePriceAdjustCoefEmpty
                showCellInfoMsg = "�L�իY�Ƥ��i�O�ťաA�Эק�"
            Case checkModePickPriceEmpty
                showCellInfoMsg = "����ݭn�H�u�B�z"
        End Select
        Label2.Caption = showCellInfoMsg
        '���ʨ�ӳB
        Worksheets(showCellInfoArray(1, i)).Select
        Range(showCellInfoArray(2, i) & showCellInfoArray(3, i)).Select
        With Selection
            .Borders(xlDiagonalDown).LineStyle = xlContinuous
            .Borders(xlDiagonalDown).Weight = xlThick
            .Borders(xlDiagonalUp).LineStyle = xlContinuous
            .Borders(xlDiagonalUp).Weight = xlThick
            Application.Wait Now + 1 / (24 * 60 * 60# * 1)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
    End If
End Sub
Private Sub CommandButton1_Click()
'��ثe���w���ˬd�覡�A���s�ˬd
    Me.Hide
    Call userFormShowCellInfoExcuteCheck(userFormShowCellInfoCheckMode)
    If (UBound(showCellInfoArray, 2) <> 0) Then
        ListBox1.Clear
        ListBox1.List = showCellInfoListArray
        skipUserFormShowCellInfo = False
        ListBox1.ListIndex = 0
        UserFormShowCellInfoSimple.Show False
    Else
        '�ץ�������A�A���ˬd��p�G�S���D��ܪ��T��
''�d��
'        msgTitle = "�T��            "    ' �w�q���D�C
'
'        msgText = "�S���o�{���~" + vbLf   ' �w�q�T���C
'        msgText = msgText + "====================================" + vbLf  ' �w�q�T���C
'        msgText = msgText + "�ЦA������H�U�\��G" + vbLf  ' �w�q�T��
'        msgText = msgText + "�u�@�� [" & operateSN & "] => ���s [���ͳ���]"   ' �w�q�T��
'        msgStyle = vbInformation '���"i"�Ϯ�
'        MsgBox msgText, msgStyle, msgTitle
'        Unload Me
        '        msgTitle = "�T��            "    ' �w�q���D�C
        If (userFormShowCellInfoCheckMode = checkModePriceAdjustCoefEmpty) Then
            msgText = "�S���o�{���~" + vbLf   ' �w�q�T���C
            msgText = msgText + "====================================" + vbLf  ' �w�q�T���C
            msgText = msgText + "�ЦA������H�U�\��G" + vbLf  ' �w�q�T��
            msgText = msgText + "�u�@�� [" & operateSN & "] => ���s [���ͳ���]"   ' �w�q�T��
            msgStyle = vbInformation '���"i"�Ϯ�
            MsgBox msgText, msgStyle, msgTitle
            Unload Me
        ElseIf (userFormShowCellInfoCheckMode = checkModePickPriceEmpty) Then
            msgText = "�S���o�{�ťճB"    ' �w�q�T���C
            msgStyle = vbInformation '���"i"�Ϯ�
            MsgBox msgText, msgStyle, msgTitle
            Unload Me
        End If
    End If
End Sub

