Attribute VB_Name = "����SHEET�ާ@"
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Private Sub Worksheet_Change(ByVal Target As Range)
'�u�@��S�w�a��Q���ܡA�hĲ�o�{��
    Set editRg = Target
    If (Target.address = dateSelectYearRg.address) Then
        For Each iRg In editRg
        Next
    End If
End Sub
Sub �b���ɤ��R���ª�÷s�W����_��Ʈw()
Dim mySht As Worksheet
    '�j��R���ª�
    Application.DisplayAlerts = False
        On Error Resume Next
            Sheets("�ϯä��R").Delete
        On Error GoTo 0
    Application.DisplayAlerts = True
    '�H�ƻs�覡���ͤu�@��
    '   ����ĵ�i�A�H�K�]���榡�ΦW�ٸ��X�@��T��
    Application.DisplayAlerts = False
        '�ƻs��̫�@�i�u�@��᭱
        Sheets("��l��").copy After:=Sheets(Sheets.Count)
        '�ƻs��S�w�u�@��᭱(�ݥ�.index���o�ӱi�u�@��s��)
        Sheets("��l��").copy After:=Sheets(mySht.Index)
    Application.DisplayAlerts = True
    '�H�s�W�覡���ͤu�@��
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.Count)
    ActiveSheet.Name = "�ϯä��R"
End Sub
Sub �ˬd�n�s�W����O�_�s�b()
Dim iStr1 As String
    iStr1 = "�n�s�W���u�@��W��"
    For Each iVar1 In ThisWorkbook.Worksheets
        If (iVar1.Name = iStr1) Then
            ibool1 = True
        End If
    Next iVar1
    

    '           �p�G�s�b�A��ܱN�R���̭�����ơA�Ϊ̫ݨϥΪ̭ק�u�@��W��
    If (ibool1 = True) Then
        msgTitle = "�`�N           "    ' �w�q���D�C
        msgText = "�H�U�u�@��N [�R��] �í��s����" + vbLf  ' �w�q�T���C
        msgText = msgText + "====================================" + vbLf ' �w�q�T���C
        msgText = msgText + iStr1 + vbLf
        msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
        msgText = msgText + "   ==>�p�G�P�N�A�п��[�T�w]�F�ο��[����]���}�{��" + vbLf   ' �w�q�T���C"
        msgText = msgText + "       *��ܨ����̡A�Эק� �W�z�u�@���W�� �A�M��A������" + vbLf    ' �w�q�T��
        answer = MsgBox(msgText, vbOKCancel + vbExclamation, msgTitle)
        If answer = vbCancel Then
            GoTo 993
        Else
            '�R���u�@��
            Application.DisplayAlerts = False
                ThisWorkbook.Worksheets(iStr1).Delete
            Application.DisplayAlerts = True
        End If
    End If
    '       �s�W�ӱi�u�@��
    Application.DisplayAlerts = False
        ThisWorkbook.Sheets(eachMonthBaseSN).copy After:=ThisWorkbook.Sheets(Sheets.Count)
    Application.DisplayAlerts = True
    ActiveSheet.Name = iStr1
    Set eachMonthSht = ThisWorkbook.Worksheets(iStr1)

End Sub
Sub SHEET���ҤW��_��Ʈw()
'SHEET�W��:��
Sheets("SYSTEM").Tab.Color = 6299648
'SHEET�W��:��
Sheets("Data").Tab.Color = 65535
'SHEET�W��:��
Sheets("�ϯä��R").Tab.Color = 65535
'SHEET�W��:��
Sheets("ALLOCATION LIST").Tab.Color = 255
End Sub
Sub ���ϥ�activesheet_�s�Wsheet_��Ʈw()
Dim myWB As Workbook, mySht As Worksheet
Dim mySN As String
'�`�N,�ݥ�set myWB
    Set myWB = ThisWorkbook
    mySN = "�s�u�@��"
    
    Set mySht = myWB.Sheets.Add(After:=myWB.Worksheets(myWB.Worksheets.Count))
    mySht.Name = mySN
    
End Sub
Sub �ƻsSHEET���ʨ�Ҧ�SHEET�᭱_��Ʈw()
    Sheets("��l��").copy After:=Sheets(Sheets.Count)
    With ActiveSheet
        .Name = "TOTAL"
        With .Tab
        .ColorIndex = xlColorIndexNone
        End With
    End With
End Sub
Sub �P�_���ɤ��O�_��SHEET�W�s�S�w�W��_��Ʈw()


For Each EBE In Worksheets
    If EBE.Name = "SPEC" Then
        TEMP_B = 1
        Exit For
    End If
Next
 
If TEMP_B = 0 Then
    MsgBox "�Ҷ}�Ҫ��ɮפ��䤣��W��""SPEC""���u�@��A�нT�w�O�_�}���ɮ�!"
    With Workbooks(PM_NAME)
        .RunAutoMacros xlAutoClose
        .Close
    End With
    Exit Sub
End If
End Sub
Sub �j��R���u�@��_��Ʈw()
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("PS_GROUP_UNIT").Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
 Sub �N�u�@��PS_GROUP_UNIT��PS_ORDER�ۦ��ɧR���ñq�L�ɸ��J����u�@��_��Ʈw()
 Application.ScreenUpdating = False
 '���o�{�b���ʤ���SHEET NAME
 ORIGIN_SHEET = ActiveSheet.Name
 '---------------------------------
 '���JSPEC-PM�ɪ�PS_GROUP_UNIT��PS_ORDER�u�@��
 '---------------------------------
 '���oPM���ɦW
 PM_PATH = Application.GetOpenFilename(FileFilter:="Excel�ɮ�(*.xls;*.xlsx),*.xls;*.xlsx", Title:="���JSPEC-PM��..")
 PM_NAME = Right(PM_PATH, Len(PM_PATH) - InStrRev(PM_PATH, "\"))

 If PM_PATH = False Then
    MsgBox "�ާ@�����A�Э��s���榹�\��!"
    Exit Sub
    Else
 '�}��SPEC-PM�ӽƻs���
 '==>�P�_SPEC-PM�ɮ׬O�_�w�g�}��
TEMP_A = 0
    For Each EBE In Workbooks
        If EBE.Name = PM_NAME Then
            TEMP_A = 1
        End If
    Next
 
    If TEMP_A = 0 Then
        Workbooks.Open fileName:=PM_PATH, ReadOnly:=True
    End If
End If
'==>�R�����ɤ����¤u�@��ñq�L�ɸ��J�Ӥu�@��
Windows(ThisWorkbook.Name).Activate
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("PS_GROUP_UNIT").Delete
    Sheets("PS_ORDER").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
Windows(PM_NAME).Activate
    '���b:�}���ɩ�SHEET�W�ٹ��x�s��Ȥ���h���}SUB
    On Error GoTo 20
        If Sheets("PS_GROUP_UNIT").Range("C1").Value <> "PS_UNIT_NAME" Or Sheets("PS_ORDER").Range("B1").Value <> "PS_ORDER" Then
            MsgBox "   ���J���ɮפ����T! �нT�{��A�����榹�\��C            ", vbCritical
            If TEMP_A = 0 Then
                Application.DisplayAlerts = False
                Workbooks(PM_NAME).Close SaveChanges:=False
                Application.DisplayAlerts = True
            End If
            Exit Sub
        End If
    On Error GoTo 0


    Sheets("PS_GROUP_UNIT").copy After:=Workbooks(ThisWorkbook.Name).Sheets("SYSTEM")
Windows(PM_NAME).Activate
    Sheets("PS_ORDER").copy After:=Workbooks(ThisWorkbook.Name).Sheets("PS_GROUP_UNIT")
'�p�G��ӨS�}SPEC-PM,�h����SPEC-PM
 If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(PM_NAME).Close SaveChanges:=False
    Application.DisplayAlerts = True
 End If
 '---------------------------------
 '���PS��ƬO�_�۲�
 '---------------------------------
 Call CHK_IMPORT
'==>�NFOCUS�a�^����e��m
    Windows(ThisWorkbook.Name).Activate
    Sheets(ORIGIN_SHEET).Select
If Sheets("SYSTEM").Range("B1") = "X" Then Exit Sub
MsgBox "   ���JSPEC-PM����!            ", vbInformation
Exit Sub
'���J�����T�ɪ����~�T��
20
MsgBox "   ���J���ɮפ����T! �нT�{��A�����榹�\��C            ", vbCritical
If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(PM_NAME).Close SaveChanges:=False
    Application.DisplayAlerts = True
    Exit Sub
End If
 End Sub

Sub ����SHEET�O�_���ʤ�_��Ʈw()
'vvvvvvvvvvvvvvvvvvvv���bvvvvvvvvvvvvvvvvvvvv

'�ˬd�D��νs�X��O�_����
On Error GoTo 777
    Sheets("�D��").Select
    Sheets("PIPES�s�X��").Select
    Sheets("FITTINGS�s�X��").Select
    Sheets("FLANGES�s�X��").Select
    Sheets("BOLT&NUTS�s�X��").Select
    Sheets("GASKETS�s�X��").Select
    Sheets("VALVES�s�X��").Select
    Sheets("SCH�s�X��").Select
    
On Error GoTo 0

        
'��X���ʪ�ɸ��X���~�T��

'^^^^^^^^^^^^^^^^^^^^���b^^^^^^^^^^^^^^^^^^^^

777 '���}�{����-���ʤu�@��
'==>���Ϳ��~�T��
Title = "���~�T��            "    ' �w�q���D�C
    Msg = vbLf + "������ʤ֥H�U�C����� �@�� �� �h�� �u�@��A���ˬd��A�����榹�{���C   " + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "*�p�G�O�u�@��W�٦��~�Эץ��F�p�G�O�ʤָӤu�@��г]�k�ɤW�C   " + vbLf + vbCrLf + vbLf ' �w�q�T���C
    Msg = Msg + "   [�D��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [PIPES�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [FITTINGS�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [FLANGES�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [BOLT&NUTS�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [GASKETS�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [VALVES�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C
    Msg = Msg + "   [SCH�s�X��]" + vbLf + vbCrLf  ' �w�q�T���C

            MsgBox Msg, vbExclamation, Title
End Sub
Sub �O�@�u�@��()
'�O�@
ActiveSheet.Protect Password:="ABC", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
'�����O�@
ActiveSheet.Unprotect Password:="ABC"
End Sub
Sub �I�s��L�u�@��Worksheet_Change()
'�I�s��L�u�@�� Private Sub Worksheet_Change(ByVal Target As Range)
Dim myWS As Worksheet
Dim myShtCodeName As String 'myShtCodeName:�u�@��CodeName�ݩʭ�
Dim myRg As Range   '�u�@����ܮɡA���b�ק諸�d��
Set myWS = Sheet("�Y�i�u�@��")
Set myRg = myWS.Range("A1")
    myShtCodeName = myWS.CodeName
    Application.Run myShtCodeName & ".Worksheet_Change", myRg
End Sub
Sub ��Ʈw_�ˬd��J�ȬO�_�H�Ϥu�@��W�٭���(ByVal myVal As String)
Dim haveErr As Boolean
Dim iStr1 As String, iStr2 As String
Dim i As Long
Dim myArr(8)
    haveErr = False
    '�s���n��@�u�@��W�١A������u�@������
    '   (1)���i�W�L31�Ӧr (2)���i�������\���r�� :�G \ / ? * [ ] (3)���i�O�ť�(���\space)
    iStr1 = myVal
    iStr2 = ""
    If (iStr1 = "") Then
        iStr2 = "�ܤֶ���J�@�Ӧr"
    ElseIf (Len(iStr1) <= 31) Then
        iStr2 = "���i�W�L31�Ӧr"
    Else
        myArr(1) = ":"
        myArr(2) = "�F"
        myArr(3) = "\"
        myArr(4) = "/"
        myArr(5) = "?"
        myArr(6) = "*"
        myArr(7) = "["
        myArr(8) = "]"
        For i = 1 To UBound(myArr, 1)
            If (InStr(1, iStr1, myArr(i)) <> 0) Then
                iStr2 = "���i�ϥγo�ǲŸ� :�G \ / ? * [ ] "
                Exit For
            End If
        Next i
    End If
    
    If (iStr2 <> "") Then
        msgTitle = "���~            "    ' �w�q���D�C
        msgText = "��J���s�����H�U���~�A�Эק�" + vbLf    ' �w�q�T���C
        msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
        msgText = msgText + iStr2 + vbLf  ' �w�q�T��
        msgStyle = vbCritical '���"X"�Ϯ�
        MsgBox msgText, msgStyle, msgTitle
    End If

End Sub
