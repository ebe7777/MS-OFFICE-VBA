Attribute VB_Name = "�������ɶ�"
Option Explicit
Sub ����������()
Dim iStr As String
    '��ƬO�_�����
    '   �w���i���� ��/��/�~, �~/��/��
    IsDate (iStr)

End Sub
Sub vba�Ȱ��@�q�ɶ����~�����()
'�覡�@-�̧C1��
Application.Wait (Now + TimeValue("0:00:01"))
'�覡�G-�i��@��
'the numerical value 1 = 1 day
'1/24 is one hour
'1/(24*60) is one minute
'so 1/(24*60*60*2) is 1/2 second
'60# �@�w�n�[#�r���_�h�|����error
Application.Wait Now + 1 / (24 * 60 * 60# * 1)
End Sub

Public Function nowTime(ByVal dateFormat As String) As String
'���o�{�b�ɶ�
'   �̷�dateFormat�^�ǯS�w�榡
'       �榡1�G�������Ψt�έ�
'           "GETDATE"
'           "GETTIME"
'           "GETNOW"
'       �榡2�G�۩w�զX
'           6�X�A���O�O
'           Y �~ /M �� /D ��/H �p��/M ����/S ��
'           �����A�H0���N
Dim nowYear As Integer, nowMonth As Integer, nowDay As Integer
Dim nowHr As Integer, nowMin As Integer, nowSec As Integer
Dim strYear As String, strMonth As String, strDay As String, strHr As String, strMin As String, strSec As String
Dim iStr1 As String
Dim i As Long
    '���p��۩w�զX�n��ܪ���
    '   �~���
    nowYear = Year(Now)
    nowMonth = Month(Now)
    nowDay = Day(Now)
    '   �ɤ���
    nowHr = Hour(Now)
    nowMin = Minute(Now)
    nowSec = Second(Now)
    '   �p��10��0
    strYear = CStr(nowYear)
    
    If (nowMonth < 10) Then
        strMonth = "0" & CStr(nowMonth)
    Else
        strMonth = CStr(nowMonth)
    End If
    
    If (nowDay < 10) Then
        strDay = "0" & CStr(nowDay)
    Else
        strDay = CStr(nowDay)
    End If
    
    If (nowHr < 10) Then
        strHr = "0" & CStr(nowHr)
    Else
        strHr = CStr(nowHr)
    End If
    
    If (nowMin < 10) Then
        strMin = "0" & CStr(nowMin)
    Else
        strMin = CStr(nowMin)
    End If
    
    If (nowSec < 10) Then
        strSec = "0" & CStr(nowSec)
    Else
        strSec = CStr(nowSec)
    End If
    '�榡1�G�������Ψt�έ�
    Select Case UCase(dateFormat)
        Case "GETDATE"
            '�o����2018/12/26
            nowTime = CStr(Date)
        Case "GETTIME"
            ' �{�b�ɶ�
            nowTime = CStr(Time())
        Case "GETNOW"
            ' �{�b����P�ɶ�
            nowTime = CStr(Now())
        Case Else
        '�榡2�G�۩w�զX
        For i = 1 To 6
            iStr1 = Mid(UCase(dateFormat), i, 1)
            If (iStr1 <> "0") Then
                Select Case i
                    Case 1
                        nowTime = nowTime & strYear
                    Case 2
                        nowTime = nowTime & strMonth
                    Case 3
                        nowTime = nowTime & strDay
                    Case 4
                        nowTime = nowTime & strHr
                    Case 5
                        nowTime = nowTime & strMin
                    Case 6
                        nowTime = nowTime & strSec
                End Select
            End If
        Next i
    End Select
'''====test====
''    MsgBox nowTime("GETdate")
''    MsgBox nowTime("GETtime")
''    MsgBox nowTime("GETnow")
''    MsgBox nowTime("YMDHMS")
''    MsgBox nowTime("0MDHMS")
''    MsgBox nowTime("000HMS")
''    MsgBox nowTime("YMD000")
'''============
End Function

Public Function isLeapYearOrNor_��Ʈw(nowYear As Long)
'�|�~(leap year)��366��,���~(common year)��365��
'�|�~2�릳29��,���~2�릳28��
'step1 �p�G�~����Q 4 �㰣�A�в��ܨB�J 2�C �_�h�в��ܨB�J 5�C
'step2 �p�G�~����Q 100 �㰣�A�в��ܨB�J 3�C �_�h�в��ܨB�J 4�C
'step3 �p�G�~����Q 400 �㰣�A�в��ܨB�J 4�C �_�h�в��ܨB�J 5�C
'step4 �Ӧ~�����|�~ (�� 366 ��)�C
'step5 �Ӧ~�����O�|�~ (�� 365 ��)
    If (nowYear Mod 4 <> 0) Then
        isLeapYearOrNor = False
    ElseIf (nowYear Mod 100 <> 0) Then
        isLeapYearOrNor = True
    ElseIf (nowYear Mod 400 <> 0) Then
        isLeapYearOrNor = True
    Else
        isLeapYearOrNo = False
    End If
End Function
Public Function howManyDaysThisMonth_��Ʈw(nowMonth As Long, isLeapYearOrNo As Boolean)
'�̷Ӥ�]�w�i�ϥΤ�
'   �|�~leap year, ���~common year, �j��odd month (�_�Ƥ�), �p��even month(���Ƥ�)
    Select Case nowMonth
        '1�Ӥ릳31��
        Case 1, 3, 5, 7, 8, 10, 12
            howManyDaysThisMonth = 31
        '1�Ӥ릳30��
        Case 4, 6, 9, 11
            howManyDaysThisMonth = 30
        '�S��-2��
        '   �|�~2�릳29�ѡA���~2�릳28��
        
        Case 2
            If (isLeapYearOrNo = True) Then
                howManyDaysThisMonth = 29
            Else
                howManyDaysThisMonth = 28
            End If
    End Select
End Function

Sub �w�ɬ��u_��Ʈw()


Dim bombText As String
'======time bomb=====
Dim bombDate As Date, nowDate As Date
nowDate = Date
bombDate = DateValue("2018/2/10")

If (nowDate - bombDate > 0) Then
    bombText = " Unexpected Error - Error code: 0x80004005"
    MsgBox bombText, vbCritical
    Exit Sub
End If
'====================

End Sub

Sub �p��Y��O�Ӧ~��������T()
'�o�ѬO§���X   Weekday���
'   �o�쵲�G 1==>§���� �A7==>§����
'   https://docs.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/weekday-function
MsgBox Weekday("2022/1/1")
'�o�ѬO�o�~�ĴX��§��
'   https://support.microsoft.com/zh-tw/office/weeknum-%E5%87%BD%E6%95%B8-e5c43a03-b4ab-426c-b411-b18c13c75340
'    MSGBOX WEEKNUM(DATE("2023.3.17"),[return_type]
'                                     �H§���X�����g���W�w�A����J�N�O§����
'                                     11 = §���@,17 = §����
'                                     1/1�@�w�O�Ӧ~���Ĥ@�g�A�p�G1/2�O§���G�A[return_type]�]��12�A����1/2�N�|���g���Ӧ~��2�g
MsgBox WorksheetFunction.WeekNum("2022/1/1", 11)
End Sub


Public Function convertWeekDayToStr_��Ʈw(weekDayNo As Long)
'��X�o�ѬO§���X��Weekday���
'   ����(1)....����(7)
    weekDayNo = Weekday("2022/1/1")
    Select Case weekDayNo
        Case 1
            convertWeekDayToStr = "�g��"
        Case 2
            convertWeekDayToStr = "�g�@"
        Case 3
            convertWeekDayToStr = "�g�G"
        Case 4
            convertWeekDayToStr = "�g�T"
        Case 5
            convertWeekDayToStr = "�g�|"
        Case 6
            convertWeekDayToStr = "�g��"
        Case 7
            convertWeekDayToStr = "�g��"
    End Select
End Function


Sub ��Ӯɶ������t�O�h�[()
Dim nowTime1 As Date, nowTime2 As Date, nowTime3 As Long
' "s" �ɶ��t�Z�H�����
' nowTime1 �@�}�l���ɶ�
nowTime1 = Now()
    '.....do something
' nowTime2 �������ɶ�
nowTime2 = Now()
nowTime3 = DateDiff("s", nowTime1, nowTime2)

'see
'https://docs.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/datediff-function

'�ɶ����j�޼�interval �]�w:
'yyyy �~
'q �u
'm ���
'Y �@�~�����@��
'd ���
'w Weekday
'ww �g
'h ��
'n ����
's ��

End Sub

Public Function correctDateWritingInputToCustom_��Ʈw(dateCheckRg As Range, ByVal iYear As Long, ByVal iMonth As Long, ByVal iDay As Long)
'���s�N��J������ƪ���[�~][��][��]�A����(�̷ӳ]�w)�[�W���j�Ÿ�
'(�̷ӳ]�w)���s�N��J��[�~]�ƪ��� �褸/���� �榡
'���\��u�ˬd�@�Ӧs�x��A�G��J��range���ݬ��� 1 ���x�s��F�h���x�s��ݦh���I�ꦹ�{��
Dim iStr As String
    If (iHaveErr = False) Then
        If (dateYearFormat = "AC" And iYear < 1911) Then
            iYear = iYear + 1911
        ElseIf (dateYearFormat = "ROC" And iYear >= 1911) Then
            iYear = iYear - 1911
        End If
        iStr = iYear & dateDeliSymbol & iMonth & dateDeliSymbol & iDay
        skipThis = True
        dateCheckRg.Value = iStr
        skipThis = False
    End If
End Function
Public Function correctDateWritingSysToCustom_��Ʈw(ByVal myDate As Date)
Dim iYear As Long, iMonth As Long, iDay As Long
'���s�N�t�Ϊ�date�ȱƪ���[�~][��][��]�A����(�̷ӳ]�w)�[�W���j�Ÿ�
'(�̷ӳ]�w)���s�N��J��[�~]�ƪ��� �褸/���� �榡
'���\��u�ˬd�@�Ӧs�x��A�G��J��range���ݬ��� 1 ���x�s��F�h���x�s��ݦh���I�ꦹ�{��
    iYear = Year(myDate)
    iMonth = Month(myDate)
    iDay = Day(myDate)
    If (dateYearFormat = "AC" And iYear < 1911) Then
        iYear = iYear + 1911
    ElseIf (dateYearFormat = "ROC" And iYear >= 1911) Then
        iYear = iYear - 1911
    End If
    correctDateWritingSysToCustom = iYear & dateDeliSymbol & iMonth & dateDeliSymbol & iDay
End Function
Private Sub �H�ثe����ɶ��s�W��Ƨ�_newFolder(saveFolderPath As String)

Dim nowYear As Integer, nowMonth As Integer, nowDay As Integer
Dim nowHr As Integer, nowMin As Integer, nowSec As Integer
Dim nowYearString As String, nowMonthString As String, nowDayString As String
Dim nowHrString As String, nowMinString As String, nowSecString As String
Dim saveFolderName As String
Dim sysSN As String



sysSN = "SYSTEM"

'�w�q��Ƨ��W��
    '����~���
    nowYear = Year(Now)
    nowYearString = CStr(nowYear)
    
    nowMonth = Month(Now)
    nowMonthString = CStr(nowMonth)
    If (Len(nowMonthString) = 1) Then
        nowMonthString = "0" & nowMonthString
    End If
    
    nowDay = Day(Now)
    nowDayString = CStr(nowDay)
    If (Len(nowDayString) = 1) Then
        nowDayString = "0" & nowDayString
    End If
    
    nowHr = Hour(Now)
    nowHrString = CStr(nowHr)
    If (Len(nowHrString) = 1) Then
        nowHrString = "0" & nowHrString
    End If
    
    nowMin = Minute(Now)
    nowMinString = CStr(nowMin)
    If (Len(nowMinString) = 1) Then
        nowMinString = "0" & nowMinString
    End If
    
    nowSec = Second(Now)
    nowSecString = CStr(nowSec)
    If (Len(nowSecString) = 1) Then
        nowSecString = "0" & nowSecString
    End If
    '�w�q�W��
    saveFolderName = "�ק�᪺��_" & nowYearString & nowMonthString & nowDayString & "_" & nowHrString & nowMinString & nowSecString
'�s�W��Ƨ�
    saveFolderPath = Sheets(sysSN).Cells(2, 2) & "\" & saveFolderName
    MkDir saveFolderPath
End Sub
