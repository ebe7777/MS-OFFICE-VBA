Attribute VB_Name = "A_����[�tvba�B��"
Sub �[�ְ���t�ת��覡()
'�קK�b�ù��W��ܸ���ܰ�
'   ����ù���s
Application.ScreenUpdating = False
Application.ScreenUpdating = True
'   �����x�s��B��
Application.Calculation = xlCalculationManual
Application.Calculation = xlCalculationAutomatic
'   �קK�ϥ�Application.StatusBar

'�קK���Ʃʪ��M�u�@����
'   �N��Ƽg�i�ܼ�(�}�C)�A���p��B�z
    i = Range("A1").Value
    For ii = 1 To 1000000
        i = i + ii
    Next ii
'   Ū��/�g�J�x�s���Ʈɨϥΰ}�C�@���ʰ���
Dim myArray() As Variant
    myArray = Sheets("test").Range("A3:B4").Value
    '�i�b�}�C���s����
    myArray(1, 1) = 31
    Sheets("test").Range("A6:B7").Value = myArray

'�ۤv�g�{�����NApplication.WorksheetFunction

'�קK�ϥΩw�q��Variants���ܼ�

'�קK�P�_ ��r ���
'   �|�ҡGif (myText = "abc") then
'       select case myText : case  "abc"
'   �i�N��r�ର�Ʀr�ӧP�_�AĴ�pEnum
Public Enum enumGender
    Male = 0
End Enum
Dim Gender As enumGender
    Select Case Gender
        Case Male
    End Select

'�קK���(.select)�u�@���A���x�s�檺�ȡF�����������x�s�檺��
    myValue = Worksheets("sheet1").Cells(1, 1).Value
    
'�קK���ư���ƾǹB��
    For i = 1 To 100
        '�i�H�N(3 * 10 / 12)��X����n
        'myValue = myValue + (3 * 10 / 12)
        a = (3 * 10 / 12)
        myValue = myValue + a
    Next i
    
'���Ψϥη|�s���x�s��J����k�A�p.copy .paste .ClearContents
'   ���ާ@�x�s�檺�ݩʡA�p .value = XXX
End Sub
