VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "�{�����椤..."
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   5025
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=========
'�}�o��     brucechen1@micb2b.com
'�}�o���   2020-03-11
'�ק���   2023-12-19
'=========

Private Sub UserForm_Initialize()
'���]�wpublic variable
'   progressBarPercentNo as long

'Label1��Width��ܶi�ױ� 0~240
'       caption��ܶi�׭�,�Hstring�Φ�

Dim labelWidth As Long
    Label1.caption = "0%"
    Label1.Width = 0
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Sub updateProgressBar(ByVal percentNo As Double)
    percentNo = Round(percentNo, 1)
    ProgressBar.Label1.caption = percentNo & "%"
    ProgressBar.Label1.Width = Round(240 * (percentNo / 100), 1)
    DoEvents
End Sub

Sub �ϥήɦb�{����J�H�U()

'ppppp�}�Y-�I�s����
ThisWorkbook.Activate
progressBarPercentNo = 0
ProgressBar.Show False
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'ppppp����-�T�w�ȼg�k
ThisWorkbook.Activate
progressBarPercentNo = progressBarPercentNo + 5
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'ppppp����-�Hi�ܤƼg�k
ThisWorkbook.Activate
progressBarPercentNo = progressBarPercentNo + ((1 / maxOFi) * baseValue)
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'����ҼW�[�@�w�ƶq���i��
'ppppp-5 to 25
ThisWorkbook.Activate
'   ���q���浲���@�|�W�[iProcessRange�Ӷi�צʤ���/��Ƶ��Ʀ@iDataCounts��/�CiProcessRangeGap���q�s�p��@���i��
iProcessRange = 20
iDataCounts = 50000
iProcessRangeGap = 500
'   �p��O�_�n���s�p��i�סA�ثe��ƬO��i��
iMod = i Mod iProcessRangeGap
'   �p��@�n��s�X��
iMax = CInt(iDataCounts / iProcessRangeGap)
'   �n��s�i�׮ɡA�i�ױ��W�[ (1/iMax*iProcessRange) ���i��
If (iMod = 0) Then
    progressBarPercentNo = progressBarPercentNo + ((1 / iMax) * iProcessRange)
    Call ProgressBar.updateProgressBar(progressBarPercentNo)
End If
'ppppp

'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
End Sub


Sub �H�U�d�s�Ѧ�()
'Sub xxx()
    '
    '
    ''<<<<<<<<<<<<<<<<�g�b�����B���>>>>>>>>>>>>>>>>>>>>>>>
    '
    'STIME = Time 'START TIME�_�Ϯɶ�
    'ProgressBar.Show 0
    '
    '
    ''@@@@@@�i�ױ�����@@@@@@'
    'NTIME = Time 'NOW TIME�ثe�ɶ�
    'Call Bar_GO(STEP_NOW, ALL_CACU_ROWS, STIME, NTIME)
    'DoEvents
    ''@@@@@@@@@@@@@@@@@@@@@@
    '
    '
    'Unload ProgressBar '�Ψ������i�ת�
'End Sub
'
'
''<<<<<<<<<<<<<<<<<�i�ױ�function>>>>>>>>>>>>>>>>>
'
    'Function ProgressBar_GO(ByVal STEP, TOTAL, START_TIME, NOW_TIME) '�ܼ�Step���ثe���檺�B�J�A�ܼ�Total���`�B�J...
    '
    ''�i�ױ�FUNCTION
    '
    ''�b���n�гy�@�Ӫ��,�b�ݩʸ̭ק��W��ProgressBar;�ԥXLA�ק��WSC
    'ProgressBar.SCH.Caption = Round(STEP / TOTAL * 100, 0) & "%"
    'ProgressBar.SCH.Width = Round(STEP / TOTAL * 240, 0)
    'ProgressBar.Caption = "�w���Ѿl�ɶ��G" _
    '    & Minute((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " �� " & _
    '    Second((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " ��   "
    ''ProgressBar.Caption = "����� [" & ActiveSheet.Name & "] �ثe���Ͷi�סG" & Format(Round(STEP / TOTAL * 100, 2), "0") & "% �F �w���Ѿl�ɶ��G" _
    ''    & Minute((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " �� " & _
    ''    Second((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " ��   "
    'DoEvents '�Ψ���ܶi�צʤ���
    'End Function
    ''^^^^^���ͤu�@�ɼƳ��� �Ҧ�SUB�w�ק�ŦX1.0��^^^^^
'
'End Function
End Sub
