Attribute VB_Name = "����K�X�}��"
Sub �u�@��K�X�}��1_��Ʈw()

answer = MsgBox("�Y�N�i��}�ѡA�O�_�~��H", vbOKCancel, "�޿�}�Ѫk")

If answer = 2 Then
    MsgBox "�����}��"
    Exit Sub
End If

StartTime = Time '�̫���ܰ���ɶ���

Dim I1 As Integer, I2 As Integer, I3 As Integer, I4 As Integer, I5 As Integer, I6 As Integer
Dim I7 As Integer, I8 As Integer, I9 As Integer, I10 As Integer, I11 As Integer, I12 As Integer

On Error Resume Next
    
ProgressBar.Show 0
ProgressBar.ProgressBar_1.Min = 0
ProgressBar.ProgressBar_1.Max = 100

For I1 = 65 To 66: For I2 = 65 To 66: For I3 = 65 To 66: For I4 = 65 To 66: For I5 = 65 To 66: For I6 = 65 To 66
For I7 = 65 To 66: For I8 = 65 To 66: For I9 = 65 To 66: For I10 = 65 To 66: For I11 = 65 To 66: For I12 = 32 To 126

    ActiveSheet.Unprotect Chr(I1) & Chr(I2) & Chr(I3) & Chr(I4) & Chr(I5) & Chr(I6) & Chr(I7) & Chr(I8) & Chr(I9) & Chr(I10) & Chr(I11) & Chr(I12)
    
    ProCount = ProCount + 1
   
    If ActiveSheet.ProtectContents = False Then
        
        If ProCount < 1000 Then
            ProgressBar.Label.Caption = "�����}�Ѧ@���դF " & ProCount & " �ձK�X�C"
        Else
            ProgressBar.Label.Caption = "�����}�Ѧ@���դF " & Format(ProCount, "0,000") & " �ձK�X�C"
        End If
        
        FinalTime = Time '�̫���ܰ���ɶ���
        InputBox "�w�����O�@�A�ϥΤ��}�ѽX�p�U(���� " & Minute(FinalTime - StartTime) & " �� " & Second(FinalTime - StartTime) & " ��)�G" & vbNewLine & vbNewLine & "�`�N�I���}�ѽX�D�ϥΪ̤���l�K�X�A����̤��K�X�޿�{�w�ۦP�C�@���H�}�ѽX�i��ϦV�[�K�A��ϥΪ̥�i��J��l�K�X�Ө����O�@�C", "�޿�}�Ѫk", Chr(I1) & Chr(I2) & Chr(I3) & Chr(I4) & Chr(I5) & Chr(I6) & Chr(I7) & Chr(I8) & Chr(I9) & Chr(I10) & Chr(I11) & Chr(I12)
        Unload ProgressBar '�Ψ������i�ת�
        Exit Sub
    End If
    
    Call Bar(ProCount, 194560)

Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next

End Sub
Function Bar(ByVal STEP, TOTAL) '�ܼ�Step���ثe���檺�B�J�A�ܼ�Row���`�B�J"

ProgressBar.ProgressBar_1.Value = Round(STEP / TOTAL * 100, 0)

If (STEP / TOTAL * 100) > 20 Then
    ProgressBar.Label.Caption = "�W�j���K�X�I�����O�h�֩O�H"
ElseIf (STEP / TOTAL * 100) > 10 Then
    ProgressBar.Label.Caption = "������A�O�Ӧn�K�X�C"
ElseIf (STEP / TOTAL * 100) > 0 Then
    ProgressBar.Label.Caption = "�ոլݯण�༵�L10%�H"
End If

ProgressBar.MessageBox_1.Value = "�ثe�i�סG" & Format(Round(STEP / TOTAL * 100, 2), "0.00") & "%"
DoEvents '�Ψ���ܶi�צʤ���

End Function




Sub �u�@��K�X�}��2_��Ʈw()
Dim mySht As Variant
For Each mySht In ActiveWorkbook.Worksheets
    mySht.Protect DrawingObjects:=True, CONTENTS:=True, AllowFiltering:=True
    
    mySht.Protect DrawingObjects:=False, CONTENTS:=True, AllowFiltering:=True
    
    mySht.Unprotect
Next
MsgBox "�Ҧ��u�@�� [�O�@�u�@��] �w�Ѱ�"

End Sub

Sub �u�@��K�X�}��3_��Ʈw()
'�KVBA�覡
'   �N�ɦW�令.zip
'   ���n�����Y�A�����H���Y�n��}��
'       *�H�U�H7zip����
'   �i�J��Ƨ� xl
'   �i�J��Ƨ� worksheets
'   �s����C��sheet��xml (�I���xml > ���ƹ��k�� >�s��)
'       ��� "<sheetProtection" �}�Y���q���A�ñq  "<" ��Ӭq "/>" ������q�R��
'       �C���ɮק粒��n�x�s�A�{���|�����A "�O�_�n�b���Y�Ҥ���s�ɮ�?" ��� "�T�w"
'   �Nzip�ɧ�^��l���ɦW
End Sub

Sub ����ï�K�X�}��1_��Ʈw()
'�KVBA�覡
'   �N�ɦW�令.zip
'   �}�� (�ɦW)\xl\ > ���workbook.xml
'   �s����� > ��� <workbookProtection .... /> �ñN���R��
'   �Nzip�ɧ�^��l���ɦW
End Sub
