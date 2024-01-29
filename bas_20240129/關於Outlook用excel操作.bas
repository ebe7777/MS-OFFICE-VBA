Attribute VB_Name = "����Outlook��excel�ާ@"
'�򥻪���
'   https://docs.microsoft.com/zh-tw/office/vba/api/outlook.mailitem
'   https://support.microsoft.com/zh-tw/topic/%E4%BD%BF%E7%94%A8-word-%E6%96%87%E4%BB%B6%E5%92%8C-excel-%E6%B4%BB%E9%A0%81%E7%B0%BF%E4%B8%AD%E4%B9%8B%E8%B3%87%E6%96%99%E5%BE%9E-outlook-%E5%82%B3%E9%80%81%E9%83%B5%E4%BB%B6%E7%9A%84-vba-%E5%B7%A8%E9%9B%86-56bbd7a9-7814-9c52-2c83-e92c01fa8418
'����outlook����
'   https://learn.microsoft.com/zh-tw/office/vba/api/outlook.mailitem.close(method)
'�NWord�����e����K�W
'   https://stackoverflow.com/questions/35609112/how-to-send-a-word-document-as-body-of-an-email-with-vba
Sub ��Ʈw_autoEmail()
'20210630  �ݱNoutlook�����{���X�@���O
 
Dim objOutlook As Object, mailSendItem As Object
Dim objWord As Object, objWordDoc As Object
Dim contentTxt As String, contentFilePath As String
Dim totalRows As Long
Dim mailInfoArray()
Dim mySignature As String
Dim mailStartRow As Long, mailEndRow As Long
Dim needAttachmentOrNot As Boolean
Dim excuteTime As String
Dim i As Long, iCount1 As Long
    Call loadPublicVar
    If (loadPubicVarHaveErr = True) Then
        GoTo 999
    End If
    
    '�i���ϥΪ̳W�h�A�ݬO�_�~��
    '   �m�W�B�s���q�W��U�����ǻݻP�u�@��[����H]�@�P
    '   �i����������r��10��
    msgTitle = "�`�N            "    ' �w�q���D�C
    msgText = "�ϥΦ��\��ɡA�u�@�� [" & mailAddressSN & "] �C�@�C�� �m�W/�s��(A~B��)" + vbLf + vbCrLf     ' �w�q�T���C
    msgText = msgText + "���ݻP�u�@�� [" & mailTitleContentSN & "] �M�u�@�� [" & mailAttach1SN & "] �����e�@�P" + vbCrLf
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�p�T�w���~����I��[�O]�A�Ϋ�[�_]���}�{��  " + vbCrLf + vbCrLf  ' �w�q�T���C
    msgText = msgText + "�`�N!�Ҧ���ƪ� �m�W/�s��/email(A~C��)������аO���|�Q����"
    answer = MsgBox(msgText, vbYesNo + vbQuestion, msgTitle)
    If answer = vbNo Then
        GoTo 999
    End If
    excuteTime = Now()
    totalRows = myDataRows(ThisWorkbook.Name, mailAddressSN, "A", 65536)
    
    '���b-�S����J�H���}
    If (totalRows <= titleUsedRows) Then
        MsgBox "�u�@�� [" & mailAddressSN & "] �̧䤣���� (�`�N�A�C�@�CA~C�泣����g�A�_�h�|�y���{���~�P)"
        GoTo 999
    End If

    '���b�ˬd
    '   [����H]���m�W�B�s���q�W��U�����ǡA�M�u�@��[�l��D���P����]�̭����ȬO�_�P�u�@��@�P
    Call compareShts(mailAddressSht, mailTitleContentSht)
    If (stopRun = True) Then
        GoTo 999
    End If
    '   [����H]���m�W�B�s���q�W��U�����ǡA�M�u�@��[����1��T]�̭����ȬO�_�P�u�@��@�P
    Call compareShts(mailAddressSht, mailAddressSht)
    If (stopRun = True) Then
        GoTo 999
    End If
    '   ���b-email�a�}�O���~��
    For i = (titleUsedRows + 1) To totalRows
        stopRun = False
        With mailAddressSht
            If (IsError(.Cells(i, 3)) = True) Then
                '�����~�Хܶ���
                With .Cells(i, 3).Interior
                    .Pattern = xlSolid
                    .Color = 65535
                    stopRun = True
                End With
            Else
                '�S���~�R������
                With .Cells(i, 3).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End With
    Next i
    If (stopRun = True) Then
        msgTitle = "�����~            "    ' �w�q���D�C
        msgText = "�u�@�� [" & mailAddressSN & "] ��email��Ʀ����D" + vbLf + vbCrLf  ' �w�q�T���C
        msgText = msgText + "�w�b����H�������Х�" + vbLf + vbCrLf  ' �w�q�T��
        msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
        msgText = msgText + "�Эץ���A�����榹�{��" + vbLf   ' �w�q�T���T
        msgStyle = vbCritical '���"X"�Ϯ�
        MsgBox msgText, msgStyle, msgTitle
        GoTo 999
    End If
    
    '���ϥΪ̿�ܿ�X�d��
    UserForm1.Show
    '   ���b-�ϥΪ̫�X���}�{��
    If (stopRun = True) Then
        GoTo 999
    End If
    mailStartRow = sysSht.Range("B3").Value
    mailEndRow = sysSht.Range("D3").Value
    needAttachmentOrNot = sysSht.Range("B4").Value
    
    '�`���l�������T
    '   �N�C��email�һݪ���T���g�b�}�C��
    ReDim mailInfoArray(totalRows, 4)
    '[n]�ĴX�Ӧ����
    '[x][0]�m�W_�s��(�H�u�ˬd�ΡA�{����Ū��) [x][1]email�a�} [x][2]email�D�� [x][3]�������ɦW [x][4]����1�ɦW
    For i = mailStartRow To mailEndRow
        mailInfoArray(i, 0) = mailAddressSht.Cells(i, 1) & "_" & mailAddressSht.Cells(i, 2)
        mailInfoArray(i, 1) = mailAddressSht.Cells(i, 3)
        mailInfoArray(i, 2) = mailTitleContentSht.Cells(i, 4)
        mailInfoArray(i, 3) = sysSht.Range("D1") & "\" & mailTitleContentSht.Cells(i, 3) & ".docx"
        If (needAttachmentOrNot = True) Then
            mailInfoArray(i, 4) = sysSht.Range("D2") & "\" & mailAttach1Sht.Cells(i, 3) & ".docx"
        End If
    Next i
    
'ppppp�}�Y-�I�s����
ThisWorkbook.Activate
progressBarPercentNo = 0
ProgressBar.Show False
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

    iCount1 = 0
    needAttachmentOrNot = sysSht.Range("B4").Value
    For i = mailStartRow To mailEndRow

        contentFilePath = mailInfoArray(i, 3)
        '���X�S�wWord�ɤW���r�A�Hword��l���榡
        On Error GoTo 880
            Set objWord = CreateObject("Word.Application")
            Set objWordDoc = objWord.Documents.Open(fileName:=contentFilePath, ReadOnly:=True)
            objWordDoc.Content.copy
            objWordDoc.Close
            Set objWord = Nothing
        On Error GoTo 0
        On Error GoTo 881
            '�H�H
            Set objOutlook = CreateObject("Outlook.Application")
            Set mailSendItem = objOutlook.CreateItem(0)
            '   ���F�ϥ�outlook���w�]ñ�W�A����ܵ����ýƻs���e
            mailSendItem.Display
        On Error GoTo 0
        On Error GoTo 882

            With mailSendItem
                '�D�����ťյ������~
                If (mailInfoArray(i, 2) = "") Then
                    GoTo 883
                Else
                    .Subject = mailInfoArray(i, 2)
                End If
                '   late binding�u���Jbodyformat���N���X
                .bodyformat = 2
                '1:olFormatPlain
                '2:olFormatHTML
                '   ����|�Houtlook�覡���a
                '3:olFormatRichText
                '   ����|�H��r�����J���󪺤覡���a
                Set editor = .GetInspector.WordEditor
                editor.Content.Paste
                If (needAttachmentOrNot = True) Then
                    On Error GoTo 884
                        .Attachments.Add mailInfoArray(i, 4)
                    On Error GoTo 0
                End If
                .To = mailInfoArray(i, 1)
                On Error GoTo 885
                    .Send
                    iCount1 = iCount1 + 1
                On Error GoTo 0
            End With
            '�w�o�Xmail���A�bemail�ж����AD��g�W�H�H�ɶ�
            mailAddressSht.Range("C" & i).Interior.Color = 65535
            mailAddressSht.Range("D" & i).Value = excuteTime
        On Error GoTo 0

'����ҼW�[�@�w�ƶq���i��
'ppppp-0 to 100
ThisWorkbook.Activate
'   ���q���浲���@�|�W�[iProcessRange�Ӷi�צʤ���/��Ƶ��Ʀ@iDataCounts��/�CiProcessRangeGap���q�s�p��@���i��
iProcessRange = 100
iDataCounts = totalRows - mailStartRow + 1
iProcessRangeGap = 1
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

    Next i
    
    
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    GoTo 991
    
880
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "�u�@�� [" & mailAddressSN & "] ��" & i & "�C������̪� ������ �䤣��" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�i���]�p�U" + vbLf  ' �w�q�T��
    msgText = msgText + "(1)�����ɩ|������ --> �а���{��[���ͤ���] " + vbLf  ' �w�q�T��
    msgText = msgText + "(2)�ɮ׸��|�����T --> �а���{��[�]�w] " + vbLf  ' �w�q�T��
    msgText = msgText + "(3)�u�@�� [" & mailTitleContentSN & "] C�檺�ȩM�������ɦW���� --> �Эק��ɦW�Τu�@���e " + vbLf + vbCrLf ' �w�q�T��
    msgText = msgText + "�ק粒����A�бq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
881
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "Outlook�{���|���}�ҡA�Φ]������]����" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�Ф�ʶ}��Outlook�{���A�ñq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
882
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "�B�z��@�b�o�ͥ��������D" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�Х����ձq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
883
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "�u�@�� [" & mailTitleContentSN & "] ��" & i & "�C������̪� �l��D�� �O�ť�" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�Эק�u�@�� [" & mailTitleContentSN & "] D��A�ñq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
884
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "�u�@�� [" & mailAddressSN & "] ��" & i & "�C������̪� ������ �䤣��" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�i���]�p�U" + vbLf  ' �w�q�T��
    msgText = msgText + "(1)�����ɩ|������" + vbLf  ' �w�q�T��
    msgText = msgText + "    -> �а���{��[���ͪ���] " + vbLf  ' �w�q�T��
    msgText = msgText + "(2)�ɮ׸��|�����T" + vbLf  ' �w�q�T��
    msgText = msgText + "    -> �а���{��[�]�w] " + vbLf  ' �w�q�T��
    msgText = msgText + "(3)�u�@�� [" & mailAddressSN & "] C�檺�ȩM�������ɦW����" + vbLf ' �w�q�T��
    msgText = msgText + "    -> �Эק��ɦW�Τu�@���e " + vbLf + vbCrLf ' �w�q�T��
    msgText = msgText + "�ק粒����A�бq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
885
'ppppp����
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    '����outlook�����B���s��
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "�o�{���D           "    ' �w�q���D�C
    msgText = "�u�@�� [" & mailAddressSN & "] ��" & i & "�C������̶l��a�}�����T" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�Э��s���ͪ����ɩέ��s���w�����ɸ��|�A�ñq���_�B�~�����" + vbLf  ' �w�q�T��
    msgText = msgText + "�p�@���d�b�P�@�ӤH(���ХX�{���T��)�A���p���{���}�o��"    ' �w�q�T��
    msgStyle = vbExclamation '���"!"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    


991
    msgTitle = "�T��          "    ' �w�q���D�C
    msgText = "�B�z�����A�`�@�H�X " & iCount1 & " �ʫH" + vbLf ' �w�q�T���C
    msgText = msgText + "(�w�N���\�o�Xmail���b�u�@�� [" & mailAddressSN & "] ��C��H�������Х�)" ' �w�q�T���C
    msgStyle = vbInformation '���"i"�Ϯ�
    MsgBox msgText, msgStyle, msgTitle
    GoTo 999
999
    ThisWorkbook.Activate
    mailAddressSht.Activate
End Sub
