Attribute VB_Name = "������e�C���r���ؽu"
Sub ��z��e�C���r��_��Ʈw()
Sheets("BASE_FORM").Select
'��z��e
Columns("A:AY").ColumnWidth = 2.13
'��z���Y�C�C��
Rows("1:1").RowHeight = 19.5
Rows("2:11").RowHeight = 16.5
'�ק�r��
Cells.Font.Name = "Consolas"
End Sub




Function �Ҷq����ӼW�h���r��_��Ʈw(ByVal DATA_ADD, ORIG_WORD_NUM, CHANGE_ROW_WORD_NUM, ADD_WORD_NUM)
'(ByVal �ԭz���,��l�ԭz�r��,����r��(�֭p,�Ĥ@��44,�h�ĤG��88),�W�[�r��)

'�]�w�W�[�r�� = �W�[�r�� �Ӥ�����0,�O�Ω�P�_�ĤG��,�ĤT�洫��ɼW�[�r�ƭn�֥[�e�@�檺�W�[��;�G�ޥΦ�FUNCTION���{���A�i�J�ĤG�^�餧�e�n�N
'(���W)�̫᪺�p�⵲�G (�Ҷq����ӼW�h���r��_��Ʈw)�g�J(ADD_WORD_NUM)��
'ADD_WORD_NUM = ADD_WORD_NUM
'�p�G��Ӫ��r��+�W�[���r�Ʒ|�j�󴫦�r��(�N���ٻݭn�P�_�O�_���U�@��),�~�ݭp��
If ORIG_WORD_NUM + ADD_WORD_NUM > CHANGE_ROW_WORD_NUM Then
    '�p�G����r�ƸӦr���Ů�" "�Τ��u"-",�S�Ϊ̸Ӵ���r���U�@�Ӧr�O�Ů�,�t�η|�۰ʴ��椣�|�h�[�r��;�G���O�~�ݭp��
    If Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> " " And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> "-" And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM + 1, 1) <> " " Then
        '�P�_�����I�e�O����" "�٬O"-"�ӧP�_�ӥέ��ӨӧP�_�r��
        If InStrRev(DATA_ADD, " ", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM) < InStrRev(DATA_ADD, "-", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM) Then
            '�` �W�[�r�� = �e�@�檺�W�[�r�� + ���檺�W�[�r��
            �Ҷq����ӼW�h���r��_��Ʈw = ADD_WORD_NUM + (CHANGE_ROW_WORD_NUM - ADD_WORD_NUM - InStrRev(DATA_ADD, "-", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM))
            Else
            �Ҷq����ӼW�h���r��_��Ʈw = ADD_WORD_NUM + (CHANGE_ROW_WORD_NUM - ADD_WORD_NUM - InStrRev(DATA_ADD, " ", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM))
        End If
    '�p�G���ݭp��A�B�Ӧ楽�������U�Ӧr���Ů�A�]���Ů椣�|�ݨ�U�@���Y�ӬO�ٲ��A�G�`�r�ƭn����
    ElseIf Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> " " And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM + 1, 1) = " " Then
        �Ҷq����ӼW�h���r��_��Ʈw = ADD_WORD_NUM - 1
        Else
        '�̫�@�إi�� �Ĥ@�楽�����Ů�,�U�@�Ӧr�����Ů�
        �Ҷq����ӼW�h���r��_��Ʈw = ADD_WORD_NUM
    End If
    Else
    '�p�G���ݭp��,�{�b �` �W�[�r��=�I�ܤW�@�欰� �` �W�[�r��
    �Ҷq����ӼW�h���r��_��Ʈw = ADD_WORD_NUM
End If
    
End Function


Public Function rangeSetBoardLine_��Ʈw(workbookName As String, shtName As String, starRange As String, endRange As String, myLineStyle As XlLineStyle)
'����d�� (e.g. A:A) �� �����d�� (e.g.A1:A3) �[�W��u
    '�L�ؽu "xlNone"
    '��u "xlContinuous"
    '�I "xlDot"
    '��u "xlDash"

    With Workbooks(workbookName).Worksheets(shtName).Range(starRange & ":" & endRange)
        If (myLineStyle = xlNone) Then
            .Borders(xlDiagonalDown).lineStyle = myLineStyle
            .Borders(xlDiagonalUp).lineStyle = myLineStyle
        End If
        .Borders(xlEdgeLeft).lineStyle = myLineStyle
        .Borders(xlEdgeTop).lineStyle = myLineStyle
        .Borders(xlEdgeBottom).lineStyle = myLineStyle
        .Borders(xlEdgeRight).lineStyle = myLineStyle
        .Borders(xlInsideVertical).lineStyle = myLineStyle
        .Borders(xlInsideHorizontal).lineStyle = myLineStyle
    End With
End Function
