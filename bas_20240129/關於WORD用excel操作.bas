Attribute VB_Name = "����WORD��excel�ާ@"
Sub �s�Wword��()
Dim objWordApp As Object, objWordDoc As Object
    'Create the object of Microsoft Word
    Set objWordApp = CreateObject("Word.Application")
    'Add documents to the Word
    Set objWordDoc = objWordApp.Documents.Add
    'Make the MS Word Visible
    objWordApp.Visible = True

End Sub
Sub �}��word��()
'20210629
Dim objWordApp As Object, objWordDoc As Object
        Set objWordApp = CreateObject("Word.Application")
'===============
'�O�_���word�{��-���ծɻݶ}�ҡA�_�h�J����~��word�{���|�b�I���}�O���Ҹ�word�ɾɭP�ɮ׳Q��w
'wordApp.Visible = True
wordApp.Visible = False
'===============
        objWordApp.ScreenUpdating = False
        '�}��WORD��,��Ū
        Set objWordDoc = objWordApp.Documents.Open(fileName:=wordSampleFileFullPath, ReadOnly:=True)
End Sub


Sub �s������word�{��()
Dim saveFullPath As String
        
    '�t�s�s��,�t�s���w�]�榡 (saveFullPath�]�t���|�P�ɦW�A���i�ٲ����ɦW)
    objWordDc.SaveAs fileName:=saveFullPath, FileFormat:=wdFormatDocumentDefault
    '�t�s�s��,�t�s��docx�榡 (saveFullPath�]�t���|�P�ɦW�A���i�ٲ����ɦW)
    objWordDc.SaveAs fileName:=saveFullPath, FileFormat:=wdFormatXMLDocument
    '�s��
    objWordApp.Documents(saveFullPath).Save
    '����WORD��
    objWordApp.Documents(saveFullPath).Close
    '������WORD�{��
    objWordApp.Quit
End Sub
Sub word���������()
    '�Nword������j�̤j
    objWordDoc.ActiveWindow.WindowState = wdWindowStateMaximize
    'focus��word����
    objWordApp.Activate
    '������s�e��
    objWordApp.ScreenUpdating = False
End Sub
Sub ��¶�Ϥ覡�]�w()
'wdWrapInline    7   �ϧλP��r�ƦC�C
'wdWrapNone      3   �|�N�ϧθm���r���e�C �t�аѾ\wdWrapFront�C
'wdWrapSquare    0   ��r��¶�ϧδ���C �汵��O�b�ϧΪ��ۤϤ@���C
'wdWrapThrough   2   ��r��¶�ϧδ���C
'wdWrapTight     1   ��r�b�ϧήǺ�K����C
'wdWrapTopBottom 4   �N��r�m��ϧΤ��W�P���U�C
'wdWrapBehind    5   �N�ϧθm���r���C
'wdWrapFront     6   �N�ϧθm���r�e��C
End Sub
Sub word�����榡()
    '�W�U���k���
    ' �^�TInchesToPoints
    ' ����CentimetersToPoints
    objWordDoc.PageSetup.TopMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.BottomMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.RightMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.LeftMargin = CentimetersToPoints(1)
    
End Sub

Sub ���_table�����a1()
Dim objWordRange As Object, objWordTab As Object
Dim wordTableColumns As Integer, wordTableRows As Integer
Dim i As Integer, ii As Integer, iii As Integer
'�]�wtable row,colum�ƶq
wordTableRows = 2
wordTableColumns = 3
    'Create a Range object.
    Set objWordRange = objWordDoc.Range
    'Create Table using Range object and define no of rows and columns.
    objWordDoc.Tables.Add Range:=objWordRange, NumRows:=wordTableRows, NumColumns:=wordTableColumns
    'Get the Table object
    Set objWordTab = objWordTab.Tables(1)
    With objWordTab
        '�Ҧ��檺��r
        .Range.Font.Size = 25
        '�Ҧ��檺�e
        ' �^�TInchesToPoints
        ' ����CentimetersToPoints
        .PreferredWidth = CentimetersToPoints(10)
        With objWordTab
        '�Ϯ�u�i��(�ϥΰ򥻽u��)
        .Borders.Enable = True
        '�]�w��u�Φ�
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        For i = 1 To wordTableRows
            For ii = 1 To wordTableColumns
                '��J���
                iii = iii + 1
                iStr = dataArry(iii, 1) & Chr(10) & dataArry(iii, 2) & Chr(10) & dataArry(iii, 3)
                .Cell(i, ii).Range.Text = iStr
                '�C�C��
                .Rows(i).Height = 300
                '�C�C����align
                .Rows(i).Cells.VerticalAlignment = wdAlignVerticalCenter
                '�C�C�����C�����align
                .Rows(i).Cells(ii).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Next ii
        Next i
    End With
    

End Sub
Sub ��檺�ާ@2_�C����O������r()
Dim objWordCell As Object
    '�]�w���wcell
    Set objWordCell = objWordDoc.Range(Start:=.Tables(1).Cell(i, ii).Range.Start, _
        End:=.Tables(1).Cell(i, ii).Range.End)
        '���[����-������wcell
        'objWordCell.Select
    '�ק���wcell�����w�r�j�p-�H�������覡
    With objWordCell.Find
        .ClearFormatting
        .ClearAllFuzzyOptions
        '�j�p�g���ŦX
        .MatchCase = True
        '���r���g���ŦX
        .MatchWholeWord = True
        '��즹��r
        .Text = "test"
        '����������r
        .Replacement.Text = "123"
        .Replacement.Font.Size = 20
        .Execute Format:=True, Replace:=wdReplaceAll
    End With
End Sub
Sub ������r�P�ק�r�j�p()
    With objWordDoc.Content.Find
        .ClearFormatting
        .ClearAllFuzzyOptions
        '�j�p�g���ŦX
        .MatchCase = True
        '���r���g���ŦX
        .MatchWholeWord = True
        '��즹��r
        .Text = "ABC"
        '����������r
        With .Replacement
            .Text = "123"  'this line might be unnecessary
            .Font.Size = 20
        End With
        
        '�ϥ�early binding
        .Execute Format:=True, Replace:=wdReplaceAll
        '�ϥ�late binding
        .Execute Format:=True, Replace:=2
        '���� https://social.msdn.microsoft.com/Forums/en-US/452d6d5f-e485-41a2-967c-27e739cc8d9e/hard-to-solve-ms-word-late-binding-problem?forum=isvvba
        'Replace:=2 �N��N�� https://docs.microsoft.com/en-us/office/vba/api/word.wdreplace
    End With
End Sub

Sub ���e�g��_Ū���J���ɮץt�s��N�̭����Ϥ�������()
'���]�w �u��>�]�w�ޥζ���>Microsoft Word xx.x Object Library
Dim objWordApp As Object, objWordDoc As Object
Dim fileFullPath As String, fileSavePath As String, fileSaveName As String
Dim saveFullPath As String
Dim originalImage As InlineShape, newImage As InlineShape
Dim imageControl As ContentControl
Dim imageW As Long, imageH As Long
Dim newImagePath As String

fileFullPath = "D:\Bruce�u�@���\�i���ӽЮѲ��͵{��\���եθ��\����-�H���J�t��ƶ��D������(SAMPLE)_�w�]�w������r�P�Ϥ�.docx"
fileSavePath = "D:\Bruce�u�@���\�i���ӽЮѲ��͵{��\���եθ��"
fileSaveName = "test.docx"
saveFullPath = fileSavePath & "\" & fileSaveName
    Set objWordApp = New Word.Application
    objWordApp.Visible = True
    '�}��WORD��,��Ū
    Set objWordDoc = objWordApp.Documents.Open(fileName:=fileFullPath, ReadOnly:=True)
    '�t�s�s��
    objWordDoc.SaveAs fileName:=saveFullPath
    '�ѩ�쥻�O��Ū���A�}�ҡA�ɭP�t�s��L�k�s��A�G�b�t�s���������}���ͪ��s��
    objWordApp.Documents(saveFullPath).Close
    Set objWordDoc = objWordApp.Documents.Open(fileName:=saveFullPath, ReadOnly:=False)
    '���N��r
    With objWordDoc.Content.Find
        .Text = "�|�Ũt�ά�ުѥ��������q"
        .Replacement.Text = "TEST"
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    
    '���NInlineShape�Ϯ�

'    For Each iShape In ActiveDocument.InlineShapes
'        iShape.ConvertToShape
'        a = iShape.Title
'    Next iShape
    Set originalImage = objWordDoc.InlineShapes(1)

    If originalImage.Range.ParentContentControl Is Nothing Then
        Set imageControl = objWordDoc.ContentControls.Add(wdContentControlPicture, originalImage.Range)
    Else
        Set imageControl = originalImage.Range.ParentContentControl
    End If


    imageW = originalImage.width
    imageH = originalImage.Height

    originalImage.Delete

    
    newImagePath = "D:\Bruce�u�@���\�i���ӽЮѲ��͵{��\���եθ��\�o�ӬO�@�θ�Ƹ�Ƨ�\���ܰʹ���\�j�Y��\C123456789_�L����_P.jpg"
    objWordDoc.InlineShapes.AddPicture newImagePath, False, True, imageControl.Range

    With imageControl.Range.InlineShapes(1)
        .Height = imageH
        .width = imageW
    End With
    
    
    
    '�s�ɨ�����WORD��
    objWordApp.Documents(saveFullPath).Save
    objWordApp.Documents(saveFullPath).Close
    
    objWordApp.Quit
    '����ĵ�i
    objWordApp.DisplayAlerts = True
End Sub
