Attribute VB_Name = "����Autocad�C�L"


Private Sub �C�L��pdf_��Ʈw()
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
Dim myFileDialog As FileDialog
Dim CanonicalMediaNameArray As Variant
Dim pdfSavePath As String
Dim iStr As String
Dim iBool As Boolean
'vvvvvv�e�m�ʧ@vvvvvv
    '�T�wautoCad�{�����}��
    Set acad = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "�Х��}��AutoCAD�{��"
        Exit Sub
    End If
    '���o�s����|
    Set myFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With myFileDialog
        .Title = "����ɮצs���Ƨ�"
        .AllowMultiSelect = False
    End With
    myFileDialog.Show
    If myFileDialog.SelectedItems.Count = 0 Then
        Exit Sub
    Else
        pdfSavePath = myFileDialog.SelectedItems.Item(1)
    End If
    '�}��
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    sdi = dwgFile.GetVariable("sdi")
    dwgFile.SetVariable "sdi", 0
    '�]�w ���n �I���C�L-�o�˵{���~�|���e�@�i�X���~�~�����X�ĤG�����ʧ@ (�_�h�e�@�i�٨S�X���{�����W�e�U�@�i���C�L���O�|�ɭP����)
    If (dwgFile.GetVariable("BACKGROUNDPLOT") <> 0) Then
        dwgFile.SetVariable "BACKGROUNDPLOT", 0
    End If
'vvvvvv�q�t�mlayout�X��vvvvvv
    '���olayout�W��
    Dim myLayouts As AcadLayout
    For Each myLayouts In dwgFile.Layouts
        iStr = myLayouts.name
    Next
    
    
    
    
'vvvvvv�q�ҫ�Model�X��vvvvvvv
    Dim blockName As String, blockWid As Double, blockHei As Double
    Dim blockXScaleFactor As Double, blockYScaleFactor As Double
    Dim obj As Object
    Dim blockMinPoint As Variant, blockMaxPoint As Variant
    Dim blockXScaleFactor As Double, blockYScaleFactor As Double
    '�ϥΪ̶���ʿ�J�϶��W�١B�e�B��(�ثe�d�L�۰ʧP�_���覡)
    blockName = "FRAM"
    blockWid = 123
    blockHei = 456
    '���o�ҫ����Ҧ��ϥθӹ϶����Ϯت����󪺥��U���M�k�W��
    Cells(1, 5) = "block���U��X"
    Cells(1, 6) = "block���U��Y"
    Cells(1, 7) = "block X ���"
    Cells(1, 8) = "block Y ���"
    Cells(1, 9) = "block�k�W��X=���U��X+(�϶��e*X���)"
    Cells(1, 10) = "block�k�W��Y=���U��Y+(�϶���*Y���)"
    Cells(1, 11) = "�X�ϵ��G"
    i = 1
    For Each obj In dwgFile.ModelSpace
        If TypeOf obj Is AcadBlockReference Then
            If (obj.name = blockName) Then
                i = i + 1
                '�u����blockMinPoint -> �϶����U���y��
                ' [0]X�y�� [1]Y�y��
                obj.GetBoundingBox blockMinPoint, blockMaxPoint
                '�϶����
                blockXScaleFactor = obj.XScaleFactor
                blockYScaleFactor = obj.YScaleFactor
                '�N���o��T�g��excel�W
                Cells(i, 5) = Round(blockMinPoint(0), 4)
                Cells(i, 6) = Round(blockMinPoint(1), 4)
                Cells(i, 7) = Round(blockXScaleFactor, 4)
                Cells(i, 8) = Round(blockYScaleFactor, 4)
                '�p��X�k�W���y�Шüg��excel
                Cells(i, 9) = Round(blockMinPoint(0) + (blockXScaleFactor * blockWid), 4)
                Cells(i, 10) = Round(blockMinPoint(1) + (blockYScaleFactor * blockHei), 4)
            End If
        End If
    Next
    '�p�G�ε���X�ϡA�ݭn�վ�DVIEW
    '   �ҫ��Ŷ�+�ϥε���X�ϥi��J�즹���D : �{����n���U�k�W�I�A���ϥ�acWindow�ɵ{�����w�������m�o���b�I�W
    '       https://forums.autodesk.com/t5/visual-basic-customization/window-plot-using-vba-strange-offset/m-p/9633693#M104077
    '   �]�wDVIEW�N�i�H�ѨM�����D
    '   https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Plot-or-preview-shows-incorrect-area-of-drawing-when-plotting-limits.html
    '   DVIEW��z
    '       https://knowledge.autodesk.com/support/autocad/learn-explore/caas/CloudHelp/cloudhelp/2019/ENU/AutoCAD-Core/files/GUID-E0078D09-8449-4A0A-A5AD-6984A01CEC33-htm.html
    '   DVIEW command�Բӻ���
    '       https://help.bricsys.com/hc/en-us/articles/360006566734-DView
    
    '   �]�wDVIEW-�ثe�䤣���vba���覡�A�G�u��ζǰecommand���覡
    dwgFile.SendCommand "DVIEW" & vbCr & "ALL" & vbCr & vbCr & "POINT" & vbCr & "0,0,0" & vbCr & "0,0,1" & vbCr & vbCr
    '   zoom�^����
    ZoomAll
    
    

'vvvvvvu�C�L�q�γ]�wvvvvvv
    '�����I������A�_�h�Ĥ@�i���٨S���槹�{���N���եX�ĤG�i�|�ɭP�L�k�X��;��U�W�������p�U
    'To plot in the foreground using VBA, you must set the BACKGROUNDPLOT system variable to 0. Otherwise, plotting occurs in the background.
    cadBackPlot = dwgFile.GetVariable("BACKGROUNDPLOT")
    acad.ActiveDocument.SetVariable "BACKGROUNDPLOT", 0
    '   �L���/ø�Ͼ�>�W��
    dwgFile.ActiveLayout.ConfigName = "DWG To PDF.pc3"
    '   �ϯȤj�p
    dwgFile.ActiveLayout.CanonicalMediaName = "ISO_expand_A3_(297.00_x_420.00_MM)"
    
    '   �X�Ͻd��>�X�Ϥ��e/plot area(�ұo�Ȭ�integer)
    '       0   acDisplay
    '       1   acExtents
    '       2   acLimits
    '       3   acView
    '       4   acWindow
    '       5   acLayout
    dwgFile.ActiveLayout.plotType = acExtents
    '       �ϥ�acWindow
    dwgFile.ModelSpace.Layout.plotType = acWindow
    '           ����X�Ͻd�� (0)x����(1)y����
    Dim leftDownPointArray(1) As Double, rightUpPointArray(1) As Double
    leftDownPointArray(0) = 0
    leftDownPointArray(1) = 0
    rightUpPointArray(0) = 200
    rightUpPointArray(1) = 100
    dwgFile.ModelSpace.Layout.SetWindowToPlot leftDownPointArray, rightUpPointArray
    
    
    '   �X�ϰ����q>�m���X��
    '       �`�N!plotType����acLayout�~�i�]�w���A�_�h�|����
    dwgFile.ActiveLayout.CenterPlot = True
    '   �X�Ϥ��>���
    '       ���-���
    dwgFile.ActiveLayout.PaperUnits = acMillimeters
    '       ���-�ۭq (�M �U�Ԧ���� �u��G��@)
    dwgFile.ActiveLayout.SetCustomScale 1, 2
    '       ���-�U�Ԧ����i��쪺 (�M �ۭq �u��G��@)
    '        dwgFile.ActiveLayout.StandardScale = ac1_2
    '       �ϭ���� --> �α��ਤ�רӹ�������/�/�W�U�A��
    dwgFile.ActiveLayout.plotRotation = ac90degrees
    '       �X�ϧΦ���
    dwgFile.ActiveLayout.StyleSheet = "1880-A3-thin.ctb"
    
    '   �N�I�����檬�A�]�����Ϫ���l��
    acad.ActiveDocument.SetVariable "BACKGROUNDPLOT", cadBackPlot
    
    dwgFile.Application.ZoomExtents

'vvvvvv�X�Ϩ��ɮ�vvvvvv
    
    '   ���b-�ϥΪ̦p�G��ܬY�w�ЦӫD�Y��Ƨ�
    If (Right(pdfSavePath, 1) = "\") Then
        iBool = dwgFile.Plot.PlotToFile(pdfSavePath & "myPdf.pdf")
    Else
        iBool = dwgFile.Plot.PlotToFile(pdfSavePath & "\" & "myPdf.pdf")
    End If
    '�N �O�_�X�Ϧ��\ �g��excel�W
    If (iBool = True) Then
        Cells(i, 11) = "�X�Ϧ��\"
    Else
        Cells(i, 11) = "�X�ϥ���"
    End If
    
    
    
    '** ��ĳ���� �X�ϫ�۰ʥ��}pdf�� �H�[�֥X�ϳt�� **
    '���}������ "�X�ϫ�۰ʥ��}pdf��" ���a��bcad�W:
    '���W "A"�ϥ� > �C�L > �L���/ø�Ͼ� �� > �W�� ���"DWG To PDF.pc3" >PDF�ﶵ ���}�A�N "�b�˵�������ܵ��G" �����Ŀ�

End Sub

Private Sub ���oCanonicalMediaName_��Ʈw()
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
Dim CanonicalMediaNameArray As Variant

Set acad = GetObject(, "AutoCAD.Application")

'�}��
Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
sdi = dwgFile.GetVariable("sdi")
dwgFile.SetVariable "sdi", 0
'�]�w �L���/ø�Ͼ�>�W��
dwgFile.ActiveLayout.ConfigName = "DWG To PDF.pc3"

'���o[�L���/ø�Ͼ�>�W��]�Ҧ��i�ϥΪ�CanonicalMediaName
CanonicalMediaNameArray = dwgFile.ActiveLayout.GetCanonicalMediaNames()

End Sub

Private Sub �ϥΦL����C�L�U�ɤ����Ҧ��t�m()
'�f�tfunctionSelectPrinter.bas/FormSelectPrinters.frm�ϥ�
'   �פJfunctionSelectPrinter.bas�A���L�ۦ��@��module
'   �פJFormSelectPrinters.frm
'�ݫŧi�������ܼơA��SUB/FORM�i�H���q
'   Public printerName As String
'�ϥΥH�U�{���X
Dim obj As Object
Dim acad As AcadApplication, dwgFile As AcadDocument, myLayout As AcadLayout
Dim i As Long, ii As Long, iii As Long
    '���b - �ˬdautocad�O�_�w�Q����
    On Error Resume Next
        Set obj = GetObject(, "AutoCAD.Application")
        If Err Then
            msgTitle = "�`�N"
            msgText = "�Х��}��AutoCAD�{��" + vbLf
            msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
            msgText = msgText + "����L�{�ݨϥΨ�AutoCAD�A�G�Ф�ʶ}�ҫ�A�װ��榹�{��"   ' �w�q�T��
            msgStyle = vbExclamation '���"!"�Ϯ�
        
            MsgBox msgText, msgStyle, msgTitle
            
            stopRun = True
            Exit Sub
        End If
    On Error GoTo 0
    '��ܦL���
    Form99Printers.Show
    If (stopRun = True) Then
        GoTo 991
    End If
    
    Set acad = GetObject(, "AutoCAD.Application")
    acad.Visible = True
    '�}��
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    '�]�msdi��0
    dwgFile.SetVariable "sdi", 0
    '�]�m ���n �I���C�L
    dwgFile.SetVariable "BACKGROUNDPLOT", 0
    
    ii = 2
    iii = 0
    For Each myLayout In dwgFile.Layouts
        '�����t�m�W��
        If (myLayout.name <> "Model") Then
            ii = ii + 1
            iii = iii + 1
            autocadToolSht.Cells(i, ii) = myLayout.name
            '�]�m�C�L�d��
            dwgFile.ActiveLayout = dwgFile.Layouts(myLayout.name)
            '�]�m�L���
            dwgFile.ActiveLayout.ConfigName = printerName
            '����C�L
            dwgFile.Plot.PlotToDevice
        End If
    Next

    dwgFile.Close False
    acad.Visible = True
991
End Sub
