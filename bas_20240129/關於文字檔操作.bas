Attribute VB_Name = "�����r�ɾާ@"
Sub Ū����r�ɨä��r1_�u��Ūansi()

Dim linesArray
Dim splitArray

'�覡2�~��dim�H�U
Dim readData As String
Dim i As Long

'�}���ɮ�(���Ҭ��u�}�Ҥ@���ɮ�)
'#1 file number
'   Number used in the Open statement to open a file. Use file numbers in the range 1-255, inclusive, for files not accessible to other applications. Use file numbers in the range 256-511 for files accessible from other applications.
Open weldFilePathList For Input As #1
'�N���e�@����iarray��
'==============
'�覡1 - �i��|�����D
'LOF�N��Ū���̭����r��(�U��ۥ[)�A��pdms�X��bom��������vba�o�쪺�r�Ƥ��ɮ׹�ڦr�Ʀh�G�X�{���~62
linesArray = split(Input$(LOF(1), #1), vbNewLine)

'�覡2 - �i�קK���覡1�����D
i = 0
Do Until EOF(1)
    Line Input #1, readData
    If Not Left(readData, 1) = "*" Then
        i = i + 1
        ReDim Preserve linesArray(i)
        linesArray(i) = readData
    End If
Loop
'==============

'�����ɮ�
Close #1
'�@��椺�e�H"^"��}��g�iArray
For i = 0 To UBound(linesArray, 1)
    If linesArray(i) <> "" Then
        ReDim Preserve weldDataArray(5, i)
        splitArray = split(linesArray(i), "^")
        For ii = 0 To 4
            weldDataArray(ii + 1, i) = splitArray(ii)
        Next ii
    End If
Next i

End Sub

Sub Ū����r��2_��Ūunicode�Pansi()
Dim fso As Object
Dim myTxtFile As Object
Dim i As Long
Set fso = CreateObject("Scripting.FileSystemObject")
'��������o��iomode/Format����ϥ�constant�A�����ϥμƦr
'iomode:= (1)��Ū (2)overwrite (3)append
'Create:= ���w��fileName�p�G���s�b (true)�طs��(false)���طs��
'Format:=  (0)�}���ɬOansi (-1)�}���ɬOUnicode (-2)�ϥΨt�ηU�w�]�ȶ}����
Set myTxtFile = fso.OpenTextFile(fileName:="c:\123.txt", iomode:=1, Create:=False, Format:=-1)
'Ū��������ɳ̫�@��
Do Until myTxtFile.AtEndOfStream
    i = i + 1
    'Ū���ɮפ��C�C���
    Sheets("123").Range("A" & i) = myTxtFile.ReadLine
Loop
End Sub
Sub ��X��ƨ��r��1_�u�䴩ansi�榡()

Dim filePath As String
Dim SDTEROWS As Long, SMTEROWS As Long, ROWS As Long
Dim SDTECOLUMNSTRING As String, SMTECOLUMNSTRING As String
'==================================================================================
'
'��X��ƨ��r��
'----------------------------------------------------------------------------------
filePath = "D:\SXTE.TXT"
SDTEROWS = Sheets(SDTE).Range("A65536").End(xlUp).Row
SMTEROWS = Sheets(SMTE).Range("A65536").End(xlUp).Row
SDTECOLUMNSTRING = "E"
SMTECOLUMNSTRING = "D"
'�b���|�۰ʷs�W�ɮ�
Open filePath For Output As #1
For ROWS = 1 To SDTEROWS
    Print #1, Sheets(SDTE).Range(SDTECOLUMNSTRING & ROWS).Value
Next ROWS
For ROWS = 1 To SMTEROWS
    Print #1, Sheets(SMTE).Range(SMTECOLUMNSTRING & ROWS).Value
Next ROWS
Close #1
'Ref WEB Site
'http://www.excel-easy.com/vba/examples/write-data-to-text-file.html
'https://www.mrexcel.com/forum/excel-questions/6030-write-text-file-without-quotes-vba.html
End Sub
Sub ��X��ƨ��r��2_�䴩ansi�Punicode�榡()
Dim fso As Object
Dim myTxtFile As Object
Dim i As Integer
    '�s�W�@�Ӥ�r��
    Set fso = CreateObject("Scripting.FileSystemObject")
    '   �p�G�Oansi,�̫᪺Unicode:=false
    Set myTxtFile = fso.CreateTextFile(fileName:="c:\123.txt", OverWrite:=True, Unicode:=True)
    '�g�J���
    For i = 1 To 5
        myTxtFile.Write i & vbCrLf
    Next i
    '����
    myTxtFile.Close
    '�t�@�ӷs�Wunicode��r�ɪ��覡�O �g�Jexcel��>excel�t�s��unicode�榡txt
    'ActiveWorkbook.SaveAs fileName:="D:\123.txt", FileFormat:=xlUnicodeText, CreateBackup:=False
End Sub
