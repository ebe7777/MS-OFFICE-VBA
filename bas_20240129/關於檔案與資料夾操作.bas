Attribute VB_Name = "�����ɮ׻P��Ƨ��ާ@"
Sub �`�N()
'VBA��IDE�u�䴩ANSI�A�ɭP�ܦh����S��r���|�L�k����(�{��Ū������ܬ�?)�i�ӾɭP����ɤ��_
'�ҥH�A���ϥ�FileSystemObject�����\��A�N�i�קK�����D

'(�覡1)�n�]�w�ޥζ���:Microsoft Scripting Runtime
Dim myFileSystemObject As New FileSystemObject, myFolder As Folder, myFile As File
'(�覡2)���γ]�w�ޥζ���
Dim objFSO As Object, objFolder As Object, objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(lastOpenFolderPath)
    For Each objFile In objFolder.Files
        '...
    Next
'�S��rŪ���쪺��]:
'VBA itself supports Unicode characters, the VBA development environment does not
'FileSystemObject support Unicode characters in file names
'��ƨӷ��G
'https://stackoverflow.com/questions/33685990/working-with-unicode-file-names-in-vba-using-dir-filesystemobject-etc

'�w�qAs New FileSystemObject, As Folder, As File�n�]�w�ޥζ���:
'Microsoft Scripting Runtime
'��ƨӷ��G
'https://trumpexcel.com/vba-filesystemobject/
End Sub



Sub ����~�������ɻP��Ƨ���()
'�R����Ƨ����ɮ�-�`�N!!�p�G�R���ɧ䤣��J���i�R�����ɮ׷|����
On Error Resume Next
    Kill "c:\aveva\test\*.*"
On Error GoTo 0
'�R����Ƨ�
RmDir "c:\aveva\test"
'�s�W��Ƨ�
MkDir "c:\aveva\test"
'�}�Ҹ�Ƨ�
Shell "explorer c:\aveva\test"
'�ƻs�ɮ�-�Na�ɽƻs��b��
Call FileSystem.FileCopy("C:\a.txt", "D:\b.TXT")
'�ƻs��Ƨ�-�Na��Ƨ������Ҧ���ƽƻs��b��Ƨ�
Call FileSystem.CopyFolder("c:\mydocuments\a*", "c:\b\")
'�إߤ@�Ӥ�r��
Dim newTextFileObj As Object
Set newTextFileObj = CreateObject("Scripting.FileSystemObject").CreateTextFile("D:\123.txt", True, True)
newTextFileObj.Write "your string goes here"
newTextFileObj.Close
'���s�R�W�ɮ�
Name "D:\123.py" As "D:\456.pp"
'�ƻs�ɮ�
FileCopy "D:\456.pp", "D:\123.py"
End Sub
Sub �}�Ҥu�@��_��Ʈw()
Dim filePath As String, fileName As String
filePath = "c:\123.xls"
fileName = "123.xls"
'��Ū / �~���s������s
Workbooks.Open fileName:=filePath, ReadOnly:=True, UpdateLinks:=0
Set nowWB = Workbooks(fileName)
Workbooks(fileName).Close SaveChanges:=False

               
End Sub
Sub �ާ@�������t�s�s�ɨæ۰�����_��Ʈw()


SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, fileFilter:="Excel�ɮ�(*.xls),*.xls", Title:="�t�s�s�ɦW��")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).copy
    ActiveWorkbook.SaveAs fileName:=SAVE_NAME
 '�]�����]�wAUTO_CLOS���ɮ�
    On Error Resume Next
    With Workbooks(ActiveWorkbook.Name)
        .RunAutoMacros xlAutoClos
        .Close
    End With
    End If
On Error GoTo 0
    Else
    MsgBox "�ާ@�����å��s��!"
End If
    
    
End Sub

Sub ���J���ɪ����w�u�@��_��Ʈw()
'��UserForm1,�����b��FORM��
End Sub
Sub ���ϥΪ̳]�w�s�ɦW��()
    lastSaveFullPath = "c:\123.xls"
    saveFullPath = Application.GetSaveAsFilename(InitialFileName:=lastSaveFullPath, fileFilter:="Excel 2003(*.xls),*.xls,Excel 2007(*.xlsx),*.xlsx", Title:="��ܦs�ɦ�m�P�W��")
    If saveFullPath <> "False" Then
        lastSaveFullPath = saveFullPath
    Else
        Exit Sub
    End If
End Sub

Sub ����@�Ӹ�Ƨ��è��o���|_��Ʈw()
'FileDialog�N�⦳�����P���W�l(set filePath)�A�u�|�����̫�@����ơA�ҥH�C���ϥΫ᳣�ݱN�o�������|�t�sstring
Dim lastSeleFolderPath As String
    '��ܸ�Ƨ��ɮɥ����ϥΪ̤�ʿ���A�ҥH���줧�e��������Ƨ����W�@�h
    lastSeleFolderPath = Left(sysSht.Range("d1"), InStrRev(sysSht.Range("d1"), "\"))
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "����ɮצs���m"
        .InitialFileName = lastSeleFolderPath
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            GoTo 991
        Else
            lastSeleFolderPath = .SelectedItems(1)
            
            sysSht.Range("d1") = lastSeleFolderPath
        End If
    End With
End Sub
Sub ����@���ɮרè��o���|_��Ʈw()
'FileDialog�N�⦳�����P���W�l(set filePath)�A�u�|�����̫�@����ơA�ҥH�C���ϥΫ᳣�ݱN�o�������|�t�sstring
Dim lastSeleFilePath As String
    lastSeleFilePath = sysSht.Range("b1")
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = lastSeleFilePath
        .Title = "��ܼ˥��� (���)"
        .AllowMultiSelect = False
        .Filters.Add "Word", "*.doc;*.docx", 1
        .Filters.Add "Excel", "*.xls;*.xlsx", 2
        .Filters.Add "��L", "*.*", 3
        .FilterIndex = 1
        .Show
        If .SelectedItems.Count = 0 Then
            GoTo 991
        Else
            lastSeleFilePath = .SelectedItems(1)
            sysSht.Range("b1") = lastSeleFilePath
            'sysSht.Range("b1") = Left(lastSeleFilePath, InStrRev(lastSeleFilePath, "\"))
        End If
    End With
End Sub
Sub �t�s�s��()
Dim saveFullPath As String
Dim lastSeleFolderPath As String
    '�W���s�ɸ�Ƨ��AĴ�p c:\123\
    lastSeleFolderPath = sysSht.Range("d1")
    saveFullPath = Application.GetSaveAsFilename(InitialFileName:=lastSeleFolderPath, fileFilter:="Excel, *.xlsx")
    '   ���b-�����h������
    If (CStr(saveFullPath) = "False") Then
        MsgBox "�Э��s����"
    Else
        '�t�s�s��
        ActiveWorkbook.SaveAs fileName:=saveFullPath
        ActiveWorkbook.Activate
        MsgBox "�w�t�s�s��"
    End If
End Sub

Sub �}�Ҥ@��Ƨ��U�Ҧ��S�w�ɮ�_��Ʈw()
Dim allFile As String, filePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "��ܭn�j�M����Ƨ�"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    '���Ĥ@�ӯS�w���ɦW���ɮ�
    eachFile = Dir(filePath & "\*.docx*")
    Do While allFile <> ""
        Workbooks.Open fileName:=filePath & eachFile
        '�N�U�@��do���ؼ��ɮײ��ܤU�@�ӯS�w���ɦW���ɮ�
        eachFile = Dir()
    Loop
End Sub
Sub ���o�@���|���U�Ҧ��ɮ�()
Dim mainFolderDirectory As String
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "��ܭn�j�M����Ƨ�"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        mainFolderDirectory = .SelectedItems(1)
    End With
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(mainFolderDirectory)
    i = 1
    For Each objFile In objFolder.Files
        Cells(i + 1, 1) = objFile.Name
        Cells(i + 1, 2) = objFile.Path
        i = i + 1
    Next objFile
End Sub
Sub ���o�@���|���U�Ҧ�����Ƨ�_��Ʈw()

Dim mainFolderDirectory As String
Dim objFSO As Object
Dim subFolders As Object
Dim subFoldersCount As Integer
Dim subFolder As Object

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "��ܭn�j�M����Ƨ�"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        mainFolderDirectory = .SelectedItems(1)
    End With
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set subFolders = objFSO.getfolder(mainFolderDirectory).subFolders
    
    subFoldersCount = subFolders.Count
    
    For Each subFolder In subFolders
        'do something...
        'subPath = subFoler.path
    Next subFolder

End Sub
Sub �s�W��Ƨ�_��Ʈw()

Dim fso
Dim sFolder As String
    
    sFolder = "C:\SampleFolder" ' You can Specify Any Path and Name To Create a Folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(sFolder) Then
    fso.CreateFolder (sFolder) 'Checking if the same Folder already exists
    MsgBox "New FolderCreated Successfully", vbExclamation, "Done!"
    Else
    MsgBox "Specified Folder Already Exists", vbExclamation, "Folder Already Exists!"
    End If

End Sub
Function getOneTypeFilesUnderFolder_�i���I�S��r(folderPath, extensionName, myArray)
'�N���|folderPath���U�Ҧ����ɦW��extensionName(Ĵ�p�G".txt")���ɮ׸�T�g�imyArray��
'myArray�����O2���}�C
'   �Ĥ@���O��ƺ��� - ����2,(1)�����ɦW(2)�L���ɦW���ɦW
'   �ĤG���O��Ƶ��� - �L����A�|�Happend�覡�[�W
Dim eachFile As String, fileNameWithExt As String, fileExt As String, fileName As String, fileFullPath As String
Dim myFileSystemObject As Object, myFolder As Object, myFile As Object

    Set myFileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set myFolder = myFileSystemObject.getfolder(folderPath & "\")
    
    For Each myFile In myFolder.Files
        fileNameWithExt = myFile.Name
        fileExt = Right(fileNameWithExt, Len(fileName) - InStrRev(fileNameWithExt, "."))
        If (fileExt = extensionName) Then
            fileName = Left(fileNameWithExt, Len(fileNameWithExt) - Len(extensionName))
            ReDim Preserve myArray(2, UBound(myArray, 2) + 1)
            '[1] fileFullPath [2]fileName
            myArray(1, UBound(myArray, 2)) = myFile.Path
            myArray(2, UBound(myArray, 2)) = fileName
        End If
    Next
End Function
Function getOneTypeFilesUnderFolder_�L�k���I�S��r(folderPath, extensionName, myArray)
'�N���|folderPath���U�Ҧ����ɦW��extensionName(Ĵ�p�G".txt")���ɮ׸�T�g�imyArray��
'myArray�����O2���}�C
'   �Ĥ@���O��ƺ��� - ����2,(1)�����ɦW(2)�L���ɦW���ɦW
'   �ĤG���O��Ƶ��� - �L����A�|�Happend�覡�[�W
Dim eachFile As String, fileNameWithExt As String, fileName As String, fileFullPath As String
    fileFullPath = folderPath & "\*" & extensionName
    eachFile = Dir(fileFullPath)
    Do While eachFile <> ""
        fileNameWithExt = Right(eachFile, Len(eachFile) - InStrRev(eachFile, "\"))
        fileName = Left(fileNameWithExt, Len(fileNameWithExt) - Len(extensionName))
        ReDim Preserve myArray(2, UBound(myArray, 2) + 1)
        '[1] fileFullPath [2]fileName
        myArray(1, UBound(myArray, 2)) = eachFile
        myArray(2, UBound(myArray, 2)) = fileName
        '�N�U�@��do���ؼ��ɮײ��ܤU�@��
        eachFile = Dir()
    Loop
'�S��rŪ���쪺��]:
'VBA itself supports Unicode characters, the VBA development environment does not
'FileSystemObject support Unicode characters in file names
'��ƨӷ��G
'https://stackoverflow.com/questions/33685990/working-with-unicode-file-names-in-vba-using-dir-filesystemobject-etc

'�w�qAs New FileSystemObject, As Folder, As File�n�]�w�ޥζ���:
'Microsoft Scripting Runtime
'��ƨӷ��G
'https://trumpexcel.com/vba-filesystemobject/
End Function
Function ���o�@���|���U�Ҧ��S�w���ɦW���ɮ�()
'���o���|�����Ӥ���
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(lastOpenFolderPath)
    i = 0
    For Each objFile In objFolder.Files
        '���ɦW�Obmp,jpg�~�B�z
        tempStr1 = Right(objFile.Name, Len(objFile.Name) - InStrRev(objFile.Name, "."))
        If (LCase(tempStr1) = "jpg" Or LCase(tempStr1) = "bmp") Then
            'do somethig
        End If
    Next
End Function
Function getAllFolderUnderThisPath_���k���o�@��Ƨ����U�U�h���l��Ƨ�(folderPath, myArray)
'�ϥλ��k�覡�A�NfolderPath���U�Ҧ����h����Ƨ����|����X,Ĵ�p
'   1.folderPath���U��������Ƨ� > ������|��myArray
'   2.folderPath���U���Y��Ƨ��̪�������Ƨ� > ������|��myArray
'   3.folderPath���U���Y��Ƨ��̪��Y��Ƨ��̪�������Ƨ� > ������|��myArray
'   4. ...
'myArray�����O1���}�C
'   �s��Ʒ|append���¸�ƤW
Dim objFSO As Object
Dim subFolders As Object
Dim subFoldersCount As Integer
Dim subFolder As Object
Dim i As Integer
    ReDim Preserve myArray(UBound(myArray, 1) + 1)
    myArray(UBound(myArray, 1)) = folderPath
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set subFolders = objFSO.getfolder(folderPath).subFolders
    If (subFolders.Count > 0) Then
        For Each subFolder In subFolders
            Call getAllFolderUnderThisPath(subFolder.Path, myArray)
        Next subFolder
'    Else
'        ReDim Preserve myArray(UBound(myArray, 1) + 1)
'        myArray(UBound(myArray, 1)) = folderPath
    End If

End Function

Sub �ˬd��Ƨ��O�_�s�b_��Ʈw()
Dim folderFullPath As String
folderFullPath = "c:\123"

    If Dir(folderFullPath, vbDirectory) = vbNullString Then
        'do something
    End If
    
End Sub
Sub �ˬd�ɮ׬O�_�s�b_��Ʈw()
Dim fileFullPath As String
fileFullPath = "C:\�J���H���t��\02��r���\�J���H���M��.xlsm"
    If Dir(fileFullPath) = Empty Then
         MsgBox "��󤣦s�b�C"
    End If
End Sub
Sub ���o�ɦW()
Dim fs, fos, fd, fc, aaa, bbb
Set fos = CreateObject("Scripting.FileSystemObject")
Set fd = fos.getfolder("C:\�J���H���t��\01�Ϥ����\X120251959_���ӱj") '�ɮץؿ�
Set fc = fd.Files

For Each fs In fc
    ThisWorkbook.Sheets("test").Cells(1, 1) = fs.Name
    aaa = ThisWorkbook.Sheets("test").Cells(1, 1)
Next
End Sub
Sub �����r�ɷs�W��Ū��()
'�� �����r�ɾާ@.bas
End Sub

'20231114 �������q�_���ի��A��J�{���y�k
Function ��Ʈw_checkIfFileNameContainUnacceptableCharacter(fileName As String) As Boolean
'�ɦW���i�ϥγo�Ǧr���Ÿ� \ / : * ? "" < > |
Dim myArr(9)
Dim i As Long
    checkIfFileNameContainUnacceptableCharacter = False
    errString = ""
    myArr(1) = "\"
    myArr(2) = "/"
    myArr(3) = ":"
    myArr(4) = "*"
    myArr(5) = "?"
    myArr(6) = """"
    myArr(7) = "<"
    myArr(8) = ">"
    myArr(9) = "|"
    '���b-�p�G��J�ȬO�ťիh���L
    '   ���\����ˬd�O�_�]�t���i�ϥΪ��r���Ÿ��A������L�ˬd
    If (Trim(fileName <> "")) Then
        For i = 1 To UBound(myArr, 1)
            If (InStr(1, fileName, myArr(i)) <> 0) Then
                checkIfFileNameContainUnacceptableCharacter = True
                Exit For
            End If
        Next i
    End If
    
'    If (errString <> "") Then
'        msgTitle = "���~            "    ' �w�q���D�C
'        msgText = "��J���s�����H�U���~�A�Эק�" + vbLf    ' �w�q�T���C
'        msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
'        msgText = msgText + iStr2 + vbLf  ' �w�q�T��
'        msgStyle = vbCritical '���"X"�Ϯ�
'        MsgBox msgText, msgStyle, msgTitle
'    End If
End Function





Sub old_���o��@�ɮ׸��|_��Ʈw()
'*�p�G�P�@�q�{�����ϥΨ�Ӧ��覡�A�B�ѰϤ��P�A�h�ĤG�ӨϥΦ��k���|����-->���FileDialog
Dim weldFilePathList As String
    '���w�w�]���|
    ChDir Left("d:\123.csv", InStrRev("d:\123.csv", "\"))
    weldFilePathList = Application.GetOpenFilename(fileFilter:="(*.csv), *.csv", Title:="���WeldData�� (���i�ƿ�)", MultiSelect:=False)
    If weldFilePathList = "False" Then
        GoTo 991
    End If
End Sub
Sub old_���o�Ƽ��ɮ׸��|_��Ʈw()
'*�p�G�P�@�q�{�����ϥΨ�Ӧ��覡�A�B�ѰϤ��P�A�h�ĤG�ӨϥΦ��k���|����-->���FileDialog
Dim FilePathList() As Variant
    '���w�w�]���|
    ChDir Left("d:\123.dwg", InStrRev("d:\123.dwg", "\"))
    On Error GoTo 991
        '�P�ɿz��Xdwg,dxf
        FilePathList = Application.GetOpenFilename(fileFilter:="(*.dwg;*.dxf), *.dwg;*.dxf", Title:="��ܭn�X����cad�� (�i�ƿ�)", MultiSelect:=True)
        '���b�U�Ԧ�����ܭn�}��dwg�Ϊ�dxf
        FilePathList = Application.GetOpenFilename(fileFilter:="(*.dwg), *.dwg,(*.dxf), *.dxf", Title:="��ܭn�X����cad�� (�i�ƿ�)", MultiSelect:=True)
    On Error GoTo 0

End Sub

