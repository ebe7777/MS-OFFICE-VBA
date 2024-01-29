Attribute VB_Name = "關於檔案與資料夾操作"
Sub 注意()
'VBA的IDE只支援ANSI，導致很多中文特殊字都會無法辨識(程式讀取時顯示為?)進而導致執行時中斷
'所以，應使用FileSystemObject相關功能，就可避免此問題

'(方式1)要設定引用項目:Microsoft Scripting Runtime
Dim myFileSystemObject As New FileSystemObject, myFolder As Folder, myFile As File
'(方式2)不用設定引用項目
Dim objFSO As Object, objFolder As Object, objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(lastOpenFolderPath)
    For Each objFile In objFolder.Files
        '...
    Next
'特殊字讀不到的原因:
'VBA itself supports Unicode characters, the VBA development environment does not
'FileSystemObject support Unicode characters in file names
'資料來源：
'https://stackoverflow.com/questions/33685990/working-with-unicode-file-names-in-vba-using-dir-filesystemobject-etc

'定義As New FileSystemObject, As Folder, As File要設定引用項目:
'Microsoft Scripting Runtime
'資料來源：
'https://trumpexcel.com/vba-filesystemobject/
End Sub



Sub 關於外部控制檔與資料夾案()
'刪除資料夾內檔案-注意!!如果刪除時找不到既有可刪除的檔案會報錯
On Error Resume Next
    Kill "c:\aveva\test\*.*"
On Error GoTo 0
'刪除資料夾
RmDir "c:\aveva\test"
'新增資料夾
MkDir "c:\aveva\test"
'開啟資料夾
Shell "explorer c:\aveva\test"
'複製檔案-將a檔複製成b檔
Call FileSystem.FileCopy("C:\a.txt", "D:\b.TXT")
'複製資料夾-將a資料夾中的所有資料複製到b資料夾
Call FileSystem.CopyFolder("c:\mydocuments\a*", "c:\b\")
'建立一個文字檔
Dim newTextFileObj As Object
Set newTextFileObj = CreateObject("Scripting.FileSystemObject").CreateTextFile("D:\123.txt", True, True)
newTextFileObj.Write "your string goes here"
newTextFileObj.Close
'重新命名檔案
Name "D:\123.py" As "D:\456.pp"
'複製檔案
FileCopy "D:\456.pp", "D:\123.py"
End Sub
Sub 開啟工作表_資料庫()
Dim filePath As String, fileName As String
filePath = "c:\123.xls"
fileName = "123.xls"
'唯讀 / 外部連結不更新
Workbooks.Open fileName:=filePath, ReadOnly:=True, UpdateLinks:=0
Set nowWB = Workbooks(fileName)
Workbooks(fileName).Close SaveChanges:=False

               
End Sub
Sub 操作中視窗另存新檔並自動關閉_資料庫()


SHEET_NAME = ActiveSheet.Name

SAVE_NAME = Application.GetSaveAsFilename(InitialFileName:=SHEET_NAME, fileFilter:="Excel檔案(*.xls),*.xls", Title:="另存新檔名稱")

If SAVE_NAME <> False Then
    
    Sheets(SHEET_NAME).copy
    ActiveWorkbook.SaveAs fileName:=SAVE_NAME
 '因應有設定AUTO_CLOS的檔案
    On Error Resume Next
    With Workbooks(ActiveWorkbook.Name)
        .RunAutoMacros xlAutoClos
        .Close
    End With
    End If
On Error GoTo 0
    Else
    MsgBox "操作取消並未存檔!"
End If
    
    
End Sub

Sub 載入它檔的指定工作表_資料庫()
'詳UserForm1,巨集在該FORM內
End Sub
Sub 讓使用者設定存檔名稱()
    lastSaveFullPath = "c:\123.xls"
    saveFullPath = Application.GetSaveAsFilename(InitialFileName:=lastSaveFullPath, fileFilter:="Excel 2003(*.xls),*.xls,Excel 2007(*.xlsx),*.xlsx", Title:="選擇存檔位置與名稱")
    If saveFullPath <> "False" Then
        lastSaveFullPath = saveFullPath
    Else
        Exit Sub
    End If
End Sub

Sub 選取一個資料夾並取得路徑_資料庫()
'FileDialog就算有取不同的名子(set filePath)，只會紀錄最後一筆資料，所以每次使用後都需將這次的路徑另存string
Dim lastSeleFolderPath As String
    '選擇資料夾時時必須使用者手動選取，所以移到之前紀錄的資料夾的上一層
    lastSeleFolderPath = Left(sysSht.Range("d1"), InStrRev(sysSht.Range("d1"), "\"))
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "選擇檔案存放位置"
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
Sub 選取一個檔案並取得路徑_資料庫()
'FileDialog就算有取不同的名子(set filePath)，只會紀錄最後一筆資料，所以每次使用後都需將這次的路徑另存string
Dim lastSeleFilePath As String
    lastSeleFilePath = sysSht.Range("b1")
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = lastSeleFilePath
        .Title = "選擇樣本檔 (單選)"
        .AllowMultiSelect = False
        .Filters.Add "Word", "*.doc;*.docx", 1
        .Filters.Add "Excel", "*.xls;*.xlsx", 2
        .Filters.Add "其他", "*.*", 3
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
Sub 另存新檔()
Dim saveFullPath As String
Dim lastSeleFolderPath As String
    '上次存檔資料夾，譬如 c:\123\
    lastSeleFolderPath = sysSht.Range("d1")
    saveFullPath = Application.GetSaveAsFilename(InitialFileName:=lastSeleFolderPath, fileFilter:="Excel, *.xlsx")
    '   防呆-取消則不執行
    If (CStr(saveFullPath) = "False") Then
        MsgBox "請重新執行"
    Else
        '另存新檔
        ActiveWorkbook.SaveAs fileName:=saveFullPath
        ActiveWorkbook.Activate
        MsgBox "已另存新檔"
    End If
End Sub

Sub 開啟一資料夾下所有特定檔案_資料庫()
Dim allFile As String, filePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "選擇要搜尋的資料夾"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    '找到第一個特定副檔名的檔案
    eachFile = Dir(filePath & "\*.docx*")
    Do While allFile <> ""
        Workbooks.Open fileName:=filePath & eachFile
        '將下一個do的目標檔案移至下一個特定副檔名的檔案
        eachFile = Dir()
    Loop
End Sub
Sub 取得一路徑底下所有檔案()
Dim mainFolderDirectory As String
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "選擇要搜尋的資料夾"
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
Sub 取得一路徑底下所有的資料夾_資料庫()

Dim mainFolderDirectory As String
Dim objFSO As Object
Dim subFolders As Object
Dim subFoldersCount As Integer
Dim subFolder As Object

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "選擇要搜尋的資料夾"
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
Sub 新增資料夾_資料庫()

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
Function getOneTypeFilesUnderFolder_可應付特殊字(folderPath, extensionName, myArray)
'將路徑folderPath底下所有副檔名為extensionName(譬如：".txt")的檔案資訊寫進myArray裡
'myArray必須是2維陣列
'   第一維是資料種類 - 須為2,(1)完整檔名(2)無副檔名的檔名
'   第二維是資料筆數 - 無限制，會以append方式加上
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
Function getOneTypeFilesUnderFolder_無法應付特殊字(folderPath, extensionName, myArray)
'將路徑folderPath底下所有副檔名為extensionName(譬如：".txt")的檔案資訊寫進myArray裡
'myArray必須是2維陣列
'   第一維是資料種類 - 須為2,(1)完整檔名(2)無副檔名的檔名
'   第二維是資料筆數 - 無限制，會以append方式加上
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
        '將下一個do的目標檔案移至下一個
        eachFile = Dir()
    Loop
'特殊字讀不到的原因:
'VBA itself supports Unicode characters, the VBA development environment does not
'FileSystemObject support Unicode characters in file names
'資料來源：
'https://stackoverflow.com/questions/33685990/working-with-unicode-file-names-in-vba-using-dir-filesystemobject-etc

'定義As New FileSystemObject, As Folder, As File要設定引用項目:
'Microsoft Scripting Runtime
'資料來源：
'https://trumpexcel.com/vba-filesystemobject/
End Function
Function 取得一路徑底下所有特定副檔名的檔案()
'取得路徑中的照片檔
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(lastOpenFolderPath)
    i = 0
    For Each objFile In objFolder.Files
        '附檔名是bmp,jpg才處理
        tempStr1 = Right(objFile.Name, Len(objFile.Name) - InStrRev(objFile.Name, "."))
        If (LCase(tempStr1) = "jpg" Or LCase(tempStr1) = "bmp") Then
            'do somethig
        End If
    Next
End Function
Function getAllFolderUnderThisPath_遞歸取得一資料夾底下各層的子資料夾(folderPath, myArray)
'使用遞歸方式，將folderPath底下所有階層的資料夾路徑都找出,譬如
'   1.folderPath底下的全部資料夾 > 抓取路徑到myArray
'   2.folderPath底下的某資料夾裡的全部資料夾 > 抓取路徑到myArray
'   3.folderPath底下的某資料夾裡的某資料夾裡的全部資料夾 > 抓取路徑到myArray
'   4. ...
'myArray必須是1維陣列
'   新資料會append到舊資料上
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

Sub 檢查資料夾是否存在_資料庫()
Dim folderFullPath As String
folderFullPath = "c:\123"

    If Dir(folderFullPath, vbDirectory) = vbNullString Then
        'do something
    End If
    
End Sub
Sub 檢查檔案是否存在_資料庫()
Dim fileFullPath As String
fileFullPath = "C:\入場人員系統\02文字資料\入場人員清單.xlsm"
    If Dir(fileFullPath) = Empty Then
         MsgBox "文件不存在。"
    End If
End Sub
Sub 取得檔名()
Dim fs, fos, fd, fc, aaa, bbb
Set fos = CreateObject("Scripting.FileSystemObject")
Set fd = fos.getfolder("C:\入場人員系統\01圖片資料\X120251959_陳志強") '檔案目錄
Set fc = fd.Files

For Each fs In fc
    ThisWorkbook.Sheets("test").Cells(1, 1) = fs.Name
    aaa = ThisWorkbook.Sheets("test").Cells(1, 1)
Next
End Sub
Sub 關於文字檔新增或讀取()
'詳 關於文字檔操作.bas
End Sub

'20231114 完成此段起測試後後，放入程式語法
Function 資料庫_checkIfFileNameContainUnacceptableCharacter(fileName As String) As Boolean
'檔名不可使用這些字元符號 \ / : * ? "" < > |
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
    '防呆-如果輸入值是空白則略過
    '   此功能僅檢查是否包含不可使用的字元符號，不做其他檢查
    If (Trim(fileName <> "")) Then
        For i = 1 To UBound(myArr, 1)
            If (InStr(1, fileName, myArr(i)) <> 0) Then
                checkIfFileNameContainUnacceptableCharacter = True
                Exit For
            End If
        Next i
    End If
    
'    If (errString <> "") Then
'        msgTitle = "錯誤            "    ' 定義標題。
'        msgText = "輸入的編號有以下錯誤，請修改" + vbLf    ' 定義訊息。
'        msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
'        msgText = msgText + iStr2 + vbLf  ' 定義訊息
'        msgStyle = vbCritical '顯示"X"圖案
'        MsgBox msgText, msgStyle, msgTitle
'    End If
End Function





Sub old_取得單一檔案路徑_資料庫()
'*如果同一段程式中使用兩個此方式，且槽區不同，則第二個使用此法的會失效-->改用FileDialog
Dim weldFilePathList As String
    '指定預設路徑
    ChDir Left("d:\123.csv", InStrRev("d:\123.csv", "\"))
    weldFilePathList = Application.GetOpenFilename(fileFilter:="(*.csv), *.csv", Title:="選擇WeldData檔 (不可複選)", MultiSelect:=False)
    If weldFilePathList = "False" Then
        GoTo 991
    End If
End Sub
Sub old_取得複數檔案路徑_資料庫()
'*如果同一段程式中使用兩個此方式，且槽區不同，則第二個使用此法的會失效-->改用FileDialog
Dim FilePathList() As Variant
    '指定預設路徑
    ChDir Left("d:\123.dwg", InStrRev("d:\123.dwg", "\"))
    On Error GoTo 991
        '同時篩選出dwg,dxf
        FilePathList = Application.GetOpenFilename(fileFilter:="(*.dwg;*.dxf), *.dwg;*.dxf", Title:="選擇要出報表的cad檔 (可複選)", MultiSelect:=True)
        '須在下拉式選單選擇要開啟dwg或者dxf
        FilePathList = Application.GetOpenFilename(fileFilter:="(*.dwg), *.dwg,(*.dxf), *.dxf", Title:="選擇要出報表的cad檔 (可複選)", MultiSelect:=True)
    On Error GoTo 0

End Sub

