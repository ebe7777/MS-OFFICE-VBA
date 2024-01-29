Attribute VB_Name = "關於文字檔操作"
Sub 讀取文字檔並切字1_只能讀ansi()

Dim linesArray
Dim splitArray

'方式2才需dim以下
Dim readData As String
Dim i As Long

'開啟檔案(此例為只開啟一個檔案)
'#1 file number
'   Number used in the Open statement to open a file. Use file numbers in the range 1-255, inclusive, for files not accessible to other applications. Use file numbers in the range 256-511 for files accessible from other applications.
Open weldFilePathList For Input As #1
'將內容一行行拆進array裡
'==============
'方式1 - 可能會有問題
'LOF代表讀取裡面的字數(各行相加)，但pdms出的bom不知為何vba得到的字數比檔案實際字數多故出現錯誤62
linesArray = split(Input$(LOF(1), #1), vbNewLine)

'方式2 - 可避免掉方式1的問題
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

'關閉檔案
Close #1
'一行行內容以"^"拆開後寫進Array
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

Sub 讀取文字檔2_能讀unicode與ansi()
Dim fso As Object
Dim myTxtFile As Object
Dim i As Long
Set fso = CreateObject("Scripting.FileSystemObject")
'不知為何這邊iomode/Format不能使用constant，必須使用數字
'iomode:= (1)唯讀 (2)overwrite (3)append
'Create:= 指定的fileName如果不存在 (true)建新檔(false)不建新檔
'Format:=  (0)開啟檔是ansi (-1)開啟檔是Unicode (-2)使用系統愈預設值開啟檔
Set myTxtFile = fso.OpenTextFile(fileName:="c:\123.txt", iomode:=1, Create:=False, Format:=-1)
'讀取直到該檔最後一行
Do Until myTxtFile.AtEndOfStream
    i = i + 1
    '讀取檔案中每列資料
    Sheets("123").Range("A" & i) = myTxtFile.ReadLine
Loop
End Sub
Sub 輸出資料到文字檔1_只支援ansi格式()

Dim filePath As String
Dim SDTEROWS As Long, SMTEROWS As Long, ROWS As Long
Dim SDTECOLUMNSTRING As String, SMTECOLUMNSTRING As String
'==================================================================================
'
'輸出資料到文字檔
'----------------------------------------------------------------------------------
filePath = "D:\SXTE.TXT"
SDTEROWS = Sheets(SDTE).Range("A65536").End(xlUp).Row
SMTEROWS = Sheets(SMTE).Range("A65536").End(xlUp).Row
SDTECOLUMNSTRING = "E"
SMTECOLUMNSTRING = "D"
'在此會自動新增檔案
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
Sub 輸出資料到文字檔2_支援ansi與unicode格式()
Dim fso As Object
Dim myTxtFile As Object
Dim i As Integer
    '新增一個文字檔
    Set fso = CreateObject("Scripting.FileSystemObject")
    '   如果是ansi,最後的Unicode:=false
    Set myTxtFile = fso.CreateTextFile(fileName:="c:\123.txt", OverWrite:=True, Unicode:=True)
    '寫入資料
    For i = 1 To 5
        myTxtFile.Write i & vbCrLf
    Next i
    '關檔
    myTxtFile.Close
    '另一個新增unicode文字檔的方式是 寫入excel中>excel另存為unicode格式txt
    'ActiveWorkbook.SaveAs fileName:="D:\123.txt", FileFormat:=xlUnicodeText, CreateBackup:=False
End Sub
