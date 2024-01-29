Attribute VB_Name = "關於WORD用excel操作"
Sub 新增word檔()
Dim objWordApp As Object, objWordDoc As Object
    'Create the object of Microsoft Word
    Set objWordApp = CreateObject("Word.Application")
    'Add documents to the Word
    Set objWordDoc = objWordApp.Documents.Add
    'Make the MS Word Visible
    objWordApp.Visible = True

End Sub
Sub 開啟word檔()
'20210629
Dim objWordApp As Object, objWordDoc As Object
        Set objWordApp = CreateObject("Word.Application")
'===============
'是否顯示word程式-測試時需開啟，否則遇到錯誤時word程式會在背景開保持啟該word檔導致檔案被鎖定
'wordApp.Visible = True
wordApp.Visible = False
'===============
        objWordApp.ScreenUpdating = False
        '開啟WORD檔,唯讀
        Set objWordDoc = objWordApp.Documents.Open(fileName:=wordSampleFileFullPath, ReadOnly:=True)
End Sub


Sub 存檔關閉word程式()
Dim saveFullPath As String
        
    '另存新檔,另存為預設格式 (saveFullPath包含路徑與檔名，但可省略副檔名)
    objWordDc.SaveAs fileName:=saveFullPath, FileFormat:=wdFormatDocumentDefault
    '另存新檔,另存為docx格式 (saveFullPath包含路徑與檔名，但可省略副檔名)
    objWordDc.SaveAs fileName:=saveFullPath, FileFormat:=wdFormatXMLDocument
    '存檔
    objWordApp.Documents(saveFullPath).Save
    '關閉WORD檔
    objWordApp.Documents(saveFullPath).Close
    '並關閉WORD程式
    objWordApp.Quit
End Sub
Sub word視窗的顯示()
    '將word視窗放大最大
    objWordDoc.ActiveWindow.WindowState = wdWindowStateMaximize
    'focus到word視窗
    objWordApp.Activate
    '視窗更新畫面
    objWordApp.ScreenUpdating = False
End Sub
Sub 文繞圖方式設定()
'wdWrapInline    7   圖形與文字排列。
'wdWrapNone      3   會將圖形置於文字之前。 另請參閱wdWrapFront。
'wdWrapSquare    0   文字環繞圖形換行。 行接續是在圖形的相反一側。
'wdWrapThrough   2   文字環繞圖形換行。
'wdWrapTight     1   文字在圖形旁緊密換行。
'wdWrapTopBottom 4   將文字置於圖形之上與之下。
'wdWrapBehind    5   將圖形置於文字後方。
'wdWrapFront     6   將圖形置於文字前方。
End Sub
Sub word版面格式()
    '上下左右邊界
    ' 英吋InchesToPoints
    ' 公分CentimetersToPoints
    objWordDoc.PageSetup.TopMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.BottomMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.RightMargin = CentimetersToPoints(1)
    objWordDoc.PageSetup.LeftMargin = CentimetersToPoints(1)
    
End Sub

Sub 表格_table的操縱1()
Dim objWordRange As Object, objWordTab As Object
Dim wordTableColumns As Integer, wordTableRows As Integer
Dim i As Integer, ii As Integer, iii As Integer
'設定table row,colum數量
wordTableRows = 2
wordTableColumns = 3
    'Create a Range object.
    Set objWordRange = objWordDoc.Range
    'Create Table using Range object and define no of rows and columns.
    objWordDoc.Tables.Add Range:=objWordRange, NumRows:=wordTableRows, NumColumns:=wordTableColumns
    'Get the Table object
    Set objWordTab = objWordTab.Tables(1)
    With objWordTab
        '所有欄的文字
        .Range.Font.Size = 25
        '所有欄的寬
        ' 英吋InchesToPoints
        ' 公分CentimetersToPoints
        .PreferredWidth = CentimetersToPoints(10)
        With objWordTab
        '使格線可視(使用基本線條)
        .Borders.Enable = True
        '設定格線形式
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        For i = 1 To wordTableRows
            For ii = 1 To wordTableColumns
                '填入資料
                iii = iii + 1
                iStr = dataArry(iii, 1) & Chr(10) & dataArry(iii, 2) & Chr(10) & dataArry(iii, 3)
                .Cell(i, ii).Range.Text = iStr
                '每列高
                .Rows(i).Height = 300
                '每列垂直align
                .Rows(i).Cells.VerticalAlignment = wdAlignVerticalCenter
                '每列中的每欄水平align
                .Rows(i).Cells(ii).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Next ii
        Next i
    End With
    

End Sub
Sub 表格的操作2_每格分別替換文字()
Dim objWordCell As Object
    '設定指定cell
    Set objWordCell = objWordDoc.Range(Start:=.Tables(1).Cell(i, ii).Range.Start, _
        End:=.Tables(1).Cell(i, ii).Range.End)
        '附加說明-選取指定cell
        'objWordCell.Select
    '修改指定cell的指定字大小-以替換的方式
    With objWordCell.Find
        .ClearFormatting
        .ClearAllFuzzyOptions
        '大小寫須符合
        .MatchCase = True
        '全字拼寫須符合
        .MatchWholeWord = True
        '找到此文字
        .Text = "test"
        '替換成此文字
        .Replacement.Text = "123"
        .Replacement.Font.Size = 20
        .Execute Format:=True, Replace:=wdReplaceAll
    End With
End Sub
Sub 替換文字與修改字大小()
    With objWordDoc.Content.Find
        .ClearFormatting
        .ClearAllFuzzyOptions
        '大小寫須符合
        .MatchCase = True
        '全字拼寫須符合
        .MatchWholeWord = True
        '找到此文字
        .Text = "ABC"
        '替換成此文字
        With .Replacement
            .Text = "123"  'this line might be unnecessary
            .Font.Size = 20
        End With
        
        '使用early binding
        .Execute Format:=True, Replace:=wdReplaceAll
        '使用late binding
        .Execute Format:=True, Replace:=2
        '說明 https://social.msdn.microsoft.com/Forums/en-US/452d6d5f-e485-41a2-967c-27e739cc8d9e/hard-to-solve-ms-word-late-binding-problem?forum=isvvba
        'Replace:=2 代表意思 https://docs.microsoft.com/en-us/office/vba/api/word.wdreplace
    End With
End Sub

Sub 之前寫的_讀取既有檔案另存後將裡面的圖片替換掉()
'須設定 工具>設定引用項目>Microsoft Word xx.x Object Library
Dim objWordApp As Object, objWordDoc As Object
Dim fileFullPath As String, fileSavePath As String, fileSaveName As String
Dim saveFullPath As String
Dim originalImage As InlineShape, newImage As InlineShape
Dim imageControl As ContentControl
Dim imageW As Long, imageH As Long
Dim newImagePath As String

fileFullPath = "D:\Bruce工作資料\進場申請書產生程式\測試用資料\日月光-人員入廠資料雇主切結書(SAMPLE)_已設定替換文字與圖片.docx"
fileSavePath = "D:\Bruce工作資料\進場申請書產生程式\測試用資料"
fileSaveName = "test.docx"
saveFullPath = fileSavePath & "\" & fileSaveName
    Set objWordApp = New Word.Application
    objWordApp.Visible = True
    '開啟WORD檔,唯讀
    Set objWordDoc = objWordApp.Documents.Open(fileName:=fileFullPath, ReadOnly:=True)
    '另存新檔
    objWordDoc.SaveAs fileName:=saveFullPath
    '由於原本是唯讀狀態開啟，導致另存後無法編輯，故在另存後關掉重開產生的新檔
    objWordApp.Documents(saveFullPath).Close
    Set objWordDoc = objWordApp.Documents.Open(fileName:=saveFullPath, ReadOnly:=False)
    '取代文字
    With objWordDoc.Content.Find
        .Text = "帆宣系統科技股份有限公司"
        .Replacement.Text = "TEST"
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    
    '取代InlineShape圖案

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

    
    newImagePath = "D:\Bruce工作資料\進場申請書產生程式\測試用資料\這個是共用資料資料夾\少變動圖檔\大頭照\C123456789_林韋成_P.jpg"
    objWordDoc.InlineShapes.AddPicture newImagePath, False, True, imageControl.Range

    With imageControl.Range.InlineShapes(1)
        .Height = imageH
        .width = imageW
    End With
    
    
    
    '存檔並關閉WORD檔
    objWordApp.Documents(saveFullPath).Save
    objWordApp.Documents(saveFullPath).Close
    
    objWordApp.Quit
    '關閉警告
    objWordApp.DisplayAlerts = True
End Sub
