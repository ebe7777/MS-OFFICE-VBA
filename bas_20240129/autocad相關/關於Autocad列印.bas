Attribute VB_Name = "關於Autocad列印"


Private Sub 列印成pdf_資料庫()
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
Dim myFileDialog As FileDialog
Dim CanonicalMediaNameArray As Variant
Dim pdfSavePath As String
Dim iStr As String
Dim iBool As Boolean
'vvvvvv前置動作vvvvvv
    '確定autoCad程式有開啟
    Set acad = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "請先開啟AutoCAD程式"
        Exit Sub
    End If
    '取得存放路徑
    Set myFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With myFileDialog
        .Title = "選擇檔案存放資料夾"
        .AllowMultiSelect = False
    End With
    myFileDialog.Show
    If myFileDialog.SelectedItems.Count = 0 Then
        Exit Sub
    Else
        pdfSavePath = myFileDialog.SelectedItems.Item(1)
    End If
    '開檔
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    sdi = dwgFile.GetVariable("sdi")
    dwgFile.SetVariable "sdi", 0
    '設定 不要 背景列印-這樣程式才會等前一張出完才繼續執行出第二章的動作 (否則前一張還沒出完程式馬上送下一張的列印指令會導致報錯)
    If (dwgFile.GetVariable("BACKGROUNDPLOT") <> 0) Then
        dwgFile.SetVariable "BACKGROUNDPLOT", 0
    End If
'vvvvvv從配置layout出圖vvvvvv
    '取得layout名稱
    Dim myLayouts As AcadLayout
    For Each myLayouts In dwgFile.Layouts
        iStr = myLayouts.name
    Next
    
    
    
    
'vvvvvv從模型Model出圖vvvvvvv
    Dim blockName As String, blockWid As Double, blockHei As Double
    Dim blockXScaleFactor As Double, blockYScaleFactor As Double
    Dim obj As Object
    Dim blockMinPoint As Variant, blockMaxPoint As Variant
    Dim blockXScaleFactor As Double, blockYScaleFactor As Double
    '使用者須手動輸入圖塊名稱、寬、高(目前查無自動判斷的方式)
    blockName = "FRAM"
    blockWid = 123
    blockHei = 456
    '取得模型中所有使用該圖塊為圖框的物件的左下角和右上角
    Cells(1, 5) = "block左下角X"
    Cells(1, 6) = "block左下角Y"
    Cells(1, 7) = "block X 比例"
    Cells(1, 8) = "block Y 比例"
    Cells(1, 9) = "block右上角X=左下角X+(圖塊寬*X比例)"
    Cells(1, 10) = "block右上角Y=左下角Y+(圖塊高*Y比例)"
    Cells(1, 11) = "出圖結果"
    i = 1
    For Each obj In dwgFile.ModelSpace
        If TypeOf obj Is AcadBlockReference Then
            If (obj.name = blockName) Then
                i = i + 1
                '只取用blockMinPoint -> 圖塊左下角座標
                ' [0]X座標 [1]Y座標
                obj.GetBoundingBox blockMinPoint, blockMaxPoint
                '圖塊比例
                blockXScaleFactor = obj.XScaleFactor
                blockYScaleFactor = obj.YScaleFactor
                '將取得資訊寫到excel上
                Cells(i, 5) = Round(blockMinPoint(0), 4)
                Cells(i, 6) = Round(blockMinPoint(1), 4)
                Cells(i, 7) = Round(blockXScaleFactor, 4)
                Cells(i, 8) = Round(blockYScaleFactor, 4)
                '計算出右上角座標並寫到excel
                Cells(i, 9) = Round(blockMinPoint(0) + (blockXScaleFactor * blockWid), 4)
                Cells(i, 10) = Round(blockMinPoint(1) + (blockYScaleFactor * blockHei), 4)
            End If
        End If
    Next
    '如果用窗選出圖，需要調整DVIEW
    '   模型空間+使用窗選出圖可能遇到此問題 : 程式抓好左下右上點，但使用acWindow時程式指定的窗選位置卻不在點上
    '       https://forums.autodesk.com/t5/visual-basic-customization/window-plot-using-vba-strange-offset/m-p/9633693#M104077
    '   設定DVIEW就可以解決此問題
    '   https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Plot-or-preview-shows-incorrect-area-of-drawing-when-plotting-limits.html
    '   DVIEW原理
    '       https://knowledge.autodesk.com/support/autocad/learn-explore/caas/CloudHelp/cloudhelp/2019/ENU/AutoCAD-Core/files/GUID-E0078D09-8449-4A0A-A5AD-6984A01CEC33-htm.html
    '   DVIEW command詳細說明
    '       https://help.bricsys.com/hc/en-us/articles/360006566734-DView
    
    '   設定DVIEW-目前找不到用vba的方式，故只能用傳送command的方式
    dwgFile.SendCommand "DVIEW" & vbCr & "ALL" & vbCr & vbCr & "POINT" & vbCr & "0,0,0" & vbCr & "0,0,1" & vbCr & vbCr
    '   zoom回物件
    ZoomAll
    
    

'vvvvvvu列印通用設定vvvvvv
    '關掉背景執行，否則第一張圖還沒執行完程式就嘗試出第二張會導致無法出圖;手冊上的說明如下
    'To plot in the foreground using VBA, you must set the BACKGROUNDPLOT system variable to 0. Otherwise, plotting occurs in the background.
    cadBackPlot = dwgFile.GetVariable("BACKGROUNDPLOT")
    acad.ActiveDocument.SetVariable "BACKGROUNDPLOT", 0
    '   印表機/繪圖機>名稱
    dwgFile.ActiveLayout.ConfigName = "DWG To PDF.pc3"
    '   圖紙大小
    dwgFile.ActiveLayout.CanonicalMediaName = "ISO_expand_A3_(297.00_x_420.00_MM)"
    
    '   出圖範圍>出圖內容/plot area(所得值為integer)
    '       0   acDisplay
    '       1   acExtents
    '       2   acLimits
    '       3   acView
    '       4   acWindow
    '       5   acLayout
    dwgFile.ActiveLayout.plotType = acExtents
    '       使用acWindow
    dwgFile.ModelSpace.Layout.plotType = acWindow
    '           窗選出圖範圍 (0)x水平(1)y垂直
    Dim leftDownPointArray(1) As Double, rightUpPointArray(1) As Double
    leftDownPointArray(0) = 0
    leftDownPointArray(1) = 0
    rightUpPointArray(0) = 200
    rightUpPointArray(1) = 100
    dwgFile.ModelSpace.Layout.SetWindowToPlot leftDownPointArray, rightUpPointArray
    
    
    '   出圖偏移量>置中出圖
    '       注意!plotType不為acLayout才可設定此，否則會報錯
    dwgFile.ActiveLayout.CenterPlot = True
    '   出圖比例>比例
    '       比例-單位
    dwgFile.ActiveLayout.PaperUnits = acMillimeters
    '       比例-自訂 (和 下拉式選單 只能二選一)
    dwgFile.ActiveLayout.SetCustomScale 1, 2
    '       比例-下拉式選單可選到的 (和 自訂 只能二選一)
    '        dwgFile.ActiveLayout.StandardScale = ac1_2
    '       圖面方位 --> 用旋轉角度來對應直式/橫式/上下顛倒
    dwgFile.ActiveLayout.plotRotation = ac90degrees
    '       出圖形式表
    dwgFile.ActiveLayout.StyleSheet = "1880-A3-thin.ctb"
    
    '   將背景執行狀態設為此圖的原始值
    acad.ActiveDocument.SetVariable "BACKGROUNDPLOT", cadBackPlot
    
    dwgFile.Application.ZoomExtents

'vvvvvv出圖到檔案vvvvvv
    
    '   防呆-使用者如果選擇某硬碟而非某資料夾
    If (Right(pdfSavePath, 1) = "\") Then
        iBool = dwgFile.Plot.PlotToFile(pdfSavePath & "myPdf.pdf")
    Else
        iBool = dwgFile.Plot.PlotToFile(pdfSavePath & "\" & "myPdf.pdf")
    End If
    '將 是否出圖成功 寫到excel上
    If (iBool = True) Then
        Cells(i, 11) = "出圖成功"
    Else
        Cells(i, 11) = "出圖失敗"
    End If
    
    
    
    '** 建議關閉 出圖後自動打開pdf檔 以加快出圖速度 **
    '打開或關閉 "出圖後自動打開pdf檔" 的地方在cad上:
    '左上 "A"圖示 > 列印 > 印表機/繪圖機 當中 > 名稱 選擇"DWG To PDF.pc3" >PDF選項 打開，將 "在檢視器中顯示結果" 取消勾選

End Sub

Private Sub 取得CanonicalMediaName_資料庫()
Dim acad As AcadApplication, dwgFile As AcadDocument, sdi As String
Dim CanonicalMediaNameArray As Variant

Set acad = GetObject(, "AutoCAD.Application")

'開檔
Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
sdi = dwgFile.GetVariable("sdi")
dwgFile.SetVariable "sdi", 0
'設定 印表機/繪圖機>名稱
dwgFile.ActiveLayout.ConfigName = "DWG To PDF.pc3"

'取得[印表機/繪圖機>名稱]所有可使用的CanonicalMediaName
CanonicalMediaNameArray = dwgFile.ActiveLayout.GetCanonicalMediaNames()

End Sub

Private Sub 使用印表機列印各檔中的所有配置()
'搭配functionSelectPrinter.bas/FormSelectPrinters.frm使用
'   匯入functionSelectPrinter.bas，讓他自成一個module
'   匯入FormSelectPrinters.frm
'需宣告此全域變數，讓SUB/FORM可以溝通
'   Public printerName As String
'使用以下程式碼
Dim obj As Object
Dim acad As AcadApplication, dwgFile As AcadDocument, myLayout As AcadLayout
Dim i As Long, ii As Long, iii As Long
    '防呆 - 檢查autocad是否已被執行
    On Error Resume Next
        Set obj = GetObject(, "AutoCAD.Application")
        If Err Then
            msgTitle = "注意"
            msgText = "請先開啟AutoCAD程式" + vbLf
            msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
            msgText = msgText + "執行過程需使用到AutoCAD，故請手動開啟後再度執行此程式"   ' 定義訊息
            msgStyle = vbExclamation '顯示"!"圖案
        
            MsgBox msgText, msgStyle, msgTitle
            
            stopRun = True
            Exit Sub
        End If
    On Error GoTo 0
    '選擇印表機
    Form99Printers.Show
    If (stopRun = True) Then
        GoTo 991
    End If
    
    Set acad = GetObject(, "AutoCAD.Application")
    acad.Visible = True
    '開檔
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    '設置sdi為0
    dwgFile.SetVariable "sdi", 0
    '設置 不要 背景列印
    dwgFile.SetVariable "BACKGROUNDPLOT", 0
    
    ii = 2
    iii = 0
    For Each myLayout In dwgFile.Layouts
        '紀錄配置名稱
        If (myLayout.name <> "Model") Then
            ii = ii + 1
            iii = iii + 1
            autocadToolSht.Cells(i, ii) = myLayout.name
            '設置列印範圍
            dwgFile.ActiveLayout = dwgFile.Layouts(myLayout.name)
            '設置印表機
            dwgFile.ActiveLayout.ConfigName = printerName
            '執行列印
            dwgFile.Plot.PlotToDevice
        End If
    Next

    dwgFile.Close False
    acad.Visible = True
991
End Sub
