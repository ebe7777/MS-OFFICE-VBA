Attribute VB_Name = "關於Outlook用excel操作"
'基本知識
'   https://docs.microsoft.com/zh-tw/office/vba/api/outlook.mailitem
'   https://support.microsoft.com/zh-tw/topic/%E4%BD%BF%E7%94%A8-word-%E6%96%87%E4%BB%B6%E5%92%8C-excel-%E6%B4%BB%E9%A0%81%E7%B0%BF%E4%B8%AD%E4%B9%8B%E8%B3%87%E6%96%99%E5%BE%9E-outlook-%E5%82%B3%E9%80%81%E9%83%B5%E4%BB%B6%E7%9A%84-vba-%E5%B7%A8%E9%9B%86-56bbd7a9-7814-9c52-2c83-e92c01fa8418
'關閉outlook物件
'   https://learn.microsoft.com/zh-tw/office/vba/api/outlook.mailitem.close(method)
'將Word的內容完整貼上
'   https://stackoverflow.com/questions/35609112/how-to-send-a-word-document-as-body-of-an-email-with-vba
Sub 資料庫_autoEmail()
'20210630  待將outlook相關程式碼作筆記
 
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
    
    '告知使用者規則，問是否繼續
    '   姓名、編號從上到下的順序需與工作表[收件人]一致
    '   可替換的關鍵字為10個
    msgTitle = "注意            "    ' 定義標題。
    msgText = "使用此功能時，工作表 [" & mailAddressSN & "] 每一列的 姓名/編號(A~B欄)" + vbLf + vbCrLf     ' 定義訊息。
    msgText = msgText + "必需與工作表 [" & mailTitleContentSN & "] 和工作表 [" & mailAttach1SN & "] 的內容一致" + vbCrLf
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "如確定並繼續請點選[是]，或按[否]離開程式  " + vbCrLf + vbCrLf  ' 定義訊息。
    msgText = msgText + "注意!所有資料的 姓名/編號/email(A~C欄)的底色標記都會被移除"
    answer = MsgBox(msgText, vbYesNo + vbQuestion, msgTitle)
    If answer = vbNo Then
        GoTo 999
    End If
    excuteTime = Now()
    totalRows = myDataRows(ThisWorkbook.Name, mailAddressSN, "A", 65536)
    
    '防呆-沒有輸入信件位址
    If (totalRows <= titleUsedRows) Then
        MsgBox "工作表 [" & mailAddressSN & "] 裡找不到資料 (注意，每一列A~C欄都須填寫，否則會造成程式誤判)"
        GoTo 999
    End If

    '防呆檢查
    '   [收件人]的姓名、編號從上到下的順序，和工作表[郵件主旨與內文]裡面的值是否與工作表一致
    Call compareShts(mailAddressSht, mailTitleContentSht)
    If (stopRun = True) Then
        GoTo 999
    End If
    '   [收件人]的姓名、編號從上到下的順序，和工作表[附件1資訊]裡面的值是否與工作表一致
    Call compareShts(mailAddressSht, mailAddressSht)
    If (stopRun = True) Then
        GoTo 999
    End If
    '   防呆-email地址是錯誤值
    For i = (titleUsedRows + 1) To totalRows
        stopRun = False
        With mailAddressSht
            If (IsError(.Cells(i, 3)) = True) Then
                '有錯誤標示黃底
                With .Cells(i, 3).Interior
                    .Pattern = xlSolid
                    .Color = 65535
                    stopRun = True
                End With
            Else
                '沒錯誤刪除底色
                With .Cells(i, 3).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End With
    Next i
    If (stopRun = True) Then
        msgTitle = "有錯誤            "    ' 定義標題。
        msgText = "工作表 [" & mailAddressSN & "] 的email資料有問題" + vbLf + vbCrLf  ' 定義訊息。
        msgText = msgText + "已在該欄以黃底做標示" + vbLf + vbCrLf  ' 定義訊息
        msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
        msgText = msgText + "請修正後再次執行此程式" + vbLf   ' 定義訊息訊
        msgStyle = vbCritical '顯示"X"圖案
        MsgBox msgText, msgStyle, msgTitle
        GoTo 999
    End If
    
    '讓使用者選擇輸出範圍
    UserForm1.Show
    '   防呆-使用者按X離開程式
    If (stopRun = True) Then
        GoTo 999
    End If
    mailStartRow = sysSht.Range("B3").Value
    mailEndRow = sysSht.Range("D3").Value
    needAttachmentOrNot = sysSht.Range("B4").Value
    
    '蒐集郵件相關資訊
    '   將每封email所需的資訊都寫在陣列中
    ReDim mailInfoArray(totalRows, 4)
    '[n]第幾個收件者
    '[x][0]姓名_編號(人工檢查用，程式不讀取) [x][1]email地址 [x][2]email主旨 [x][3]內文檔檔名 [x][4]附件1檔名
    For i = mailStartRow To mailEndRow
        mailInfoArray(i, 0) = mailAddressSht.Cells(i, 1) & "_" & mailAddressSht.Cells(i, 2)
        mailInfoArray(i, 1) = mailAddressSht.Cells(i, 3)
        mailInfoArray(i, 2) = mailTitleContentSht.Cells(i, 4)
        mailInfoArray(i, 3) = sysSht.Range("D1") & "\" & mailTitleContentSht.Cells(i, 3) & ".docx"
        If (needAttachmentOrNot = True) Then
            mailInfoArray(i, 4) = sysSht.Range("D2") & "\" & mailAttach1Sht.Cells(i, 3) & ".docx"
        End If
    Next i
    
'ppppp開頭-呼叫視窗
ThisWorkbook.Activate
progressBarPercentNo = 0
ProgressBar.Show False
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

    iCount1 = 0
    needAttachmentOrNot = sysSht.Range("B4").Value
    For i = mailStartRow To mailEndRow

        contentFilePath = mailInfoArray(i, 3)
        '取出特定Word檔上的字，以word原始的格式
        On Error GoTo 880
            Set objWord = CreateObject("Word.Application")
            Set objWordDoc = objWord.Documents.Open(fileName:=contentFilePath, ReadOnly:=True)
            objWordDoc.Content.copy
            objWordDoc.Close
            Set objWord = Nothing
        On Error GoTo 0
        On Error GoTo 881
            '寄信
            Set objOutlook = CreateObject("Outlook.Application")
            Set mailSendItem = objOutlook.CreateItem(0)
            '   為了使用outlook的預設簽名，需顯示視窗並複製內容
            mailSendItem.Display
        On Error GoTo 0
        On Error GoTo 882

            With mailSendItem
                '主旨為空白視為錯誤
                If (mailInfoArray(i, 2) = "") Then
                    GoTo 883
                Else
                    .Subject = mailInfoArray(i, 2)
                End If
                '   late binding只能輸入bodyformat的代表號碼
                .bodyformat = 2
                '1:olFormatPlain
                '2:olFormatHTML
                '   附件會以outlook方式夾帶
                '3:olFormatRichText
                '   附件會以文字中插入物件的方式夾帶
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
            '已發出mail的，在email標黃底，D欄寫上寄信時間
            mailAddressSht.Range("C" & i).Interior.Color = 65535
            mailAddressSht.Range("D" & i).Value = excuteTime
        On Error GoTo 0

'按比例增加一定數量的進度
'ppppp-0 to 100
ThisWorkbook.Activate
'   此段執行結束共會增加iProcessRange個進度百分比/資料筆數共iDataCounts筆/每iProcessRangeGap筆從新計算一次進度
iProcessRange = 100
iDataCounts = totalRows - mailStartRow + 1
iProcessRangeGap = 1
'   計算是否要重新計算進度，目前資料是第i筆
iMod = i Mod iProcessRangeGap
'   計算共要更新幾次
iMax = CInt(iDataCounts / iProcessRangeGap)
'   要更新進度時，進度條增加 (1/iMax*iProcessRange) 的進度
If (iMod = 0) Then
    progressBarPercentNo = progressBarPercentNo + ((1 / iMax) * iProcessRange)
    Call ProgressBar.updateProgressBar(progressBarPercentNo)
End If
'ppppp

    Next i
    
    
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    GoTo 991
    
880
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "工作表 [" & mailAddressSN & "] 第" & i & "列的收件者的 內文檔 找不到" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "可能原因如下" + vbLf  ' 定義訊息
    msgText = msgText + "(1)內文檔尚未產生 --> 請執行程式[產生內文] " + vbLf  ' 定義訊息
    msgText = msgText + "(2)檔案路徑不正確 --> 請執行程式[設定] " + vbLf  ' 定義訊息
    msgText = msgText + "(3)工作表 [" & mailTitleContentSN & "] C欄的值和內文檔檔名不符 --> 請修改檔名或工作表內容 " + vbLf + vbCrLf ' 定義訊息
    msgText = msgText + "修改完成後，請從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
881
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "Outlook程式尚未開啟，或因不明原因關閉" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "請手動開啟Outlook程式，並從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
882
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "處理到一半發生未知的問題" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "請先嘗試從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
883
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp

    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "工作表 [" & mailTitleContentSN & "] 第" & i & "列的收件者的 郵件主旨 是空白" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "請修改工作表 [" & mailTitleContentSN & "] D欄，並從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
884
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "工作表 [" & mailAddressSN & "] 第" & i & "列的收件者的 附件檔 找不到" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "可能原因如下" + vbLf  ' 定義訊息
    msgText = msgText + "(1)附件檔尚未產生" + vbLf  ' 定義訊息
    msgText = msgText + "    -> 請執行程式[產生附件] " + vbLf  ' 定義訊息
    msgText = msgText + "(2)檔案路徑不正確" + vbLf  ' 定義訊息
    msgText = msgText + "    -> 請執行程式[設定] " + vbLf  ' 定義訊息
    msgText = msgText + "(3)工作表 [" & mailAddressSN & "] C欄的值和附件檔檔名不符" + vbLf ' 定義訊息
    msgText = msgText + "    -> 請修改檔名或工作表內容 " + vbLf + vbCrLf ' 定義訊息
    msgText = msgText + "修改完成後，請從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    
885
'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
    '關掉outlook視窗、不存檔
    mailSendItem.Close 1 'olDiscard
    
    msgTitle = "發現問題           "    ' 定義標題。
    msgText = "工作表 [" & mailAddressSN & "] 第" & i & "列的收件者郵件地址不正確" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "請重新產生附件檔或重新指定附件檔路徑，並從中斷處繼續執行" + vbLf  ' 定義訊息
    msgText = msgText + "如一直卡在同一個人(反覆出現此訊息)，請聯絡程式開發者"    ' 定義訊息
    msgStyle = vbExclamation '顯示"!"圖案
    MsgBox msgText, msgStyle, msgTitle
    
    GoTo 999
    


991
    msgTitle = "訊息          "    ' 定義標題。
    msgText = "處理完畢，總共寄出 " & iCount1 & " 封信" + vbLf ' 定義訊息。
    msgText = msgText + "(已將成功發出mail的在工作表 [" & mailAddressSN & "] 的C欄以黃底做標示)" ' 定義訊息。
    msgStyle = vbInformation '顯示"i"圖案
    MsgBox msgText, msgStyle, msgTitle
    GoTo 999
999
    ThisWorkbook.Activate
    mailAddressSht.Activate
End Sub
