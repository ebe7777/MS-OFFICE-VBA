Attribute VB_Name = "關於MSGBOX"
Sub MSGBOX內容_資料庫()
Dim msgTitle As String, msgText As String, msgStyle As String

'開場白
    msgTitle = "有錯誤產生            "    ' 定義標題。

    msgText = "說明如下:" + vbLf   ' 定義訊息。
    msgText = msgText + "發現有部分SUP TYPE與工作表 [" & sNType & "] 中資的資料不符。" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "請執行以下動作：" + vbLf   ' 定義訊息
    msgText = msgText + "(1)參照工作表 [" & sNPdmsReport & "] 中L欄與Q欄標示 ERROR 處並修改資料理。" + vbLf  ' 定義訊息
    msgText = msgText + "(2)再次執行此程式。"    ' 定義訊息
        
'    msgStyle = vbOKOnly '顯示確定
'    msgStyle = vbOKCancel '顯示確定/取消
'    msgStyle = vbYesNo '顯示是/否
'    msgStyle = vbYesNoCancel '顯示是/否/取消
'    msgStyle = vbCritical '顯示"X"圖案
'    msgStyle = vbQuestion '顯示"?"圖案
'    msgStyle = vbExclamation '顯示"!"圖案
'    msgStyle = vbInformation '顯示"i"圖案
    
    MsgBox msgText, msgStyle, msgTitle
End Sub
Sub 執行完畢()
    msgTitle = "訊息" ' 定義標題。
    msgText = "執行完畢"
    msgStyle = vbInformation '顯示"i"圖案
    MsgBox msgText, msgStyle, msgTitle
End Sub
Sub VBOKCANCEL寫法_資料庫()
Dim msgTitle As String, msgText As String, msgStyle As String, answer As Variant

msgTitle = "重要訊息            "    ' 定義標題。
msgText = "使用本程式有以下條件限制 :" + vbLf + vbCrLf + vbLf  ' 定義訊息。
msgText = msgText + "   1.SPEC表必須符合PM格式" + vbLf + vbCrLf  ' 定義訊息。
msgText = msgText + "     (參照本檔中的工作表""SPEC格式範本"",A~K欄必須與之相同)" + vbLf + vbCrLf  ' 定義訊息
msgText = msgText + "   2.SPEC表中相同CLASS的管件需排序在一起" + vbLf + vbCrLf + vbLf  ' 定義訊息
msgText = msgText + "如果確定無誤請按[ 確定 ]繼續，或者按[ 取消 ]離開程式 !" + vbLf + vbCrLf ' 定義訊息。

answer = MsgBox(msgText, vbOKCancel + vbExclamation, msgTitle)

If answer = vbCancel Then
    Exit Sub
    Else
End If
End Sub

Sub VBYESNO寫法_資料庫()
Dim msgTitle As String, msgText As String, msgStyle As String, answer As Variant

msgTitle = "重要訊息            "    ' 定義標題。
msgText = "所產生的MTO1~MTO4總表是否要列印 ? 如果要列印請選擇 [是] " + vbLf + vbCrLf + vbLf  ' 定義訊息。
msgText = msgText + "   ==>如選擇[是]，程式會為了列印進行排版。" + vbLf + vbCrLf  ' 定義訊息。"
msgText = msgText + "       *需耗時 30秒 以上(資料量越多越久)" + vbLf + vbCrLf   ' 定義訊息


answer = MsgBox(msgText, vbYesNo + vbQuestion, msgTitle)

If answer = vbYes Then
    
    Else
End If
End Sub
