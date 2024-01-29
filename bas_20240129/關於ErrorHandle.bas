Attribute VB_Name = "關於ErrorHandle"
'新手問題:VBA當遇到錯誤且使用者有下On Error時，會進入使用者自訂的錯誤處理模式，並且此錯誤尚未處理完不會再度進入On Error
'  舉例: 用for對與 "使用者指定名稱"同名的工作表做處理，並使用on error來防止找不到同名的工作表
'           當第一次遇到問題時，會進入on error指定的處理方式，但再度遇到時就會跳出系統錯誤訊息而非進入on error
'原因:在on error後使用者須要使用以下任一個方式來告知VBA此錯誤已經處理完畢
'   Resume Next ' go to the line following error
'   Resume ' go back to the same line of code
'   Exit Sub ' go out of this routine
Sub 正確errorHandle範例()
    For i = 1 To 10
        haveErr = False
        '防呆-該工作表不存在
        On Error GoTo 880
        '可能會錯誤處 - 如果工作表不存在進入880標記有問題
        Set nowSht = Sheets(Cells(i, 1))
        '將on error改成不處理-否則接下來遇到的任何錯誤都會以goto 880處理
        On Error GoTo 0
        '沒問題才執行動作
        If (haveErr = False) Then
           'when no error,do something
        End If
        GoTo 881
880
        '標記有問題
        haveErr = True
        '告訴程式剛剛的錯誤已經處理完畢
        Resume Next
881
    Next i
End Sub
Sub 關於991與999_資料庫()
Dim msgTitle As String, msgText As String, msgStyle As String

991
    msgTitle = "警告            "   ' 定義標題。
    msgText = "  請重新執行此程式 !"
    MsgBox msgText, vbExclamation, msgTitle
    Exit Sub

999
'告知user執行完畢
    msgTitle = "訊息            "    ' 定義標題。
    msgText = "  程式執行完畢 !"
    MsgBox msgText, vbOKOnly, msgTitle
    

End Sub

