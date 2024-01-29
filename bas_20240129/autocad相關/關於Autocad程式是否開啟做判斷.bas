Attribute VB_Name = "關於Autocad程式是否開啟做判斷"
Option Explicit

Private Sub checkAutoCad()
Dim obj As Object
  
    On Error Resume Next
    Set obj = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "請先執行AutoCAD。", vbExclamation, "未執行AutoCAD"
        allStopRun = True
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub checkAutocadExecuted()
Dim obj As Object
    stopRun = False
    On Error Resume Next
    Set obj = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "請先開啟AutoCAD程式" + vbLf
        msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
        msgText = msgText + "執行過程需使用到AutoCAD，故請手動開啟後再度執行此程式"   ' 定義訊息
        msgStyle = vbExclamation '顯示"!"圖案
    
        MsgBox msgText, msgStyle, msgTitle
        
        stopRun = True
        Exit Sub
    End If
    On Error GoTo 0
End Sub
