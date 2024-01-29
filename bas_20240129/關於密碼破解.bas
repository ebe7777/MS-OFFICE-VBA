Attribute VB_Name = "關於密碼破解"
Sub 工作表密碼破解1_資料庫()

answer = MsgBox("即將進行破解，是否繼續？", vbOKCancel, "邏輯破解法")

If answer = 2 Then
    MsgBox "取消破解"
    Exit Sub
End If

StartTime = Time '最後顯示執行時間用

Dim I1 As Integer, I2 As Integer, I3 As Integer, I4 As Integer, I5 As Integer, I6 As Integer
Dim I7 As Integer, I8 As Integer, I9 As Integer, I10 As Integer, I11 As Integer, I12 As Integer

On Error Resume Next
    
ProgressBar.Show 0
ProgressBar.ProgressBar_1.Min = 0
ProgressBar.ProgressBar_1.Max = 100

For I1 = 65 To 66: For I2 = 65 To 66: For I3 = 65 To 66: For I4 = 65 To 66: For I5 = 65 To 66: For I6 = 65 To 66
For I7 = 65 To 66: For I8 = 65 To 66: For I9 = 65 To 66: For I10 = 65 To 66: For I11 = 65 To 66: For I12 = 32 To 126

    ActiveSheet.Unprotect Chr(I1) & Chr(I2) & Chr(I3) & Chr(I4) & Chr(I5) & Chr(I6) & Chr(I7) & Chr(I8) & Chr(I9) & Chr(I10) & Chr(I11) & Chr(I12)
    
    ProCount = ProCount + 1
   
    If ActiveSheet.ProtectContents = False Then
        
        If ProCount < 1000 Then
            ProgressBar.Label.Caption = "本次破解共嘗試了 " & ProCount & " 組密碼。"
        Else
            ProgressBar.Label.Caption = "本次破解共嘗試了 " & Format(ProCount, "0,000") & " 組密碼。"
        End If
        
        FinalTime = Time '最後顯示執行時間用
        InputBox "已取消保護，使用之破解碼如下(歷時 " & Minute(FinalTime - StartTime) & " 分 " & Second(FinalTime - StartTime) & " 秒)：" & vbNewLine & vbNewLine & "注意！此破解碼非使用者之原始密碼，但兩者之密碼邏輯認定相同。一旦以破解碼進行反向加密，原使用者亦可輸入原始密碼來取消保護。", "邏輯破解法", Chr(I1) & Chr(I2) & Chr(I3) & Chr(I4) & Chr(I5) & Chr(I6) & Chr(I7) & Chr(I8) & Chr(I9) & Chr(I10) & Chr(I11) & Chr(I12)
        Unload ProgressBar '用來關閉進度表
        Exit Sub
    End If
    
    Call Bar(ProCount, 194560)

Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next

End Sub
Function Bar(ByVal STEP, TOTAL) '變數Step為目前執行的步驟，變數Row為總步驟"

ProgressBar.ProgressBar_1.Value = Round(STEP / TOTAL * 100, 0)

If (STEP / TOTAL * 100) > 20 Then
    ProgressBar.Label.Caption = "超強的密碼！極限是多少呢？"
ElseIf (STEP / TOTAL * 100) > 10 Then
    ProgressBar.Label.Caption = "不錯喔，是個好密碼。"
ElseIf (STEP / TOTAL * 100) > 0 Then
    ProgressBar.Label.Caption = "試試看能不能撐過10%？"
End If

ProgressBar.MessageBox_1.Value = "目前進度：" & Format(Round(STEP / TOTAL * 100, 2), "0.00") & "%"
DoEvents '用來顯示進度百分比

End Function




Sub 工作表密碼破解2_資料庫()
Dim mySht As Variant
For Each mySht In ActiveWorkbook.Worksheets
    mySht.Protect DrawingObjects:=True, CONTENTS:=True, AllowFiltering:=True
    
    mySht.Protect DrawingObjects:=False, CONTENTS:=True, AllowFiltering:=True
    
    mySht.Unprotect
Next
MsgBox "所有工作表的 [保護工作表] 已解除"

End Sub

Sub 工作表密碼破解3_資料庫()
'免VBA方式
'   將檔名改成.zip
'   不要解壓縮，直接以壓縮軟體開啟
'       *以下以7zip為例
'   進入資料夾 xl
'   進入資料夾 worksheets
'   編輯當中每個sheet的xml (點選該xml > 按滑鼠右鍵 >編輯)
'       找到 "<sheetProtection" 開頭的段落，並從  "<" 到該段 "/>" 結尾整段刪除
'       每個檔案改完後要儲存，程式會提醒你 "是否要在壓縮黨內更新檔案?" 選擇 "確定"
'   將zip檔改回原始副檔名
End Sub

Sub 活頁簿密碼破解1_資料庫()
'免VBA方式
'   將檔名改成.zip
'   開啟 (檔名)\xl\ > 找到workbook.xml
'   編輯該檔 > 找到 <workbookProtection .... /> 並將之刪除
'   將zip檔改回原始副檔名
End Sub
