Attribute VB_Name = "關於SHEET操作"
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Private Sub Worksheet_Change(ByVal Target As Range)
'工作表特定地方被改變，則觸發程式
    Set editRg = Target
    If (Target.address = dateSelectYearRg.address) Then
        For Each iRg In editRg
        Next
    End If
End Sub
Sub 在此檔內刪除舊表並新增此表_資料庫()
Dim mySht As Worksheet
    '強制刪除舊表
    Application.DisplayAlerts = False
        On Error Resume Next
            Sheets("樞紐分析").Delete
        On Error GoTo 0
    Application.DisplayAlerts = True
    '以複製方式產生工作表
    '   關掉警告，以免因為格式或名稱跳出一堆訊息
    Application.DisplayAlerts = False
        '複製到最後一張工作表後面
        Sheets("原始表").copy After:=Sheets(Sheets.Count)
        '複製到特定工作表後面(需用.index取得該張工作表編號)
        Sheets("原始表").copy After:=Sheets(mySht.Index)
    Application.DisplayAlerts = True
    '以新增方式產生工作表
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.Count)
    ActiveSheet.Name = "樞紐分析"
End Sub
Sub 檢查要新增的表是否存在()
Dim iStr1 As String
    iStr1 = "要新增的工作表名稱"
    For Each iVar1 In ThisWorkbook.Worksheets
        If (iVar1.Name = iStr1) Then
            ibool1 = True
        End If
    Next iVar1
    

    '           如果存在，選擇將刪除裡面的資料，或者待使用者修改工作表名稱
    If (ibool1 = True) Then
        msgTitle = "注意           "    ' 定義標題。
        msgText = "以下工作表將 [刪除] 並重新產生" + vbLf  ' 定義訊息。
        msgText = msgText + "====================================" + vbLf ' 定義訊息。
        msgText = msgText + iStr1 + vbLf
        msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
        msgText = msgText + "   ==>如果同意，請選擇[確定]；或選擇[取消]離開程式" + vbLf   ' 定義訊息。"
        msgText = msgText + "       *選擇取消者，請修改 上述工作表的名稱 ，然後再次執行" + vbLf    ' 定義訊息
        answer = MsgBox(msgText, vbOKCancel + vbExclamation, msgTitle)
        If answer = vbCancel Then
            GoTo 993
        Else
            '刪除工作表
            Application.DisplayAlerts = False
                ThisWorkbook.Worksheets(iStr1).Delete
            Application.DisplayAlerts = True
        End If
    End If
    '       新增該張工作表
    Application.DisplayAlerts = False
        ThisWorkbook.Sheets(eachMonthBaseSN).copy After:=ThisWorkbook.Sheets(Sheets.Count)
    Application.DisplayAlerts = True
    ActiveSheet.Name = iStr1
    Set eachMonthSht = ThisWorkbook.Worksheets(iStr1)

End Sub
Sub SHEET標籤上色_資料庫()
'SHEET上色:黑
Sheets("SYSTEM").Tab.Color = 6299648
'SHEET上色:黃
Sheets("Data").Tab.Color = 65535
'SHEET上色:黃
Sheets("樞紐分析").Tab.Color = 65535
'SHEET上色:紅
Sheets("ALLOCATION LIST").Tab.Color = 255
End Sub
Sub 不使用activesheet_新增sheet_資料庫()
Dim myWB As Workbook, mySht As Worksheet
Dim mySN As String
'注意,需先set myWB
    Set myWB = ThisWorkbook
    mySN = "新工作表"
    
    Set mySht = myWB.Sheets.Add(After:=myWB.Worksheets(myWB.Worksheets.Count))
    mySht.Name = mySN
    
End Sub
Sub 複製SHEET移動到所有SHEET後面_資料庫()
    Sheets("原始表").copy After:=Sheets(Sheets.Count)
    With ActiveSheet
        .Name = "TOTAL"
        With .Tab
        .ColorIndex = xlColorIndexNone
        End With
    End With
End Sub
Sub 判斷此檔內是否有SHEET名叫特定名稱_資料庫()


For Each EBE In Worksheets
    If EBE.Name = "SPEC" Then
        TEMP_B = 1
        Exit For
    End If
Next
 
If TEMP_B = 0 Then
    MsgBox "所開啟的檔案內找不到名為""SPEC""的工作表，請確定是否開錯檔案!"
    With Workbooks(PM_NAME)
        .RunAutoMacros xlAutoClose
        .Close
    End With
    Exit Sub
End If
End Sub
Sub 強制刪除工作表_資料庫()
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("PS_GROUP_UNIT").Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
 Sub 將工作表PS_GROUP_UNIT及PS_ORDER自此檔刪除並從他檔載入此兩工作表_資料庫()
 Application.ScreenUpdating = False
 '取得現在活動中的SHEET NAME
 ORIGIN_SHEET = ActiveSheet.Name
 '---------------------------------
 '載入SPEC-PM檔的PS_GROUP_UNIT及PS_ORDER工作表
 '---------------------------------
 '取得PM的檔名
 PM_PATH = Application.GetOpenFilename(FileFilter:="Excel檔案(*.xls;*.xlsx),*.xls;*.xlsx", Title:="載入SPEC-PM檔..")
 PM_NAME = Right(PM_PATH, Len(PM_PATH) - InStrRev(PM_PATH, "\"))

 If PM_PATH = False Then
    MsgBox "操作取消，請重新執行此功能!"
    Exit Sub
    Else
 '開啟SPEC-PM來複製資料
 '==>判斷SPEC-PM檔案是否已經開啟
TEMP_A = 0
    For Each EBE In Workbooks
        If EBE.Name = PM_NAME Then
            TEMP_A = 1
        End If
    Next
 
    If TEMP_A = 0 Then
        Workbooks.Open fileName:=PM_PATH, ReadOnly:=True
    End If
End If
'==>刪除此檔內的舊工作表並從他檔載入該工作表
Windows(ThisWorkbook.Name).Activate
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("PS_GROUP_UNIT").Delete
    Sheets("PS_ORDER").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
Windows(PM_NAME).Activate
    '防呆:開錯檔或SHEET名稱對儲存格值不對則離開SUB
    On Error GoTo 20
        If Sheets("PS_GROUP_UNIT").Range("C1").Value <> "PS_UNIT_NAME" Or Sheets("PS_ORDER").Range("B1").Value <> "PS_ORDER" Then
            MsgBox "   載入的檔案不正確! 請確認後再次執行此功能。            ", vbCritical
            If TEMP_A = 0 Then
                Application.DisplayAlerts = False
                Workbooks(PM_NAME).Close SaveChanges:=False
                Application.DisplayAlerts = True
            End If
            Exit Sub
        End If
    On Error GoTo 0


    Sheets("PS_GROUP_UNIT").copy After:=Workbooks(ThisWorkbook.Name).Sheets("SYSTEM")
Windows(PM_NAME).Activate
    Sheets("PS_ORDER").copy After:=Workbooks(ThisWorkbook.Name).Sheets("PS_GROUP_UNIT")
'如果原來沒開SPEC-PM,則關閉SPEC-PM
 If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(PM_NAME).Close SaveChanges:=False
    Application.DisplayAlerts = True
 End If
 '---------------------------------
 '比對PS資料是否相符
 '---------------------------------
 Call CHK_IMPORT
'==>將FOCUS帶回執行前位置
    Windows(ThisWorkbook.Name).Activate
    Sheets(ORIGIN_SHEET).Select
If Sheets("SYSTEM").Range("B1") = "X" Then Exit Sub
MsgBox "   載入SPEC-PM完畢!            ", vbInformation
Exit Sub
'載入不正確時的錯誤訊息
20
MsgBox "   載入的檔案不正確! 請確認後再次執行此功能。            ", vbCritical
If TEMP_A = 0 Then
    Application.DisplayAlerts = False
    Workbooks(PM_NAME).Close SaveChanges:=False
    Application.DisplayAlerts = True
    Exit Sub
End If
 End Sub

Sub 偵測SHEET是否有缺少_資料庫()
'vvvvvvvvvvvvvvvvvvvv防呆vvvvvvvvvvvvvvvvvvvv

'檢查主表及編碼表是否有缺
On Error GoTo 777
    Sheets("主表").Select
    Sheets("PIPES編碼表").Select
    Sheets("FITTINGS編碼表").Select
    Sheets("FLANGES編碼表").Select
    Sheets("BOLT&NUTS編碼表").Select
    Sheets("GASKETS編碼表").Select
    Sheets("VALVES編碼表").Select
    Sheets("SCH編碼表").Select
    
On Error GoTo 0

        
'邊碼錶缺表時跳出錯誤訊息

'^^^^^^^^^^^^^^^^^^^^防呆^^^^^^^^^^^^^^^^^^^^

777 '離開程式用-有缺工作表版
'==>產生錯誤訊息
Title = "錯誤訊息            "    ' 定義標題。
    Msg = vbLf + "偵測到缺少以下列表當中的 一個 或 多個 工作表，請檢查後再次執行此程式。   " + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "*如果是工作表名稱有誤請修正；如果是缺少該工作表請設法補上。   " + vbLf + vbCrLf + vbLf ' 定義訊息。
    Msg = Msg + "   [主表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [PIPES編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [FITTINGS編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [FLANGES編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [BOLT&NUTS編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [GASKETS編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [VALVES編碼表]" + vbLf + vbCrLf  ' 定義訊息。
    Msg = Msg + "   [SCH編碼表]" + vbLf + vbCrLf  ' 定義訊息。

            MsgBox Msg, vbExclamation, Title
End Sub
Sub 保護工作表()
'保護
ActiveSheet.Protect Password:="ABC", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
'取消保護
ActiveSheet.Unprotect Password:="ABC"
End Sub
Sub 呼叫其他工作表的Worksheet_Change()
'呼叫其他工作表的 Private Sub Worksheet_Change(ByVal Target As Range)
Dim myWS As Worksheet
Dim myShtCodeName As String 'myShtCodeName:工作表的CodeName屬性值
Dim myRg As Range   '工作表改變時，正在修改的範圍
Set myWS = Sheet("某張工作表")
Set myRg = myWS.Range("A1")
    myShtCodeName = myWS.CodeName
    Application.Run myShtCodeName & ".Worksheet_Change", myRg
End Sub
Sub 資料庫_檢查輸入值是否違反工作表名稱限制(ByVal myVal As String)
Dim haveErr As Boolean
Dim iStr1 As String, iStr2 As String
Dim i As Long
Dim myArr(8)
    haveErr = False
    '編號要當作工作表名稱，受限於工作表的限制
    '   (1)不可超過31個字 (2)不可有不允許的字元 :： \ / ? * [ ] (3)不可是空白(允許space)
    iStr1 = myVal
    iStr2 = ""
    If (iStr1 = "") Then
        iStr2 = "至少須輸入一個字"
    ElseIf (Len(iStr1) <= 31) Then
        iStr2 = "不可超過31個字"
    Else
        myArr(1) = ":"
        myArr(2) = "；"
        myArr(3) = "\"
        myArr(4) = "/"
        myArr(5) = "?"
        myArr(6) = "*"
        myArr(7) = "["
        myArr(8) = "]"
        For i = 1 To UBound(myArr, 1)
            If (InStr(1, iStr1, myArr(i)) <> 0) Then
                iStr2 = "不可使用這些符號 :： \ / ? * [ ] "
                Exit For
            End If
        Next i
    End If
    
    If (iStr2 <> "") Then
        msgTitle = "錯誤            "    ' 定義標題。
        msgText = "輸入的編號有以下錯誤，請修改" + vbLf    ' 定義訊息。
        msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
        msgText = msgText + iStr2 + vbLf  ' 定義訊息
        msgStyle = vbCritical '顯示"X"圖案
        MsgBox msgText, msgStyle, msgTitle
    End If

End Sub
