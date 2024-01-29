VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "程式執行中..."
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   5025
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=========
'開發者     brucechen1@micb2b.com
'開發日期   2020-03-11
'修改日期   2023-12-19
'=========

Private Sub UserForm_Initialize()
'須設定public variable
'   progressBarPercentNo as long

'Label1的Width顯示進度條 0~240
'       caption顯示進度值,以string形式

Dim labelWidth As Long
    Label1.caption = "0%"
    Label1.Width = 0
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Sub updateProgressBar(ByVal percentNo As Double)
    percentNo = Round(percentNo, 1)
    ProgressBar.Label1.caption = percentNo & "%"
    ProgressBar.Label1.Width = Round(240 * (percentNo / 100), 1)
    DoEvents
End Sub

Sub 使用時在程式放入以下()

'ppppp開頭-呼叫視窗
ThisWorkbook.Activate
progressBarPercentNo = 0
ProgressBar.Show False
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'ppppp中間-固定值寫法
ThisWorkbook.Activate
progressBarPercentNo = progressBarPercentNo + 5
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'ppppp中間-隨i變化寫法
ThisWorkbook.Activate
progressBarPercentNo = progressBarPercentNo + ((1 / maxOFi) * baseValue)
Call ProgressBar.updateProgressBar(progressBarPercentNo)
'ppppp

'按比例增加一定數量的進度
'ppppp-5 to 25
ThisWorkbook.Activate
'   此段執行結束共會增加iProcessRange個進度百分比/資料筆數共iDataCounts筆/每iProcessRangeGap筆從新計算一次進度
iProcessRange = 20
iDataCounts = 50000
iProcessRangeGap = 500
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

'ppppp結束
Unload ProgressBar
ThisWorkbook.Activate
'ppppp
End Sub


Sub 以下留存參考()
'Sub xxx()
    '
    '
    ''<<<<<<<<<<<<<<<<寫在巨集運算時>>>>>>>>>>>>>>>>>>>>>>>
    '
    'STIME = Time 'START TIME起使時間
    'ProgressBar.Show 0
    '
    '
    ''@@@@@@進度條相關@@@@@@'
    'NTIME = Time 'NOW TIME目前時間
    'Call Bar_GO(STEP_NOW, ALL_CACU_ROWS, STIME, NTIME)
    'DoEvents
    ''@@@@@@@@@@@@@@@@@@@@@@
    '
    '
    'Unload ProgressBar '用來關閉進度表
'End Sub
'
'
''<<<<<<<<<<<<<<<<<進度條function>>>>>>>>>>>>>>>>>
'
    'Function ProgressBar_GO(ByVal STEP, TOTAL, START_TIME, NOW_TIME) '變數Step為目前執行的步驟，變數Total為總步驟...
    '
    ''進度條FUNCTION
    '
    ''在此要創造一個表單,在屬性裡修改其名為ProgressBar;拉出LA修改其名SC
    'ProgressBar.SCH.Caption = Round(STEP / TOTAL * 100, 0) & "%"
    'ProgressBar.SCH.Width = Round(STEP / TOTAL * 240, 0)
    'ProgressBar.Caption = "預估剩餘時間：" _
    '    & Minute((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " 分 " & _
    '    Second((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " 秒   "
    ''ProgressBar.Caption = "月報表 [" & ActiveSheet.Name & "] 目前產生進度：" & Format(Round(STEP / TOTAL * 100, 2), "0") & "% ； 預估剩餘時間：" _
    ''    & Minute((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " 分 " & _
    ''    Second((NOW_TIME - START_TIME) * (TOTAL / STEP) - (NOW_TIME - START_TIME)) & " 秒   "
    'DoEvents '用來顯示進度百分比
    'End Function
    ''^^^^^產生工作時數報表 所有SUB已修改符合1.0版^^^^^
'
'End Function
End Sub
