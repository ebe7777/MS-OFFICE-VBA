Attribute VB_Name = "關於螢幕顯示資訊及警示"
Sub 資料庫()
'關閉螢幕顯示更新
Application.ScreenUpdating = False

'關閉警告示窗顯示
Application.DisplayAlerts = False

'關閉events執行-在工作表使用event時可使用此避免無限迴圈
Application.EnableEvents = False

'包含修改Listbox的rawsource所讀取的工作表導致觸發listbox
'   關閉自動計算
Application.Calculation = xlCalculationManual
'   開啟自動計算
Application.Calculation = xlCalculationAutomatic

'excel視窗尺寸
'   隱藏
Application.Visible = False
Application.Visible = True
'   最大化
Application.WindowState = xlMaximized
'   目前高
myHeig = Application.Height
'   目前寬
myWid = Application.Width
'   視窗左上角靠螢幕位置；注意，視窗不可在最大化的狀態
Application.WindowState = xlNormal
Application.Top = 1
Application.Left = 1

'focus到excel
AppActivate Application.Caption

'暫停
Application.Wait Now + TimeValue("00:00:05")

'不確定，但應該是介面可不可以操作
Application.Interactive = False '  禁止交互模式 :

'不確定
Application.StatusBar = False '關閉狀態列

'excel不重算 自定義的工作表含數時,強制重算全部工作表，功能同按下CTRL+ALT+F9
'   https://docs.microsoft.com/en-us/office/troubleshoot/excel/custom-function-calculate-wrong-value
Application.Volatile
End Sub
