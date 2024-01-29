Attribute VB_Name = "關於Autocad各種資訊顯示"
Sub 常用()
    Dim acad As AcadApplication
    Set acad = GetObject(, "AutoCAD.Application")
    
    '執行vba時隱藏autocad()
    acad.Visible = True
    
    '視窗位置()
    '   最大化
    acad.WindowState = acMax
    '   取得目前視窗的尺寸
    i = acad.Application.Height
    ii = acad.Application.Width
    '   設定視窗左的左上角距離螢幕左上角位置 ;注意，不可在最大化的情況下使用
    acad.WindowState = acNorm
    acad.WindowTop = 1
    acad.WindowLeft = (ii / 2) - 1
    '   設定視窗尺寸
    acad.Height = i
    acad.Width = ii / 2
    '   windows foucus到視窗
    AppActivate acad.Caption
End Sub
Sub 開檔時會遇到的警告訊息()
Dim acad As AcadApplication, dwgFile As AcadDocument
Dim preferences As AcadPreferences, currShowProxyDialogBox As Boolean
    Set acad = GetObject(, "AutoCAD.Application")

'===因custom objects產生的警告訊息
    Set preferences = acad.Application.preferences
    '設置 開檔時不要顯示因custom objects產生的警告訊息
    '   Retrieve the current ShowProxyDialogBox value
    currShowProxyDialogBox = preferences.OpenSave.ShowProxyDialogBox
    '   Change the value for ShowProxyDialogBox
    preferences.OpenSave.ShowProxyDialogBox = Not (currShowProxyDialogBox)

    '開啟autocad
    Set dwgFile = acad.Documents.Open("c:\123.dwg", False)
    '   do something...
    
'===因為AEC物件(if a drawing has AEC object references)
'   無法以VBA處理處理
'   相關說明
'       https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Error-opening-a-drawing-This-application-has-detected-a-mixed-version-of-AEC-objects.html
'       https://forums.autodesk.com/t5/net/how-to-disable-aec-warning-message/td-p/4703443

'===檔案版本較新(This drawing contains objects from a newer version..)
'   查無VBA是否可處理資訊
'   相關說明
'       https://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Error-This-drawing-contains-objects-from-a-newer-version-when-opening-drawing-in-AutoCAD.html
End Sub
