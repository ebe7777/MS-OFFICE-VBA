VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form02ChooseAutoCadVersion 
   Caption         =   "選擇軟體版次"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4095
   OleObjectBlob   =   "Form02ChooseAutoCadVersion.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Form02ChooseAutoCadVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'20200622待執行 :
'   說明書告知使用者需要設定 [檔案]>[選項]>[信任中心]>[信任中心設定(T)...]>[巨集設定]>勾選[信任存取VBA專案物件模型(V)]

'需搭配VBA REFERENCE SETTING使用
Private Sub UserForm_Activate()
    ComboBox01.Clear
    ComboBox01.AddItem "2010            (64位元)"
    ComboBox01.AddItem "2014            (64位元)"
    ComboBox01.AddItem "2015/2016 (64位元)"
    ComboBox01.AddItem "2017            (64位元)"
    ComboBox01.AddItem "2018            (64位元)"
    ComboBox01.ListIndex = ThisWorkbook.Worksheets("VBA REFERENCE SETTING").Range("B2").Value
End Sub
Private Sub CommandButton01_Click()
    Me.Hide
'==啟動以下兩者會使的錯誤發生時無法使用偵錯==
    Call clearReferenceAutoCAD
    Call loadReferenceAutoCAD
'============================================
    '執行後續的AutoCAD操作
'    Call run0203_AutocadBatchPlotLayout("pdf")
End Sub
'刪除此軟體的引用項目
Private Function clearReferenceAutoCAD()
Dim ref As Object
Dim refs As Object

    Set refs = Application.VBE.ActiveVBProject.References
    For Each ref In refs
      On Error Resume Next
      If ref.name = "AutoCAD" Then
          refs.Remove ref
      End If
    Next
End Function
'新增此軟體的引用項目
Private Function loadReferenceAutoCAD()
Dim obj As Object
Dim guid As String
Dim softwareVersion As String

    '記錄此次combobox的listIndex選擇值
    ThisWorkbook.Worksheets("VBA REFERENCE SETTING").Range("B2").Value = ComboBox01.ListIndex
    
'    ComboBox01.AddItem "2010            (64位元)"
'    ComboBox01.AddItem "2014            (64位元)"
'    ComboBox01.AddItem "2015/2016 (64位元)"
'    ComboBox01.AddItem "2017            (64位元)"
'    ComboBox01.AddItem "2018            (64位元)"
    softwareVersion = ComboBox01.ListIndex
    
    Select Case softwareVersion
    '2010 64bits
    Case 0
        guid = "{E072BCE4-9027-4F86-BAE2-EF119FD0A0D3}"
    '2014 64bits
    Case 1
        guid = "{D5C3CB6F-AA0A-4D45-B02D-CF2974EFD4BE}"
    '2015,2016 64bits
    Case 2
        guid = "{4E3F492A-FB57-4439-9BF0-1567ED84A3A9}"
    '2017 64bits
    Case 3
        guid = "{5B3245BE-661C-4324-BB55-3AD94EBBFDD7}"
    '2018 64bits
    Case 4
        guid = "{644614D2-93DC-48C6-A061-21ABCE65A4C0}"
    End Select
    
    '防呆-使用者選擇了電腦沒安裝的版次
    On Error GoTo 991
    Application.VBE.ActiveVBProject.References.AddFromGuid guid, 1, 0
    On Error GoTo 0
    GoTo 999
991
    msgTitle = "訊息"    ' 定義標題。
    msgText = "此電腦尚未安裝此版本的軟體" + vbLf   ' 定義訊息。
    msgText = msgText + "==================" + vbLf ' 定義訊息。
    msgText = msgText + "-->請重新選擇"  ' 定義訊息
    MsgBox msgText, vbExclamation, msgTitle
    Me.Show False
    Exit Function
999

End Function




