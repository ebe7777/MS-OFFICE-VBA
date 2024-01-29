VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form99Printers 
   Caption         =   "列印設定"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   OleObjectBlob   =   "關於取得印表機名稱.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Form99Printers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Activate()
Dim iStr1 As String, iStr2 As String
    iStr1 = Application.ActivePrinter
    printerName = Left(iStr1, InStr(1, iStr1, " on") - 1)
    TextBox1.Value = printerName
End Sub
Private Sub CommandButton1_Click()
    Application.Dialogs(xlDialogPrinterSetup).Show
    iStr1 = Application.ActivePrinter
    printerName = Left(iStr1, InStr(1, iStr1, " on") - 1)
    TextBox1.Value = printerName
End Sub
Private Sub CommandButton2_Click()
    Form99Printers.Hide
    printerName = TextBox1.Value
End Sub
Private Sub UserForm_Terminate()
    printerName = ""
    stopRun = True
    Unload Form99Printers
End Sub

