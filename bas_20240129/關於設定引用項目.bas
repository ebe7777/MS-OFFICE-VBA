Attribute VB_Name = "關於設定引用項目"
Sub 設定和不用設定引用項目的寫法()


'https://stackoverflow.com/questions/47056390/how-to-do-late-binding-in-vba
'Late binding does not require a reference to Outlook Library 16.0 whereas early binding does. However, note that late binding is a bit slower and you won't get intellisense for that object.

Dim objWordApp As Object, objWordDoc As Object
'要設定引用項目
'This is early binding:
Dim objWordApp As Outlook.Application
Set objWordApp = New Outlook.Application

'不用設定引用項目
'And this is late binding:
Dim objWordApp As Object
Set objWordApp = CreateObject("Outlook.Application")

End Sub
