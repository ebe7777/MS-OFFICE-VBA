Attribute VB_Name = "����]�w�ޥζ���"
Sub �]�w�M���γ]�w�ޥζ��ت��g�k()


'https://stackoverflow.com/questions/47056390/how-to-do-late-binding-in-vba
'Late binding does not require a reference to Outlook Library 16.0 whereas early binding does. However, note that late binding is a bit slower and you won't get intellisense for that object.

Dim objWordApp As Object, objWordDoc As Object
'�n�]�w�ޥζ���
'This is early binding:
Dim objWordApp As Outlook.Application
Set objWordApp = New Outlook.Application

'���γ]�w�ޥζ���
'And this is late binding:
Dim objWordApp As Object
Set objWordApp = CreateObject("Outlook.Application")

End Sub
