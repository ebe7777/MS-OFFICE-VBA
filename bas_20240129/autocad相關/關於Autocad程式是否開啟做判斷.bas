Attribute VB_Name = "����Autocad�{���O�_�}�Ұ��P�_"
Option Explicit

Private Sub checkAutoCad()
Dim obj As Object
  
    On Error Resume Next
    Set obj = GetObject(, "AutoCAD.Application")
    If Err Then
        MsgBox "�Х�����AutoCAD�C", vbExclamation, "������AutoCAD"
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
        MsgBox "�Х��}��AutoCAD�{��" + vbLf
        msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
        msgText = msgText + "����L�{�ݨϥΨ�AutoCAD�A�G�Ф�ʶ}�ҫ�A�װ��榹�{��"   ' �w�q�T��
        msgStyle = vbExclamation '���"!"�Ϯ�
    
        MsgBox msgText, msgStyle, msgTitle
        
        stopRun = True
        Exit Sub
    End If
    On Error GoTo 0
End Sub
