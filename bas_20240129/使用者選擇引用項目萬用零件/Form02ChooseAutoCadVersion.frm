VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form02ChooseAutoCadVersion 
   Caption         =   "��ܳn�骩��"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4095
   OleObjectBlob   =   "Form02ChooseAutoCadVersion.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Form02ChooseAutoCadVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'20200622�ݰ��� :
'   �����ѧi���ϥΪ̻ݭn�]�w [�ɮ�]>[�ﶵ]>[�H������]>[�H�����߳]�w(T)...]>[�����]�w]>�Ŀ�[�H���s��VBA�M�ת���ҫ�(V)]

'�ݷf�tVBA REFERENCE SETTING�ϥ�
Private Sub UserForm_Activate()
    ComboBox01.Clear
    ComboBox01.AddItem "2010            (64�줸)"
    ComboBox01.AddItem "2014            (64�줸)"
    ComboBox01.AddItem "2015/2016 (64�줸)"
    ComboBox01.AddItem "2017            (64�줸)"
    ComboBox01.AddItem "2018            (64�줸)"
    ComboBox01.ListIndex = ThisWorkbook.Worksheets("VBA REFERENCE SETTING").Range("B2").Value
End Sub
Private Sub CommandButton01_Click()
    Me.Hide
'==�ҰʥH�U��̷|�Ϫ����~�o�ͮɵL�k�ϥΰ���==
    Call clearReferenceAutoCAD
    Call loadReferenceAutoCAD
'============================================
    '�������AutoCAD�ާ@
'    Call run0203_AutocadBatchPlotLayout("pdf")
End Sub
'�R�����n�骺�ޥζ���
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
'�s�W���n�骺�ޥζ���
Private Function loadReferenceAutoCAD()
Dim obj As Object
Dim guid As String
Dim softwareVersion As String

    '�O������combobox��listIndex��ܭ�
    ThisWorkbook.Worksheets("VBA REFERENCE SETTING").Range("B2").Value = ComboBox01.ListIndex
    
'    ComboBox01.AddItem "2010            (64�줸)"
'    ComboBox01.AddItem "2014            (64�줸)"
'    ComboBox01.AddItem "2015/2016 (64�줸)"
'    ComboBox01.AddItem "2017            (64�줸)"
'    ComboBox01.AddItem "2018            (64�줸)"
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
    
    '���b-�ϥΪ̿�ܤF�q���S�w�˪�����
    On Error GoTo 991
    Application.VBE.ActiveVBProject.References.AddFromGuid guid, 1, 0
    On Error GoTo 0
    GoTo 999
991
    msgTitle = "�T��"    ' �w�q���D�C
    msgText = "���q���|���w�˦��������n��" + vbLf   ' �w�q�T���C
    msgText = msgText + "==================" + vbLf ' �w�q�T���C
    msgText = msgText + "-->�Э��s���"  ' �w�q�T��
    MsgBox msgText, vbExclamation, msgTitle
    Me.Show False
    Exit Function
999

End Function




