Attribute VB_Name = "����ץX��PDMS��MACRO"
Private Sub �N��ƾ�z��PDMS�ϥΪ�MACRO_��Ʈw()
Dim mainSN As String
Dim macroArray()
Dim totalRows As Integer
Dim i As Integer, ii As Integer
Dim outputFilePath As FileDialog
Dim outputFilePathString As String
Dim outputFileName As String
Dim macroString As String
Dim msgTitle As String, msgText As String, msgStyle As String

mainSN = "�D��"

'��ܿ�X��Ƨ�
    Set outputFilePath = Application.FileDialog(msoFileDialogFolderPicker)
    With outputFilePath
        .Title = "��ܭn�ɮצs���m(��Ƨ�)"
        .AllowMultiSelect = False
    End With
    outputFilePath.Show
    If outputFilePath.SelectedItems.Count = 0 Then
        GoTo 991
    Else
        outputFilePathString = outputFilePath.SelectedItems.Item(1)
    End If

'�̾ڥD���Ƽg��macro�A�üg�Jarray��
    totalRows = dataRows(mainSN, "A")
    ii = 0
    For i = 2 To totalRows
        Call arrayAddNewData(ii, macroArray, "!skipThis = false")
        Call arrayAddNewData(ii, macroArray, "/" & Cells(i, 1).Value)
        Call arrayAddNewData(ii, macroArray, "handle(2,109)")
        Call arrayAddNewData(ii, macroArray, "  !skipThis = true")
        Call arrayAddNewData(ii, macroArray, "endhandle")
        Call arrayAddNewData(ii, macroArray, "if (!skipThis eq false) then")
        Call arrayAddNewData(ii, macroArray, "  :XOPRESS '" & Cells(i, 2) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XDPRESS '" & Cells(i, 3) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XOTEMP '" & Cells(i, 4) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XDTEMP '" & Cells(i, 5) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XHYDRO '" & Cells(i, 6) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XPNEUM '" & Cells(i, 7) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XNDTPT '" & Cells(i, 8) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XNDTMT '" & Cells(i, 9) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XNDTRT '" & Cells(i, 10) & "'")
        Call arrayAddNewData(ii, macroArray, "  :XREFDWG '" & Cells(i, 11) & "'")
        Call arrayAddNewData(ii, macroArray, "endif")
    Next i
    Call arrayAddNewData(ii, macroArray, "$* complete message.")
    Call arrayAddNewData(ii, macroArray, "!!alert.message(|Line list data input completed.|)")

    '==================================================================================
    '
    '��X��ƨ��r��
    '----------------------------------------------------------------------------------
    outputFileName = outputFilePathString & "\LineListInputMacro.mac"

    totalRows = UBound(macroArray, 1)
    
    Open outputFileName For Output As #1
    For i = 1 To totalRows
        Print #1, macroArray(i)
    Next i

    Close #1
    '���槹���T��
    msgTitle = "�T��            "    ' �w�q���D�C
    msgText = "���槹�� !�C" + vbLf + vbCrLf  ' �w�q�T���C
    msgText = msgText + "====================================" + vbLf + vbCrLf ' �w�q�T���C
    msgText = msgText + "�����ɮת���m�G" + vbLf + vbCrLf  ' �w�q�T��
    msgText = msgText + outputFileName + vbLf + vbCrLf   ' �w�q�T��
    msgText = msgText + "-->�ЦbPDMS�����榹MAC��" + vbLf + vbCrLf   ' �w�q�T��
    MsgBox msgText, vbInformation, msgTitle
    Exit Sub
991
    '���_����T��
    msgTitle = "�T��            "   ' �w�q���D�C
    msgText = "  �Э��s���榹�{�� !"
    MsgBox msgText, vbInformation, msgTitle
    Exit Sub
End Sub

Function dataRows(sheetName, columnName)
    ' �p��̫�@�C�g��function
    dataRows = Sheets(sheetName).Range(columnName & "100000").End(xlUp).Row
End Function
Private Function arrayAddNewData(i, arrayName, data)
    i = i + 1
    ReDim Preserve arrayName(i)
    arrayName(i) = data
End Function
