Attribute VB_Name = "����b�u�@��WŪ���P�g�J���"
Public Function �Pvlookup�Ϊk_���@��ƨæ^�ǦP�C�Y����_vlookupF1R1(findDataInThisWorksht As Worksheet, findDataValue As String, findDataInthisColumn As String, returnDataInThisColumn As String)
Dim i As Integer
Dim totalRows As Integer
    totalRows = myDataRows(findDataInThisWorksht.Name, "A")
    For i = 2 To totalRows
        If (sysSht.Range(findDataInthisColumn & i).Value = findDataValue) Then
            vlookupF1R1 = sysSht.Range(returnDataInThisColumn & i).Value
        End If
        Exit For
    Next i
End Function
Public Function �ҥ�vlookup�Ϊk_���@��ƨæb�P�C�Y��g�J���_vlookupF1W1(findDataInThisWorksht As Worksheet, findDataValue As String, findDataInthisColumn As String, writeThisValue As Variant, writeDataInThisColumn As String)
Dim i As Integer
Dim totalRows As Integer
    totalRows = myDataRows(findDataInThisWorksht.Name, "A")
    For i = 2 To totalRows
        If (sysSht.Range(findDataInthisColumn & i).Value = findDataValue) Then
            sysSht.Range(writeDataInThisColumn & i).Value = writeThisValue
        End If
        Exit For
    Next i
End Function

Sub ����H�u���檺��J�P�R��()
'��J
result = "123" & Chr(10) & "456"
'�R��
result = Replace(originalText, Chr(10), "")
End Sub

Sub ����ƻs�K�W()
Worksheets("Sheet1").Range("A1:D4").Copy Destination:=Worksheets("Sheet2").Range("E5")
End Sub
