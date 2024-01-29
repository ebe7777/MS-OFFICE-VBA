Attribute VB_Name = "關於在工作表上讀取與寫入資料"
Public Function 同vlookup用法_找到一資料並回傳同列某欄資料_vlookupF1R1(findDataInThisWorksht As Worksheet, findDataValue As String, findDataInthisColumn As String, returnDataInThisColumn As String)
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
Public Function 模仿vlookup用法_找到一資料並在同列某欄寫入資料_vlookupF1W1(findDataInThisWorksht As Worksheet, findDataValue As String, findDataInthisColumn As String, writeThisValue As Variant, writeDataInThisColumn As String)
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

Sub 關於人工換行的輸入與刪除()
'輸入
result = "123" & Chr(10) & "456"
'刪除
result = Replace(originalText, Chr(10), "")
End Sub

Sub 關於複製貼上()
Worksheets("Sheet1").Range("A1:D4").Copy Destination:=Worksheets("Sheet2").Range("E5")
End Sub
