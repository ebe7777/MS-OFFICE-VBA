Attribute VB_Name = "關於匯出成PDMS的MACRO"
Private Sub 將資料整理成PDMS使用的MACRO_資料庫()
Dim mainSN As String
Dim macroArray()
Dim totalRows As Integer
Dim i As Integer, ii As Integer
Dim outputFilePath As FileDialog
Dim outputFilePathString As String
Dim outputFileName As String
Dim macroString As String
Dim msgTitle As String, msgText As String, msgStyle As String

mainSN = "主表"

'選擇輸出資料夾
    Set outputFilePath = Application.FileDialog(msoFileDialogFolderPicker)
    With outputFilePath
        .Title = "選擇要檔案存放位置(資料夾)"
        .AllowMultiSelect = False
    End With
    outputFilePath.Show
    If outputFilePath.SelectedItems.Count = 0 Then
        GoTo 991
    Else
        outputFilePathString = outputFilePath.SelectedItems.Item(1)
    End If

'依據主表資料寫成macro，並寫入array內
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
    '輸出資料到文字檔
    '----------------------------------------------------------------------------------
    outputFileName = outputFilePathString & "\LineListInputMacro.mac"

    totalRows = UBound(macroArray, 1)
    
    Open outputFileName For Output As #1
    For i = 1 To totalRows
        Print #1, macroArray(i)
    Next i

    Close #1
    '執行完畢訊息
    msgTitle = "訊息            "    ' 定義標題。
    msgText = "執行完畢 !。" + vbLf + vbCrLf  ' 定義訊息。
    msgText = msgText + "====================================" + vbLf + vbCrLf ' 定義訊息。
    msgText = msgText + "產生檔案的位置：" + vbLf + vbCrLf  ' 定義訊息
    msgText = msgText + outputFileName + vbLf + vbCrLf   ' 定義訊息
    msgText = msgText + "-->請在PDMS中執行此MAC檔" + vbLf + vbCrLf   ' 定義訊息
    MsgBox msgText, vbInformation, msgTitle
    Exit Sub
991
    '中斷執行訊息
    msgTitle = "訊息            "   ' 定義標題。
    msgText = "  請重新執行此程式 !"
    MsgBox msgText, vbInformation, msgTitle
    Exit Sub
End Sub

Function dataRows(sheetName, columnName)
    ' 計算最後一列寫成function
    dataRows = Sheets(sheetName).Range(columnName & "100000").End(xlUp).Row
End Function
Private Function arrayAddNewData(i, arrayName, data)
    i = i + 1
    ReDim Preserve arrayName(i)
    arrayName(i) = data
End Function
