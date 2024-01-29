Attribute VB_Name = "關於資料筆數或資料範圍"

Public Function myDataRows_資料庫(ByVal bookName As String, ByVal sheetName As String, ByVal columnName As String, ByVal countUpwardFromThisRow As Long)
'以下舊方式如果遇到最後幾列被篩選掉/隱藏，則該部分不會納入計算範圍
'    myDataRows = Workbooks(bookName).Sheets(sheetName).Range(columnName & countUpwardFromThisRow).End(xlUp).Row

Dim myArray As Variant
Dim i As Long
    '為加速讀取速度，先將儲存格值寫入陣列
    With Workbooks(bookName).Sheets(sheetName)
        myArray = .Range(columnName & "1:" & columnName & countUpwardFromThisRow).Formula
    End With
    For i = UBound(myArray, 1) To 1 Step -1
        If (myArray(i, 1) <> "") Then
            myDataRows = i
            Exit For
        End If
    Next i
' 如果將function放置在一個sub中，另一個sub要呼叫此sub的function，使用 call Module名稱1.Sub名稱
End Function
Public Function myDataColumns_資料庫(ByVal bookName As String, ByVal sheetName As String, ByVal rowNumber As Long, ByVal countLeftwardFromThisColumn As String)
'以下舊方式如果遇到最後幾欄被隱藏，則該部分不會納入計算範圍
'   myDataColumns = Workbooks(bookName).Sheets(sheetName).Range(countLeftwardFromThisColumn & rowNumber).End(xlToLeft).Column
'需搭配 convertABCto123 使用
'   限於於 convertABCto123 目前只能處理到 ZZ，所以欄最多只能找到ZZ

Dim myArray As Variant
Dim i As Long, maxCol As Long
    '為加速讀取速度，先將儲存格值寫入陣列
    maxCol = convertABCto123(countLeftwardFromThisColumn)
    With Workbooks(bookName).Sheets(sheetName)
        myArray = .Range(.Cells(rowNumber, 1), .Cells(rowNumber, maxCol)).Formula
    End With
    For i = UBound(myArray, 2) To 1 Step -1
        If (myArray(1, i) <> "") Then
            myDataColumns = i
            Exit For
        End If
    Next i
End Function
Public Function findMaxRowNo_資料庫(bookName As String, sheetName As String, startColName As String, endColName As String)
'找到某檔某工作表指定的欄範圍內使用的最多列的號碼
Dim iStart As Long, iEnd As Long, iRows As Long, iMaxRows As Long
Dim i As Long
    iStart = convertABCto123(startColName)
    iEnd = convertABCto123(endColName)
    iMaxRows = 0
    For i = iStart To iEnd
        iRows = myDataRows(bookName, sheetName, convert123toABC(i), 65536)
        If (iRows > iMaxRows) Then
            iMaxRows = iRows
        End If
    Next i
    findMaxRowNo = iMaxRows
End Function
Public Function findMaxColNo_資料庫(bookName As String, sheetName As String, startRowNo As Long, endRolNo As Long)
'找到某檔某工作表指定的列範圍內使用的最多欄的號碼
'   限於於 convertABCto123 目前只能處理到 ZZ，所以欄最多只能找到ZZ
Dim iCols As Long, iMaxCols As Long
Dim i As Long

    iMaxCols = 0
    For i = startRowNo To endRolNo
        iCols = myDataColumns(bookName, sheetName, i, "ZZ")
        If (iCols > iMaxCols) Then
            iMaxCols = iCols
        End If
    Next i
    findMaxColNo = iMaxCols
End Function
Sub 選取已使用的工作範圍_資料庫()
ActiveSheet.UsedRange.Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'取得剛剛修改的存儲格的range，並與自訂的一個範圍比較得知 "修改的是否在此特定範圍內"
'此sub只能寫在工作表中
'https://docs.microsoft.com/zh-tw/office/troubleshoot/excel/run-macro-cells-change
'https://docs.microsoft.com/zh-tw/office/vba/api/excel.application.intersect
Dim myRange1 As Range
Dim myRange2 As Range
    Set myRange1 = Range("C10:C11")
    Set myRange2 = Range(Target.address)
        
    If (Application.Intersect(myRange1, myRange2) Is Nothing) Then
        MsgBox "修改的儲存格 不在 指定範圍內"
    Else
        MsgBox "修改的儲存 在 指定範圍內"
    End If
End Sub

