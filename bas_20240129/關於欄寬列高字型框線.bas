Attribute VB_Name = "關於欄寬列高字型框線"
Sub 整理欄寬列高字型_資料庫()
Sheets("BASE_FORM").Select
'整理欄寬
Columns("A:AY").ColumnWidth = 2.13
'整理抬頭列列高
Rows("1:1").RowHeight = 19.5
Rows("2:11").RowHeight = 16.5
'修改字形
Cells.Font.Name = "Consolas"
End Sub




Function 考量換行而增多的字數_資料庫(ByVal DATA_ADD, ORIG_WORD_NUM, CHANGE_ROW_WORD_NUM, ADD_WORD_NUM)
'(ByVal 敘述欄位,原始敘述字數,換行字數(累計,第一行44,則第二行88),增加字數)

'設定增加字數 = 增加字數 而不等於0,是用於判斷第二行,第三行換行時增加字數要累加前一行的增加數;故引用此FUNCTION的程式再進入第二回圈之前要將
'(接上)最後的計算結果 (考量換行而增多的字數_資料庫)寫入(ADD_WORD_NUM)裡
'ADD_WORD_NUM = ADD_WORD_NUM
'如果原來的字數+增加的字數會大於換行字數(代表還需要判斷是否換下一行),才需計算
If ORIG_WORD_NUM + ADD_WORD_NUM > CHANGE_ROW_WORD_NUM Then
    '如果換行字數該字為空格" "或中線"-",又或者該換行字的下一個字是空格,系統會自動換行不會多加字數;故不是才需計算
    If Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> " " And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> "-" And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM + 1, 1) <> " " Then
        '判斷分頁點前是先有" "還是"-"來判斷該用哪個來判斷字數
        If InStrRev(DATA_ADD, " ", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM) < InStrRev(DATA_ADD, "-", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM) Then
            '總 增加字數 = 前一行的增加字數 + 此行的增加字數
            考量換行而增多的字數_資料庫 = ADD_WORD_NUM + (CHANGE_ROW_WORD_NUM - ADD_WORD_NUM - InStrRev(DATA_ADD, "-", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM))
            Else
            考量換行而增多的字數_資料庫 = ADD_WORD_NUM + (CHANGE_ROW_WORD_NUM - ADD_WORD_NUM - InStrRev(DATA_ADD, " ", CHANGE_ROW_WORD_NUM - ADD_WORD_NUM))
        End If
    '如果不需計算，且該行末尾的接下來字為空格，因為空格不會待到下一行頭而是省略，故總字數要扣減
    ElseIf Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM, 1) <> " " And Mid(DATA_ADD, CHANGE_ROW_WORD_NUM - ADD_WORD_NUM + 1, 1) = " " Then
        考量換行而增多的字數_資料庫 = ADD_WORD_NUM - 1
        Else
        '最後一種可能 第一行末尾為空格,下一個字不為空格
        考量換行而增多的字數_資料庫 = ADD_WORD_NUM
    End If
    Else
    '如果不需計算,現在 總 增加字數=截至上一行為止的 總 增加字數
    考量換行而增多的字數_資料庫 = ADD_WORD_NUM
End If
    
End Function


Public Function rangeSetBoardLine_資料庫(workbookName As String, shtName As String, starRange As String, endRange As String, myLineStyle As XlLineStyle)
'整欄範圍 (e.g. A:A) 或 有限範圍 (e.g.A1:A3) 加上格線
    '無框線 "xlNone"
    '實線 "xlContinuous"
    '點 "xlDot"
    '虛線 "xlDash"

    With Workbooks(workbookName).Worksheets(shtName).Range(starRange & ":" & endRange)
        If (myLineStyle = xlNone) Then
            .Borders(xlDiagonalDown).lineStyle = myLineStyle
            .Borders(xlDiagonalUp).lineStyle = myLineStyle
        End If
        .Borders(xlEdgeLeft).lineStyle = myLineStyle
        .Borders(xlEdgeTop).lineStyle = myLineStyle
        .Borders(xlEdgeBottom).lineStyle = myLineStyle
        .Borders(xlEdgeRight).lineStyle = myLineStyle
        .Borders(xlInsideVertical).lineStyle = myLineStyle
        .Borders(xlInsideHorizontal).lineStyle = myLineStyle
    End With
End Function
