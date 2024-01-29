Attribute VB_Name = "關於物件的操作"
Sub 修改核取方塊的儲存格連結()
Dim iVar As CheckBox
Dim iCol As Long, iRow As Long

For Each iVar In ActiveSheet.CheckBoxes
    '取得核取方塊所在欄位(選取核取方塊後，其周圍限定他的方塊白點要全部在該儲存格的欄列分界內)
    iCol = iVar.BottomRightCell.Column
    iRow = iVar.BottomRightCell.Row
    '設定核取方塊的儲存格連結
    iVar.LinkedCell = Cells(iRow, iCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
Next
MsgBox "執行完畢"
End Sub
