Attribute VB_Name = "關於設定保護"
Function protectWorkbook_資料庫()
    'workbook.Protect([Password], [Structure], [Windows])
    'structure:工作表不可移動、改名、新增刪除
    '   True protects the order of sheets in the workbook; False does not protect. Default is False.
    'windows:不確定
    '   True protects the location and appearance of the Excel windows used to display the workbook; False does not protect. Default is False.
    ActiveWorkbook.Protect "770916", True, False
End Function

Function unprotectWorkbook_資料庫()
    ActiveWorkbook.Unprotect "770916"
End Function
