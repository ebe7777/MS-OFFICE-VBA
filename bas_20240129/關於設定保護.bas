Attribute VB_Name = "����]�w�O�@"
Function protectWorkbook_��Ʈw()
    'workbook.Protect([Password], [Structure], [Windows])
    'structure:�u�@���i���ʡB��W�B�s�W�R��
    '   True protects the order of sheets in the workbook; False does not protect. Default is False.
    'windows:���T�w
    '   True protects the location and appearance of the Excel windows used to display the workbook; False does not protect. Default is False.
    ActiveWorkbook.Protect "770916", True, False
End Function

Function unprotectWorkbook_��Ʈw()
    ActiveWorkbook.Unprotect "770916"
End Function
