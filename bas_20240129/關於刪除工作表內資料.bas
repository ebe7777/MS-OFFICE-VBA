Attribute VB_Name = "����R���u�@�����"
Sub �M�Žd���x�s����_��Ʈw()
'�u�O�R����ƨëD�R����C
Sheets("SYSTEM").Range("A:G").ClearContents
End Sub

Public Function clearRange_��Ʈw(myRange As Range)
'�M���x�s�� ���e�B����B�r��
    With myRange
        .Formula = ""
        '.ClearComments
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
End Function
