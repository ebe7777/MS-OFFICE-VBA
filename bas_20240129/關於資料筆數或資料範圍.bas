Attribute VB_Name = "�����Ƶ��Ʃθ�ƽd��"

Public Function myDataRows_��Ʈw(ByVal bookName As String, ByVal sheetName As String, ByVal columnName As String, ByVal countUpwardFromThisRow As Long)
'�H�U�¤覡�p�G�J��̫�X�C�Q�z�ﱼ/���áA�h�ӳ������|�ǤJ�p��d��
'    myDataRows = Workbooks(bookName).Sheets(sheetName).Range(columnName & countUpwardFromThisRow).End(xlUp).Row

Dim myArray As Variant
Dim i As Long
    '���[�tŪ���t�סA���N�x�s��ȼg�J�}�C
    With Workbooks(bookName).Sheets(sheetName)
        myArray = .Range(columnName & "1:" & columnName & countUpwardFromThisRow).Formula
    End With
    For i = UBound(myArray, 1) To 1 Step -1
        If (myArray(i, 1) <> "") Then
            myDataRows = i
            Exit For
        End If
    Next i
' �p�G�Nfunction��m�b�@��sub���A�t�@��sub�n�I�s��sub��function�A�ϥ� call Module�W��1.Sub�W��
End Function
Public Function myDataColumns_��Ʈw(ByVal bookName As String, ByVal sheetName As String, ByVal rowNumber As Long, ByVal countLeftwardFromThisColumn As String)
'�H�U�¤覡�p�G�J��̫�X��Q���áA�h�ӳ������|�ǤJ�p��d��
'   myDataColumns = Workbooks(bookName).Sheets(sheetName).Range(countLeftwardFromThisColumn & rowNumber).End(xlToLeft).Column
'�ݷf�t convertABCto123 �ϥ�
'   ����� convertABCto123 �ثe�u��B�z�� ZZ�A�ҥH��̦h�u����ZZ

Dim myArray As Variant
Dim i As Long, maxCol As Long
    '���[�tŪ���t�סA���N�x�s��ȼg�J�}�C
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
Public Function findMaxRowNo_��Ʈw(bookName As String, sheetName As String, startColName As String, endColName As String)
'���Y�ɬY�u�@����w����d�򤺨ϥΪ��̦h�C�����X
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
Public Function findMaxColNo_��Ʈw(bookName As String, sheetName As String, startRowNo As Long, endRolNo As Long)
'���Y�ɬY�u�@����w���C�d�򤺨ϥΪ��̦h�檺���X
'   ����� convertABCto123 �ثe�u��B�z�� ZZ�A�ҥH��̦h�u����ZZ
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
Sub ����w�ϥΪ��u�@�d��_��Ʈw()
ActiveSheet.UsedRange.Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'���o���ק諸�s�x�檺range�A�ûP�ۭq���@�ӽd�����o�� "�ק諸�O�_�b���S�w�d��"
'��sub�u��g�b�u�@��
'https://docs.microsoft.com/zh-tw/office/troubleshoot/excel/run-macro-cells-change
'https://docs.microsoft.com/zh-tw/office/vba/api/excel.application.intersect
Dim myRange1 As Range
Dim myRange2 As Range
    Set myRange1 = Range("C10:C11")
    Set myRange2 = Range(Target.address)
        
    If (Application.Intersect(myRange1, myRange2) Is Nothing) Then
        MsgBox "�ק諸�x�s�� ���b ���w�d��"
    Else
        MsgBox "�ק諸�x�s �b ���w�d��"
    End If
End Sub

