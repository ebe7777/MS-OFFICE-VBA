Attribute VB_Name = "�`���ݩʻP��k"

Sub METHOD_��Ʈw()
'�N�����ܦ��ŭ�
Set myObj = Nothing
'�N�ܼ��ܦ��ŭ�
myVar = Empty

'�ܼ�(���]�t�}�C)�O�_���ŭ�
'   Returns a Boolean value indicating whether a variable has been initialized.
a = IsEmpty(myVar)

'�}�C�O�_���ŭ�
'   if using "dim myArray()",then this array is not empty,so canot use IsEmtpy
'   if using "dim myArray()",and not redim it,then it's uninitialized
'       uninitialized array (Not myArray) = -1  (Not Not myArray) = 0
If ((Not myArray) = -1) Then
End If

'�O�_���Ʀr - ���G��BOOLEAN
a = IsNumeric(Cells(3, "M"))

'���A�ഫ
a = CStr(123)

'CBool   Boolean    Any valid string or numeric expression.
'CByte   Byte       0 to 255.
'CCur    Currency   -922,337,203,685,477.5808 to 922,337,203,685,477.5807.
'CDate   Date       Any valid date expression.
'CDbl    Double     -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values.
'CDec    Decimal    79,228,162,514,264,337,593,543,950,335 for zero-scaled numbers, that is, numbers with no decimal places. For numbers with 28 decimal places, the range is 7.9228162514264337593543950335. The smallest possible non-zero number is 0.0000000000000000000000000001.
'CInt    Integer    -32,768 to 32,767; fractions are rounded(���Ʒ|�|�ˤ��J).
'CLng    Long       -2,147,483,648 to 2,147,483,647; fractions are rounded.
'CLngLng LongLong   -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807; fractions are rounded. (Valid on 64-bit platforms only.)
'CLngPtr LongPtr    -2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems; fractions are rounded for 32-bit and 64-bit systems.
'CSng    Single     -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values.
'CStr    String     Returns for CStr depend on the expression argument.
'CVar    Variant    Same range as Double for numerics. Same range as String for non-numerics.

'�p��l��
a = 132 Mod 20

'�O����(even)�Ω_��(odd)
If (a Mod 2 = 0) Then
    '�O����
End If

'���Φr��
Dim a As Variant
a = Split("1,2,3", ",")

'�N���i�C�L�r��(Ĵ�palt+Enter����)�R��
a = Application.WorksheetFunction.Clean(Range("A1"))

'�N�r���ର�j�g
a = UCase("abc")

'�Hcell�g�{���A���N���x�s��W�l�H"A1"�Φ���X
a = Cells(1, 1).address(RowAbsolute:=False, ColumnAbsolute:=False)

'�N�r�ꤤ���r������
a = Replace("1,2,3", ",", "")
End Sub

Sub ����()
'�w�q���󬰤u�@��
Dim sysSht As Worksheet
Set sysSht = Worksheets("SYSTEM")
'�M��array���e
Erase threeDimArray, twoDimArray
End Sub


