Attribute VB_Name = "常用屬性與方法"

Sub METHOD_資料庫()
'將物件變成空值
Set myObj = Nothing
'將變數變成空值
myVar = Empty

'變數(不包含陣列)是否為空值
'   Returns a Boolean value indicating whether a variable has been initialized.
a = IsEmpty(myVar)

'陣列是否為空值
'   if using "dim myArray()",then this array is not empty,so canot use IsEmtpy
'   if using "dim myArray()",and not redim it,then it's uninitialized
'       uninitialized array (Not myArray) = -1  (Not Not myArray) = 0
If ((Not myArray) = -1) Then
End If

'是否為數字 - 結果為BOOLEAN
a = IsNumeric(Cells(3, "M"))

'型態轉換
a = CStr(123)

'CBool   Boolean    Any valid string or numeric expression.
'CByte   Byte       0 to 255.
'CCur    Currency   -922,337,203,685,477.5808 to 922,337,203,685,477.5807.
'CDate   Date       Any valid date expression.
'CDbl    Double     -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values.
'CDec    Decimal    79,228,162,514,264,337,593,543,950,335 for zero-scaled numbers, that is, numbers with no decimal places. For numbers with 28 decimal places, the range is 7.9228162514264337593543950335. The smallest possible non-zero number is 0.0000000000000000000000000001.
'CInt    Integer    -32,768 to 32,767; fractions are rounded(分數會四捨五入).
'CLng    Long       -2,147,483,648 to 2,147,483,647; fractions are rounded.
'CLngLng LongLong   -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807; fractions are rounded. (Valid on 64-bit platforms only.)
'CLngPtr LongPtr    -2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems; fractions are rounded for 32-bit and 64-bit systems.
'CSng    Single     -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values.
'CStr    String     Returns for CStr depend on the expression argument.
'CVar    Variant    Same range as Double for numerics. Same range as String for non-numerics.

'計算餘數
a = 132 Mod 20

'是偶數(even)或奇數(odd)
If (a Mod 2 = 0) Then
    '是偶數
End If

'切割字串
Dim a As Variant
a = Split("1,2,3", ",")

'將不可列印字元(譬如alt+Enter換行)刪除
a = Application.WorksheetFunction.Clean(Range("A1"))

'將字串轉為大寫
a = UCase("abc")

'以cell寫程式，但將該儲存格名子以"A1"形式抓出
a = Cells(1, 1).address(RowAbsolute:=False, ColumnAbsolute:=False)

'將字串中的字替換掉
a = Replace("1,2,3", ",", "")
End Sub

Sub 雜項()
'定義物件為工作表
Dim sysSht As Worksheet
Set sysSht = Worksheets("SYSTEM")
'清除array內容
Erase threeDimArray, twoDimArray
End Sub


