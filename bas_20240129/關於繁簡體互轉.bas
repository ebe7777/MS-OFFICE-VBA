Attribute VB_Name = "�����c²�餬��"
'�ѦҺ���
'https://www.itread01.com/content/1543804631.html

Private Declare PtrSafe Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, _
ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Private Declare PtrSafe Function lStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'�ഫ�禡�A0=²���c�A1=�c��²
Function Jian_Fan_Conv(ByVal strString As String, Optional ByVal iMode As Integer = 0) As String
    Dim lStrLength As Long
    Dim strNew As String
    Const J2F_MAPFLAG = &H4000000
    Const F2J_MAPFLAG = &H2000000
    Jian_Fan_Conv = ""
    lStrLength = lStrLen(strString)
    strNew = Space(lStrLength)
    '0²���c
    If iMode = 0 Then
        LCMapString &H804, J2F_MAPFLAG, strString, lStrLength, strNew, lStrLength
    '1�c��²
    ElseIf iMode = 1 Then
        LCMapString &H804, F2J_MAPFLAG, strString, lStrLength, strNew, lStrLength
    End If
    Jian_Fan_Conv = strNew
End Function

Sub j2f()
    sj = Range("c8").Value
    sf = Jian_Fan_Conv(sj)
    MsgBox sj & sf
'    sx = Jian_Fan_Conv(sj, 1)
'    sy = Jian_Fan_Conv(sf, 1)
'    MsgBox sx & sy

End Sub
