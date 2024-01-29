Attribute VB_Name = "���]�榡�d��"
Sub resetForm_��Ʈw()
Dim mainSN As String
Dim dataStartRow As Integer, dataEndRow As Integer

mainSN = "CvKv�p���"
dataStartRow = 2
dataEndRow = 500

Application.ScreenUpdating = False

    Sheets(mainSN).Select


'�]�^�w�]��======================
'����m��
    Range("A:I").Select
    Selection.UnMerge
'�U�Ԧ����
    Cells.Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
'���e============================


'�r�C��
    '�{���۰ʶ�W����
    Range("A:J").Select
    With Selection.Font
        .Color = -4165632
        .TintAndShade = 0
    End With
    '�ϥΪ̤�ʿ�J����
    Range("B:G").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    '�ϥΪ̨ϥΤU�Ԧ�����J��
    Range("C:C,E:E,G:G").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    

    
'�r��
    
    '����
   Range("A:I").Select
    With Selection.Font
        .Name = "�s�ө���"
        .Size = 12
    End With
    '���D�C
    Rows("1:1").Select
    With Selection.Font
        .Name = "�з���"
        .Size = 12
    End With
'�r���
    '�m��
    Range("A:A,C:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    '�a��
    Range("B:B,E:E,G:G").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    '�a�k
    Range("D:D,F:F,H:I").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    '����P�Y�p
    Cells.Select
    With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
'��e�C���۰ʽվ�
    Cells.EntireColumn.AutoFit
    
'��e
    Columns("A:D").ColumnWidth = 24
    Cells.Select

    
 '�ؽu
    '����
    Range("A:J").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    '�S��B�z
    Columns("D:E").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("F:G").Select
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'�U�Ԧ����
    Range("C:C").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="15A,20A,25A,32A,40A,50A,65A,80A,100A,125A"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���~"
        .InputMessage = ""
        .ErrorMessage = "��J�ȥ����P�U�Ԧ����ۦP"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("E:E").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="GPM,LPM,LPS,m3/hr"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���~"
        .InputMessage = ""
        .ErrorMessage = "��J�ȥ����P�U�Ԧ����ۦP"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("G:G").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="psi,ft-W.G.,M-W.G.,kPa,kg/cm2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���~"
        .InputMessage = ""
        .ErrorMessage = "��J�ȥ����P�U�Ԧ����ۦP"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    
'���D============================

'�r���e
    i = 1
    Cells(1, i) = "����"
    i = i + 1
    Cells(1, i) = "TAG NAME"
    i = i + 1
    Cells(1, i) = "�޸��ؤo"
    i = i + 1
    Cells(1, i) = "Q�y�q"
    i = i + 1
    Cells(1, i) = ""
    i = i + 1
    Cells(1, i) = "��P���t"
    i = i + 1
    Cells(1, i) = ""
    i = i + 1
    Cells(1, i) = "Cv"
    i = i + 1
    Cells(1, i) = "Kv"
    i = i + 1
    Cells(1, i) = "<����>"
'����m��
    Range("D1:E1").Select
    Selection.Merge
    Range("F1:G1").Select
    Selection.Merge
'�r���
    '�m��
    Range("A1:J1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
'�r�C��
    Rows("1:1").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
'�x�s���
    Range("A1:I1").Select
    With Selection.Interior
        .Color = 5296274
    End With
    Range("J1:J1").Select
    With Selection.Interior
        .Color = 65535
    End With
    
'�U�Ԧ����(�M��)
    Range("A1:J1").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
Application.ScreenUpdating = True
End Sub

