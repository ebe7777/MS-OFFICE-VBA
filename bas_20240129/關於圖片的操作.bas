Attribute VB_Name = "����Ϥ����ާ@"
Sub ���J�Ϥ�_��Ʈw()
ActiveSheet.Pictures.Insert("D:\1.png").Select
End Sub
Sub ����Ϥ���T_��Ʈw()
'���J���Ϥ���ؤo�M�C��,�d�e�����O�ۦP��
'   ���n�S�O�`�N�A�C���d�e�ä����M�O�ηƹ��ާ@�ݨ쪺�ȡA���ΤU�C�d�Ҩ��o


Dim myCell As Range
Dim cellWidth As Double, cellHeight As Double
Dim cellMergeAreaWidth As Double, cellMergeAreaHeight As Double
Dim shape As Excel.shape
'======���J�Ϥ��èϤ��P �x�s�� / �x�s��Ҧb���X�ֽd�� �P���P�e
'���o�x�s��ؤo
Set myCell = Cells(2, 2)
myCell.Select
'   �x�s�楻��
cellHeight = myCell.Height
cellWidth = myCell.width
'   �x�s��Ҧb���X�ֽd��
cellMergeAreaHeight = myCell.MergeArea.Height
cellMergeAreaWidth = myCell.MergeArea.width
Debug.Print "hei=" & cellHeight & ",wid=" & cellWidth
Debug.Print "hei=" & cellMergeAreaHeight & ",wid=" & cellMergeAreaWidth

'====�H�s���覡���J�Ϥ� (���ɲ���EXCEL�N�|��ܥ]�l��)
'   ���J�Ϥ��b��w�s�x��(�Ϥ����W���|�۰ʻP�x�s����;�p�G�s�x��b�@�ӦX�ֽd�򤺡A�h�۰ʻP�X�ֽd�򥪤W�����)
ActiveSheet.Pictures.Insert("D:\1.png").Select
'   �]�w �Ϥ� ����������
Selection.ShapeRange.LockAspectRatio = msoFalse
'   �]�w�Ϥ��ؤo
Selection.ShapeRange.Height = cellHeight
Selection.ShapeRange.width = cellWidth
Selection.ShapeRange.Height = cellMergeAreaHeight
Selection.ShapeRange.width = cellMergeAreaWidth

'====�H���O�覡���J�Ϥ�
Set myShape = ActiveSheet.Shapes.AddPicture(Filename:="D:\1.png", linktofile:=msoFalse, savewithdocument:=msoCTrue, Left:=10, Top:=20, width:=50, Height:=50)

'�R���Ҧ��u�@��W���Ϥ�
For Each shape In ActiveSheet.Shapes
    Debug.Print shape.Name
    shape.Delete
Next

End Sub

Sub �����Ϥ��ɭ�l�ɮת��ؤo()

Dim objShell As Object, objFolder As Object, objFile As Object
Dim objDim As String, objWidth As Long, objHeig As Long

Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace("C:\�J���H���t��\01�Ϥ����\L123456789")
Set objFile = objFolder.ParseName("L123456789_�N�ثH_H02.jpg")
objDim = objFile.ExtendedProperty("Dimensions")
objWidth = Mid(objDim, 2, InStr(1, objDim, "x") - 2)
objHeig = Mid(objDim, InStr(1, objDim, "x") + 1, Len(objDim) - InStr(1, objDim, "x") - 1)

End Sub

Sub �����Ϥ��ɭ�l�ɮת��ؤo_old()
Dim myPicture As Object
Dim filePath As String
Dim imageH As Long, imageW As Long
filePath = "C:\123.jpg"
'���o�s�Ϥؤo
Set myPicture = CreateObject("WIA.ImageFile")
myPicture.LoadFile filePath
'�`�N!�γ\�O����k�O�L�n���¤�k�A�p�G�O��win10���k�����A����k�|�{���èS������
'�p�G�O��office 2010 picture manager(OPM)����᪺���P�e�A����k�N�i���Ѫ��X��
imageH = myPicture.Height
imageW = myPicture.width
End Sub
