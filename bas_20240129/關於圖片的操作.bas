Attribute VB_Name = "關於圖片的操作"
Sub 插入圖片_資料庫()
ActiveSheet.Pictures.Insert("D:\1.png").Select
End Sub
Sub 關於圖片資訊_資料庫()
'插入的圖片其尺寸和列高,攔寬的單位是相同的
'   但要特別注意，列高攔寬並不全然是用滑鼠操作看到的值，須用下列範例取得


Dim myCell As Range
Dim cellWidth As Double, cellHeight As Double
Dim cellMergeAreaWidth As Double, cellMergeAreaHeight As Double
Dim shape As Excel.shape
'======插入圖片並使之與 儲存格 / 儲存格所在的合併範圍 同高同寬
'取得儲存格尺寸
Set myCell = Cells(2, 2)
myCell.Select
'   儲存格本身
cellHeight = myCell.Height
cellWidth = myCell.width
'   儲存格所在的合併範圍
cellMergeAreaHeight = myCell.MergeArea.Height
cellMergeAreaWidth = myCell.MergeArea.width
Debug.Print "hei=" & cellHeight & ",wid=" & cellWidth
Debug.Print "hei=" & cellMergeAreaHeight & ",wid=" & cellMergeAreaWidth

'====以連結方式插入圖片 (圖檔移除EXCEL就會顯示包子圖)
'   插入圖片在選定存儲格(圖片左上角會自動與儲存格對齊;如果存儲格在一個合併範圍內，則自動與合併範圍左上角對齊)
ActiveSheet.Pictures.Insert("D:\1.png").Select
'   設定 圖片 不等比例顯示
Selection.ShapeRange.LockAspectRatio = msoFalse
'   設定圖片尺寸
Selection.ShapeRange.Height = cellHeight
Selection.ShapeRange.width = cellWidth
Selection.ShapeRange.Height = cellMergeAreaHeight
Selection.ShapeRange.width = cellMergeAreaWidth

'====以內嵌方式插入圖片
Set myShape = ActiveSheet.Shapes.AddPicture(Filename:="D:\1.png", linktofile:=msoFalse, savewithdocument:=msoCTrue, Left:=10, Top:=20, width:=50, Height:=50)

'刪除所有工作表上的圖片
For Each shape In ActiveSheet.Shapes
    Debug.Print shape.Name
    shape.Delete
Next

End Sub

Sub 取的圖片檔原始檔案的尺寸()

Dim objShell As Object, objFolder As Object, objFile As Object
Dim objDim As String, objWidth As Long, objHeig As Long

Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace("C:\入場人員系統\01圖片資料\L123456789")
Set objFile = objFolder.ParseName("L123456789_吉佩信_H02.jpg")
objDim = objFile.ExtendedProperty("Dimensions")
objWidth = Mid(objDim, 2, InStr(1, objDim, "x") - 2)
objHeig = Mid(objDim, InStr(1, objDim, "x") + 1, Len(objDim) - InStr(1, objDim, "x") - 1)

End Sub

Sub 取的圖片檔原始檔案的尺寸_old()
Dim myPicture As Object
Dim filePath As String
Dim imageH As Long, imageW As Long
filePath = "C:\123.jpg"
'取得新圖尺寸
Set myPicture = CreateObject("WIA.ImageFile")
myPicture.LoadFile filePath
'注意!或許是此方法是微軟的舊方法，如果是用win10的右鍵選轉，此方法會認為並沒有旋轉
'如果是用office 2010 picture manager(OPM)旋轉後的高與寬，此方法就可辨識的出來
imageH = myPicture.Height
imageW = myPicture.width
End Sub
