Attribute VB_Name = "關於WORD引用項目"
'取得與設定教學 https://www.wiseowl.co.uk/blog/s204/vbe-references.htm
'官方說明 https://docs.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/addfromfile-method-vba-add-in-object-model


Sub 目前已經設定的引用項目的資訊()
'用來偵測各版本 "Microsoft Word XX.0 Object Library"的GUID
'先引用該項目,然後執行此程式,打開 區域變數視窗 看 "ref"內的資訊

'create a variable to refer to each reference

Dim ref As Reference

'list out all of the current references

For Each ref In Application.VBE.ActiveVBProject.References

Debug.Print ref.Name, ref.Description

Next ref

End Sub

Sub 偵測目前打開的word的版本()
Dim wordApp As Object, wordDoc As Object, wordVer As String
Dim wordSampleFileFullPath As String

wordSampleFileFullPath = "C:\123\123.docx"

Set wordApp = New Word.Application
Set wordDoc = wordApp.Documents.Open(fileName:=wordSampleFileFullPath, ReadOnly:=True)

wordVer = wordDoc.Application.Version
End Sub
Sub 相關資訊資料庫()
'Office Word 2016
'Application.Version取得值: "16.0" String
'引用項目名稱: "Microsoft Word 16.0 Object Library" String
'GUID: "{00020905-0000-0000-C000-000000000046}" String

'Office Word 2013
'Application.Version取得值: "15.0" String
'引用項目名稱: "Microsoft Word 15.0 Object Library" String
'GUID: "" String
End Sub

Sub OpenWordDocsFromExcelLateBinding()

    ' Declare objects
    Dim wrdApplication As Object
    Dim wrdDocument As Object

    ' Declare other variables
    Dim wrdDocumentFullPath As String
    Dim wrdDocumentName As String
    Dim documentCounter As Integer

    ' Check if Word is already opened
    On Error Resume Next

    Set wrdApplication = GetObject(, "Word.Application")

    If Err.Number <> 0 Then
        ' Open a new instance
        Set wrdApplication = CreateObject("Word.Application")
        wrdApplication.Visible = True
    End If

    ' Reset error handling
    Err.Clear
    On Error GoTo 0

    ' Open file dialog
    With Application.FileDialog(1)  'msoFileDialogOpen
        .AllowMultiSelect = True
        .Show

        'Set wrdApplication = New Word.Application
        documentCounter = .SelectedItems.Count

        ' For each document selected in dialog
        For documentCounter = 1 To .SelectedItems.Count

            ' Get full path and name of each file selected
            wrdDocumentFullPath = .SelectedItems(documentCounter)
            wrdDocumentName = Mid(.SelectedItems(documentCounter), InStrRev(.SelectedItems(documentCounter), "\") + 1)

            ' Check if document is already opened
            On Error Resume Next

            Set wrdDocument = wrdApplication.Documents(wrdDocumentName)

            If Err.Number <> 0 Then
                ' Open word document
                Set wrdDocument = wrdApplication.Documents.Open(wrdDocumentFullPath)
            End If

            ' Reset error handling
            Err.Clear
            On Error GoTo 0

        Next documentCounter

    End With


End Sub
