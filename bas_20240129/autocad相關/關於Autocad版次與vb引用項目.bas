Attribute VB_Name = "����Autocad�����Pvb�ޥζ���"
'�����]�w�ޥζ��� "Microsoft Visual Basic for Applications Extensibility x.x"
Private Sub LoadRef()
  Dim obj As Object
  Dim guid$
  
  On Error Resume Next
  Set obj = GetObject(, "AutoCAD.Application")
  
  obj.Visible = True
  
  Call ClearAcadRef
  
  '2017-12-08 Update
  Select Case Left(obj.Version, 2)
    '2007
    Case "17"
      guid = "{851A4561-F4EC-4631-9B0C-E7DC407512C9}"
    '2010
    Case "18"
      'guid = "{D32C213D-6096-40EF-A216-89A3A6FB82F7}" '32bits
      guid = "{E072BCE4-9027-4F86-BAE2-EF119FD0A0D3}" '64bits
    '2014
    Case "19"
      'guid = "{852B2D4E-B1F4-4BD6-8672-9993177C1A40}" '32bit
      guid = "{D5C3CB6F-AA0A-4D45-B02D-CF2974EFD4BE}" '64bits
    '2015,2016
    Case "20"
      guid = "{4E3F492A-FB57-4439-9BF0-1567ED84A3A9}" '64bits
    '2017
    Case "21"
      guid = "{5B3245BE-661C-4324-BB55-3AD94EBBFDD7}" '64bits
    '2018
    Case "22"
      guid = "{644614D2-93DC-48C6-A061-21ABCE65A4C0}" '64bits
  End Select
  Application.VBE.ActiveVBProject.References.AddFromGuid guid, 1, 0
  'Application.VBE.ActiveVBProject.References.AddFromFile _
  '  "C:\Program Files\Common Files\Autodesk Shared\acax" & Left(obj.Version, 2) & "enu.tlb"
End Sub

Private Sub ClearAcadRef()
'�H�U��dim�k�����]�w�ޥζ��� "Microsoft Visual Basic for Applications Extensibility x.x"
'Dim ref As VBIDE.Reference
'Dim refs As VBIDE.References
'�H�U��dim�k����
Dim ref As Object
Dim refs As Object
'�`�N�A�b�ާ@reference�ɡA�p�G�]�w���_�I�A�{�����M�|���_�A���@�w�|�X�{���~�T��[���ɵL�k�i�J���_�Ҧ�]�F�L���z�|�ӿ��~�T��
    Set refs = Application.VBE.ActiveVBProject.References
    For Each ref In refs
        If ref.name = "AutoCAD" Then
            Call refs.Remove(ref)
        End If
    Next
End Sub
Private Sub ListProjectReferencesList()
    Dim i                   As Long
    Dim VBProj              As Object  'VBIDE.VBProject
    Dim VBComp              As Object 'VBIDE.VBComponent
    Set VBProj = Application.VBE.ActiveVBProject
    Dim strTmp              As String
    On Error Resume Next
    For i = 1 To VBProj.References.Count
        With VBProj.References.Item(i)
            Debug.Print "Description: " & .Description & vbNewLine & _
                        "FullPath: " & .FullPath & vbNewLine & _
                        "Major.Minor: " & .Major & "." & .Minor & vbNewLine & _
                        "Name: " & .name & vbNewLine & _
                        "GUID: " & .guid & vbNewLine & _
                        "Type: " & .Type
            Debug.Print "-------------------"
        End With 'VBProj.References.Item(i)
    Next i
End Sub
Sub test()
Dim i As String
For i = 1 To UBound(myArray, 1)
Next i
End Sub
