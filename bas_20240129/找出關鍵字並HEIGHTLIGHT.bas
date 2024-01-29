Attribute VB_Name = "找出關鍵字並HEIGHTLIGHT"

' from 研究生2.0
' http://newgenerationresearcher.blogspot.com
' by 苦情研究生
' Last update: 2009/01/30

Sub 改變關鍵字顏色_資料庫()

Dim strName As String
    
    strName = InputBox(Prompt:="Your keyword please.", Title:="Type a word you want to highlight", Default:="Type a word you want to highlight")


        If strName = "Type a word you want to highlight" Or strName = vbNullString Then

           Exit Sub

        Else

           doit (strName)
           
        End If
 

End Sub

Sub doit(keywords)
    
    Dim vntWords As Variant
    Dim lngIndex As Long
    Dim rngFind As Range
    Dim strFirstAddress As String
    Dim lngPos As Long
    
    vntWords = keywords
    
    With ActiveSheet.UsedRange
            
            Set rngFind = .Find(vntWords, LookIn:=xlValues, LookAt:=xlPart)
            If Not rngFind Is Nothing Then
                strFirstAddress = rngFind.Address
                Do
                    lngPos = 0
                    Do
                        lngPos = InStr(lngPos + 1, rngFind.Value, vntWords, vbTextCompare)
                        If lngPos > 0 Then
                            With rngFind.Characters(lngPos, Len(vntWords))
                                .Font.Bold = True
                                .Font.Size = .Font.Size + 2
                                .Font.ColorIndex = 3
                            End With
                        End If
                    Loop While lngPos > 0
                    Set rngFind = .FindNext(rngFind)
                Loop While rngFind.Address <> strFirstAddress
            End If

    End With
    
End Sub

