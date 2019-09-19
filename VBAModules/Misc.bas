Attribute VB_Name = "Misc"
Function changeFileExtension(relativePath, newFileExtension)
If Len(relativePath) > 0 Then
    changeFileExtension = Left(relativePath, Len(relativePath) - 2) + newFileExtension
Else
    changeFileExtension = ""
End If
End Function



Private Sub fixLinks_old()
    ' Fixar länkar. Markera de celler som ska fixas och kör makrot
    Dim currRow, currCol, currRowCnt
    
    Selection.Hyperlinks.Delete
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
    currRow = Selection.Row
    currCol = Selection.Column
    currRowCnt = Selection.Rows.Count
    
    For i = 1 To currRowCnt
        Cells(currRow + i - 1, currCol).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=ActiveSheet.Name & "!" & Selection.Address
        MsgBox i
    Next i
End Sub

