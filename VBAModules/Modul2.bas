Attribute VB_Name = "Modul2"
Sub GetFiles()
    Dim fso As New FileSystemObject
    Dim myFolder As Folder
    
    Set myFolder = fso.GetFolder(ThisWorkbook.Path)
    
    'if "R5Post_v1.0beta.xlsm" is in myFolder.Files then
    
    
    Debug.Print myFolder.Files("R5Post_v1.0beta.xlsm").DateLastModified
    
    Dim objFile As File
    For Each objFile In myFolder.Files
        Debug.Print objFile.Name & ": " & objFile.DateLastModified
    Next objFile
    
    


End Sub




Sub FFF()

    Dim fso As New FileSystemObject
    Dim myFolder As Folder
    
    Debug.Print fso.GetFileName("Hej\Apa.txt")
    
    

End Sub
