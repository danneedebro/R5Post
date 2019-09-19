VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pFilesBefore As CollectionFileList    ' (Tracked) Files before process
Private pFilesAfter As CollectionFileList     ' (Tracked) Files after process
Private pFilesRequired As CollectionFileList  '  Files required to run the process (but not created during run)

Public Sub Create(FilesRequired As CollectionFileList, FilesBefore As CollectionFileList, FilesAfter As CollectionFileList)
    Set pFilesRequired = FilesRequired
    Set pFilesBefore = FilesBefore
    Set pFilesAfter = FilesAfter
End Sub

Public Property Get UpToDate() As Boolean
' Action: Returns true if process is up to date
'         All files in the final state should be present AND the date of some
'         outputfile should be greater than some input file
'
    Dim i As Integer
    Dim AllFilesExists As Boolean
    AllFilesExists = True
    
    ' Check if all files that are supposed to be created is created
    For i = 1 To pFilesAfter.Count
        If pFilesAfter(i).FileExists = False Then AllFilesExists = False
    Next i
    
    ' If all files created, check date of one of the output files to see if it's created after the input file
    If AllFilesExists = True Then
        
        If pFilesAfter(1).DateLastModified > pFilesBefore(1).DateLastModified Then
            Debug.Print "All files exists and is up to date"
            UpToDate = True
        Else
            Debug.Print "All files exists but not up to date"
            UpToDate = False
        End If
    Else
        Debug.Print "Not all files exists"
        UpToDate = False
    End If
End Property
