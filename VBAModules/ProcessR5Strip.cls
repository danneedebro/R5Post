VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessR5Strip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iProcess

'
' MEMBER VARIABLES
Public FilesBefore As CollectionFileList    ' (Tracked) Files before process
Public FilesAfter As CollectionFileList     ' (Tracked) Files after process
Public FilesRequired As CollectionFileList  '  Files required to run the process (but not created during run)
Private pBase As ProcessBase
'
'
' INTERFACE PROPERTIES AND METHODS
Private Property Get iProcess_UpToDate() As Boolean
    iProcess_UpToDate = Me.UpToDate
End Property

Public Property Get iProcess_FilesRequired() As CollectionFileList
    Set iProcess_FilesRequired = Me.FilesRequired
End Property

Public Property Get iProcess_FilesBefore() As CollectionFileList
    Set iProcess_FilesBefore = Me.FilesBefore
End Property

Public Property Get iProcess_FilesAfter() As CollectionFileList
    Set iProcess_FilesAfter = Me.FilesAfter
End Property

Private Function iProcess_GetShellCommand(Optional WorkingFolder As String = "") As String
    iProcess_GetShellCommand = Me.GetShellCommand(WorkingFolder)
End Function
'
'
' PROPERTIES
Public Property Get UpToDate() As Boolean
' Action: Returns true if process is up to date
'         All files in the final state should be present AND the date of some
'         outputfile should be greater than some input file
'
    UpToDate = pBase.UpToDate
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set FilesBefore = New CollectionFileList
    Set FilesAfter = New CollectionFileList
    Set FilesRequired = New CollectionFileList
    
    Set pBase = New ProcessBase
    pBase.Create FilesRequired, FilesBefore, FilesAfter
End Sub
'
'
' CLASS METHODS
Public Sub Create(R5Path As R5PostFileObject, StripFile As R5PostFileObject, Rstfile As R5PostFileObject, Strfile As R5PostFileObject)
' Action: Creates the current process given different inputs
'
    FilesRequired.Add R5Path
    
    FilesBefore.Add StripFile
    FilesBefore.Add Rstfile
    FilesAfter.Add Strfile
    Dim i As Integer
    For i = 1 To FilesBefore.Count
        FilesAfter.Add FilesBefore(i)
    Next i
    
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'
    Dim Shellstring As String
    
    Dim s As New ResourceSprintf
    
    Dim RelapExe As String
    RelapExe = FilesRequired(1).FullPath
        
    Dim Striprequestfile As String, Rstfile As String, Strfile As String
    Striprequestfile = FilesBefore(1).GetRelativePath(WorkingFolder)
    Rstfile = FilesBefore(2).GetRelativePath(WorkingFolder)
    Strfile = FilesAfter(1).GetRelativePath(WorkingFolder)
    
    
    Shellstring = s.sprintf("del outdta /Q && del ""%s"" /Q && ", Strfile)
    Shellstring = Shellstring & s.sprintf("""%s"" -i ""%s"" -r ""%s"" -s ""%s""", RelapExe, Striprequestfile, Rstfile, Strfile)
    
    
    'Shellstring = "del outdta /Q && del " & Chr(34) & FilesAfter(1).GetRelativePath(WorkingFolder) & Chr(34) & " /Q" & " && "
    'Shellstring = Shellstring & """" & FilesRequired(1).FullPath & """ -i " & Chr(34) & FilesBefore(1).GetRelativePath(WorkingFolder) & Chr(34) & _
                  " -r " & Chr(34) & FilesBefore(2).GetRelativePath(WorkingFolder) & Chr(34) & _
                  " -s " & Chr(34) & FilesAfter(1).GetRelativePath(WorkingFolder) & Chr(34)
    GetShellCommand = Shellstring
End Function

