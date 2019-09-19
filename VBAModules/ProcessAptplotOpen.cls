VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessAptplotOpen"
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
Private pFileType As String  ' restart, demux or strip
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
Public Sub Create(AptplotPath As R5PostFileObject, Rstfile As R5PostFileObject, FileType As String)
' Action: Creates the current process given different inputs
'
    FilesRequired.Add AptplotPath
    FilesBefore.Add Rstfile
    FilesAfter.Add Rstfile
    pFileType = FileType
    
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'
    Dim Shellstring As String
    Dim s As New ResourceSprintf
    
    Dim AptplotExe As String
    Dim Rstfile As String
    AptplotExe = FilesRequired(1).FullPath
    
    Rstfile = FilesBefore(1).GetRelativePath(WorkingFolder)
    
    'Print #1, "relap 0 " + FileType + " " + Chr(34) + FullPath + Chr(34)
    'Print #1, "relap 0 show"
    'Close #1
    
    Shellstring = s.sprintf("ECHO relap 0 %s ""%s"" > temp.par && ECHO relap 0 show >> temp.par && ", pFileType, Rstfile)
    Shellstring = Shellstring & s.sprintf("""%s"" -param temp.par &", AptplotExe)
    
    GetShellCommand = Shellstring

End Function





