VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessR2DMX"
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
Public Sub Create(R2DMXPath As R5PostFileObject, Rstfile As R5PostFileObject, dmxFile As R5PostFileObject)
' Action: Creates the current process given different inputs
'
    FilesRequired.Add R2DMXPath
    FilesBefore.Add Rstfile
    FilesAfter.Add Rstfile
    FilesAfter.Add dmxFile
    
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'
    Dim Shellstring As String
    Dim s As New ResourceSprintf
    
    Dim rstFilePath As String
    Dim dmxFilePath As String
    rstFilePath = FilesBefore(1).GetRelativePath(WorkingFolder)
    dmxFilePath = FilesAfter(2).GetRelativePath(WorkingFolder)
    
    Shellstring = s.sprintf("""%s"" -r ""%s"" -d ""%s"" -v -m 512", FilesRequired(1).FullPath, rstFilePath, dmxFilePath)
    GetShellCommand = Shellstring

End Function
'
'
' STATIC METHODS



