VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessR5Calc"
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
Public Sub Create(R5Path As R5PostFileObject, R5SteamTable As R5PostFileObject, R5SteamTableOld As R5PostFileObject, _
                  iFile As R5PostFileObject, oFile As R5PostFileObject, Rstfile As R5PostFileObject)
' Action: Creates the current process given different inputs
'
    FilesRequired.Add R5Path
    FilesRequired.Add R5SteamTable
    FilesRequired.Add R5SteamTableOld
    FilesBefore.Add iFile
    FilesAfter.Add oFile
    FilesAfter.Add Rstfile
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
    
    Dim RelapExe As String, RelapSteam As String, RelapSteamOld As String
    RelapExe = FilesRequired(1).FullPath
    RelapSteam = FilesRequired(2).FullPath
    RelapSteamOld = FilesRequired(3).FullPath
    
    Dim Inputfile As String, Rstfile As String, Outfile As String
    Inputfile = FilesBefore(1).GetRelativePath(WorkingFolder)
    Outfile = FilesAfter(1).GetRelativePath(WorkingFolder)
    Rstfile = FilesAfter(2).GetRelativePath(WorkingFolder)
    
    Shellstring = s.sprintf("""%s"" -i ""%s"" -O ""%s"" -R ""%s"" -Z ""%s"" -w ""%s""", RelapExe, Inputfile, Outfile, Rstfile, RelapSteam, RelapSteamOld)
    GetShellCommand = Shellstring
End Function
