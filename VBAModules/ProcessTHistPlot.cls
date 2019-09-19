VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessTHistPlot"
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

Private pPlotTimeMin As Double
Private pPlotTimeMax As Double
Private pPlotTitle As String
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
Public Sub Create(MatlabPath As R5PostFileObject, ScriptPath As R5PostFileObject, StripFile As R5PostFileObject, Paramfile As R5PostFileObject, _
                  Psfile As R5PostFileObject, Title As String, tMin As Double, tMax As Double)
' Action: Creates the current process given different inputs
'
    FilesRequired.Add MatlabPath
    FilesRequired.Add ScriptPath
    FilesBefore.Add StripFile
    FilesBefore.Add Paramfile
    FilesAfter.Add Psfile
    pPlotTimeMin = tMin
    pPlotTimeMax = tMax
    pPlotTitle = Title
    
    
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
    Dim s As New ResourceSprintf
    Dim Shellstring As String
    
    Dim strFilePath As String, psFilePath As String, paramFilePath As String
    Dim MatlabPath As String, scriptFolder As String
    MatlabPath = FilesRequired(1).FullPath
    scriptFolder = FilesRequired(2).FolderPath
    
    strFilePath = FilesAfter(2).GetRelativePath(WorkingFolder)
    psFilePath = FilesAfter(1).GetRelativePath(WorkingFolder)
    paramFilePath = FilesBefore(2).GetRelativePath(WorkingFolder)
    
    Shellstring = s.sprintf("""%s"" -nosplash -nodesktop -batch ""addpath %s;THistPlot('DataFile','%s','ParamFile','%s','PlotFile','%s','tmin',%f,'tmax',%f,'Title','%s','tsamp',1);pause(3);exit""", MatlabPath, scriptFolder, strFilePath, paramFilePath, psFilePath, pPlotTimeMin, pPlotTimeMax, pPlotTitle)
    GetShellCommand = Shellstring

End Function











