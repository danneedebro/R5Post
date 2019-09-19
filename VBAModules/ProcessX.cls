VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessX"
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
Public Sub Create(RunFilesRequired As CollectionFileList, FilesBeforeProcess As CollectionFileList, FilesAfterProcess As CollectionFileList)
' Action: Creates the current process given different inputs
'
    Set FilesRequired = RunFilesRequired
    Set FilesBefore = FilesBeforeProcess
    Set FilesAfter = FilesAfterProcess
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'
    
    Dim Shellstring As String
    Dim i As Integer
    
    
    
    Shellstring = "USES: "
    For i = 1 To FilesRequired.Count
        Shellstring = Shellstring & FilesRequired(i).GetRelativePath(WorkingFolder) & IIf(i = FilesRequired.Count, vbNewLine, ", ")
    Next i
    
    Shellstring = Shellstring & "Before: "
    For i = 1 To FilesBefore.Count
        Shellstring = Shellstring & FilesBefore(i).GetRelativePath(WorkingFolder) & IIf(i = FilesBefore.Count, vbNewLine, ", ")
    Next i
    
    Shellstring = Shellstring & "After: "
    For i = 1 To FilesAfter.Count
        Shellstring = Shellstring & FilesAfter(i).GetRelativePath(WorkingFolder) & IIf(i = FilesAfter.Count, vbNewLine, ", ")
    Next i
    
    GetShellCommand = Shellstring
    ' cmd /c cd "H:\Projekt\Småjobb\2018\R2 411 Böjda rör pilotventiler MSIV\Relapmodell\Calc\Case1\" && H: && cd && timeout 2 && "H:\RELAP5\bin\relap5\relap5mod3.3patch03\relap5.exe" "-i" "Case1.i" "-O" "Case1.o" "-R" "Case1.rst" "-Z" "H:\RELAP5\bin\relap5\relap5mod3.3patch03\tpfh2onew" "-w" "H:\RELAP5\bin\relap5\relap5mod3.3patch03\tpfh2o" && timeout 2 && cp screen Case1.i.screen -f && del "Case1.str" /Q && timeout 1 && "H:\RELAP5\bin\relap5\relap5mod3.3patch03\relap5.exe" "-i" "..\Stripfile.strip" "-r" "Case1.rst" "-s" "Case1.str" "-O" "outdta" && timeout 2 && "C:\Program Files (x86)\MATLAB\R2012a\bin\matlab.exe" -nodisplay -nosplash -nodesktop -r "addpath 'H:\RELAP5\Script\plottaStripFil\';plottaStripFil('batch','-s','Case1.str','-t','Case1.ps','-p','..\Stripfile.strip.txt','tmin',10,'tmax',16,'title','Case 1 (Grundfall, 70 bar a)','tsamp',1);pause(3);exit" && "C:\Users\dero\AptPlot\bin\r2dmx.exe" "-r" "Case1.rst" "-d" "Case1.dmx" "-m" "512" && timeout 2

End Function
