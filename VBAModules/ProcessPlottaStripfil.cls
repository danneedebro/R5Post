VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessPlottaStripfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements interfaceProcess

'
' MEMBER VARIABLES
Public FilesetInitial As CollectionFileList   ' (Tracked) Files before process
Public FilesetFinal As CollectionFileList     ' (Tracked) Files after process
Public FilesetRequired As CollectionFileList  ' Files required to run the process (but not created during run)
Private PlotTimeMin As Double
Private PlotTimeMax As Double
Private PlotTitle As String


'
'
' INTERFACE PROPERTIES AND METHODS
Private Property Get interfaceProcess_UpToDate() As Boolean
    interfaceProcess_UpToDate = Me.UpToDate
End Property

Public Property Get interfaceProcess_FilesetRequired() As CollectionFileList
    Set interfaceProcess_FilesetRequired = Me.FilesetRequired
End Property

Public Property Get interfaceProcess_FilesetInitial() As CollectionFileList
    Set interfaceProcess_FilesetInitial = Me.FilesetInitial
End Property

Public Property Get interfaceProcess_FilesetFinal() As CollectionFileList
    Set interfaceProcess_FilesetFinal = Me.FilesetFinal
End Property

Private Function interfaceProcess_GetShellCommand(Optional WorkingFolder As String = "") As String
    interfaceProcess_GetShellCommand = Me.GetShellCommand(WorkingFolder)
End Function
'
'
' PROPERTIES


Public Property Get UpToDate() As Boolean
' Action: Returns true if process is up to date
'         All files in the final state should be present AND the date of some
'         outputfile should be greater than some input file
'
    Dim i As Integer
    Dim AllFilesExists As Boolean
    AllFilesExists = True
    
    ' Check if all files that are supposed to be created is created
    For i = 1 To FilesetFinal.Count
        If FilesetFinal(i).FileExists = False Then AllFilesExists = False
    Next i
    
    ' If all files created, check date of one of the output files to see if it's created after the input file
    If AllFilesExists = True Then
        
        If FilesetFinal(1).DateLastModified > FilesetInitial(1).DateLastModified Then
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
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set FilesetInitial = New CollectionFileList
    Set FilesetFinal = New CollectionFileList
    Set FilesetRequired = New CollectionFileList
End Sub
'
'
' CLASS METHODS
Public Sub Create(MatlabPath As R5PostFileObject, MatlabScriptPath As R5PostFileObject, strFile As R5PostFileObject, Paramfile As R5PostFileObject, _
                  postScriptFile As R5PostFileObject, Title As String, tMin As Double, tMax As Double)
' Action: Creates the current process given different inputs
'
    FilesetRequired.Add MatlabPath
    FilesetRequired.Add MatlabScriptPath
    FilesetInitial.Add strFile
    FilesetInitial.Add Paramfile
    FilesetFinal.Add postScriptFile
    PlotTimeMin = tMin
    PlotTimeMax = tMax
    PlotTitle = Title
    
    
    Dim i As Integer
    For i = 1 To FilesetInitial.Count
        FilesetFinal.Add FilesetInitial(i)
    Next i
    
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'

    Dim Shellstring As String
    
    Dim strFilePath As String, psFilePath As String
    
    strFilePath = FilesetFinal(2).getRelativePath(WorkingFolder)
    psFilePath = FilesetFinal(1).getRelativePath(WorkingFolder)
    
    Shellstring = """" & FilesetRequired(1).FullPath & """ -nodisplay -nosplash -nodesktop -r ""addpath " & FilesetRequired(2).FolderPath & ";plottaStripFil('batch','-s','" & strFilePath & "','-t','" & psFilePath & "','tmin'," & PlotTimeMin & ",'tmax'," & PlotTimeMax & ",'title','" & PlotTitle & "','tsamp',1);pause(3);exit"""
    GetShellCommand = Shellstring
    ' matlab -nodisplay -nosplash -nodesktop -r "addpath C:\RELAP5\Script\plottaStripFil;plottaStripFil('batch','-s','$1.str','-t','$1.ps','-m','$1_MaxMin.ps','-xls','$1_MaxMin.xls','tmin',$2,'tmax',$3,'title','$4','tsamp',1);pause(3);exit"


End Function
'
'
' STATIC METHODS






