VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessPs2Pdf"
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
Public Sub Create(GhostScriptPath As R5PostFileObject, psFile As R5PostFileObject, pdfFile As R5PostFileObject)
' Action: Creates the current process given different inputs
'
    FilesetRequired.Add GhostScriptPath
    FilesetInitial.Add psFile
    FilesetFinal.Add psFile
    FilesetFinal.Add pdfFile
    
End Sub

Public Function GetShellCommand(Optional WorkingFolder As String) As String
' Action: Returns the shell command for the current process. File paths are written
'         as a relative path (relative to WorkingFolder)
'
'
    
    Dim Shellstring As String
    
    Dim pdfFilePath As String
    Dim psFilePath As String
    pdfFilePath = FilesetFinal(2).getRelativePath(WorkingFolder)
    psFilePath = FilesetFinal(1).getRelativePath(WorkingFolder)
    
    Shellstring = """" & FilesetRequired(1).FullPath & """ -sDEVICE=pdfwrite -o """ & pdfFilePath & """ """ & psFilePath & """"
    GetShellCommand = Shellstring
    ' cmd /c cd "H:\Projekt\Småjobb\2018\R2 411 Böjda rör pilotventiler MSIV\Relapmodell\Calc\Case1\" && H: && cd && timeout 2 && "H:\RELAP5\bin\relap5\relap5mod3.3patch03\relap5.exe" "-i" "Case1.i" "-O" "Case1.o" "-R" "Case1.rst" "-Z" "H:\RELAP5\bin\relap5\relap5mod3.3patch03\tpfh2onew" "-w" "H:\RELAP5\bin\relap5\relap5mod3.3patch03\tpfh2o" && timeout 2 && cp screen Case1.i.screen -f && del "Case1.str" /Q && timeout 1 && "H:\RELAP5\bin\relap5\relap5mod3.3patch03\relap5.exe" "-i" "..\Stripfile.strip" "-r" "Case1.rst" "-s" "Case1.str" "-O" "outdta" && timeout 2 && "C:\Program Files (x86)\MATLAB\R2012a\bin\matlab.exe" -nodisplay -nosplash -nodesktop -r "addpath 'H:\RELAP5\Script\plottaStripFil\';plottaStripFil('batch','-s','Case1.str','-t','Case1.ps','-p','..\Stripfile.strip.txt','tmin',10,'tmax',16,'title','Case 1 (Grundfall, 70 bar a)','tsamp',1);pause(3);exit" && "C:\Users\dero\AptPlot\bin\r2dmx.exe" "-r" "Case1.rst" "-d" "Case1.dmx" "-m" "512" && timeout 2

End Function
'
'
' STATIC METHODS




