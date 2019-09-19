VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainProcessChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pProcesses As Collection
Private pLogfiles As Collection
Private pLogfileMessage As Collection
'
'
' PROPERTIES
Public Property Get NumberOfProcesses() As Integer
' Action: Property for number of files in intial file set
    NumberOfProcesses = pProcesses.Count
End Property

Public Property Get Process(ByVal index As Integer) As iProcess
    Set Process = pProcesses(index)
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pProcesses = New Collection
    Set pLogfiles = New Collection
    Set pLogfileMessage = New Collection
End Sub

Public Sub Add(NewProcess As iProcess, Optional Logfile As R5PostFileObject, Optional Message As String)
    pProcesses.Add NewProcess
    If Logfile Is Nothing Then
        pLogfiles.Add New R5PostFileObject
    Else
        pLogfiles.Add Logfile
    End If
    pLogfileMessage.Add Message
End Sub



'
'
' CLASS METHODS
Public Function ProcessChainOK() As Boolean
' Action: Loops through the process chain backwards and checks that every file needed for that
'         process step either exists or is created in a previous process
'
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim ChainOk As Boolean, InputFileCreated As Boolean
    Dim ErrorCount As Integer, ErrorMessage As String
    
    ErrorCount = 0
    ErrorMessage = ""
    ChainOk = True
    ' Loop through all processes backwards and check that the files specified in the 'FilesBefore'
    '  are either created in som previous step (loop through each process 'FilesAfter') or existing
    '
    For i = pProcesses.Count To 1 Step -1
        For j = 1 To Process(i).FilesBefore.Count
            Debug.Print "   Looking for input file " & Process(i).FilesBefore(j).Filename & " required in process " & CStr(i)
            ' Look for inputfile(j) in earlier processes
            'If Process(i).InputFile(j).FileExists = True Then GoTo Next_j
            InputFileCreated = False
            For k = i - 1 To 1 Step -1
                For m = 1 To Process(k).FilesAfter.Count
                    If Process(i).FilesBefore(j).FullPath = Process(k).FilesAfter(m).FullPath Then
                        Debug.Print "      It is created as output in process " & CStr(k)
                        InputFileCreated = True
                    End If
                Next m
            Next k
            
            ' If file not created
            If InputFileCreated = False Then
                If Process(i).FilesBefore(j).FileExists = True Then
                    ChainOk = True
                    Debug.Print "      Info: " & Process(i).FilesBefore(j).Filename & " exists but isn't created in previous processes"
                Else
                    ChainOk = False
                    Debug.Print "      Error: " & Process(i).FilesBefore(j).Filename & " doesn't exist and isn't created in previous processes"
                    ErrorCount = ErrorCount + 1
                    ErrorMessage = ErrorMessage & vbNewLine & vbNewLine & "Process " & i & ": Error, '" & Process(i).FilesBefore(j).Filename & "' doesn't exist and isn't created in previous processes"
                End If
            End If
Next_j:
        Next j
    Next i
    
    For i = 1 To pProcesses.Count
        For j = 1 To Process(i).FilesRequired.Count
            If Process(i).FilesRequired(j).FileExists = False Then
                ErrorMessage = ErrorMessage & vbNewLine & vbNewLine & "Process " & i & ": Warning, '" & Process(i).FilesRequired(j).FullPath & "' doesn't exist but might be in path"
                'MsgBox Process(i).FilesRequired(j).FullPath & " doesn't exist"
            End If
        Next j
    
    Next i

    Dim answ
    If ErrorMessage <> "" And ErrorCount = 0 Then
        answ = MsgBox(ErrorMessage & vbNewLine & vbNewLine & "Continue?", vbExclamation + vbOKCancel)
        If answ = vbCancel Then
            ChainOk = False
        Else
            ChainOk = True
        End If
    ElseIf ErrorCount > 0 Then
        answ = MsgBox(ErrorMessage & vbNewLine & vbNewLine & "Quitting", vbCritical + vbOKOnly)
        ChainOk = False
    End If
    ProcessChainOK = ChainOk

End Function

Public Function GetShellCommand(Optional WorkingFolder As String, Optional CaseId As String = "Case") As String
    Dim i As Integer
    Dim Shellstring As String
    Shellstring = ""
    For i = 1 To pProcesses.Count
        If pLogfiles(i).Basename <> "" Then
            Shellstring = Shellstring & "(date /T && time /T) >> " & _
                          Chr(34) & pLogfiles(i).GetRelativePath(WorkingFolder) & Chr(34) & " && "
        End If
        Shellstring = Shellstring & "title " & CaseId & " step " & i & " of " & pProcesses.Count & " && " & Process(i).GetShellCommand(WorkingFolder)
        If pLogfiles(i).FileExists Then Shellstring = Shellstring & " >> " & Chr(34) & pLogfiles(i).GetRelativePath(WorkingFolder) & Chr(34)
        If i < pProcesses.Count Then Shellstring = Shellstring & " && "
        Debug.Print Process(i).GetShellCommand(WorkingFolder)
    Next i
    
    GetShellCommand = Shellstring
End Function
'
'
' STATIC METHODS

