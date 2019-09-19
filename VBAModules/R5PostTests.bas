Attribute VB_Name = "R5PostTests"
Option Explicit

Private Sub TestFileObject()
    Debug.Print "<TESTING FILE OBJECT>"

    Dim Inputfile As New R5PostFileObject
    
    Inputfile.CreateByParts ThisWorkbook.Path, "Case1\Case1.i"
    
    Debug.Print "FullPath=" & Inputfile.FullPath
    Debug.Print "FolderPath=" & Inputfile.FolderPath
    Debug.Assert Inputfile.Basename = "Case1"
    Debug.Assert Inputfile.Extension = "i"
    Debug.Assert Inputfile.FolderPath = ThisWorkbook.Path & "\Case1"
    Debug.Assert Inputfile.GetRelativePath(ThisWorkbook.Path & "\") = "Case1\Case1.i"
    
    Debug.Print "</TESTING FILE OBJECT>"
End Sub


Private Sub TestProcessR5Calc()
    Debug.Print "<TESTING PROCESS>"

    Dim Proc As New ProcessR5Calc
    Dim Relap5Path As New R5PostFileObject
    Dim Relap5SteamTableOld As New R5PostFileObject
    Dim Relap5SteamTableNew As New R5PostFileObject
    Dim Inputfile As New R5PostFileObject
    Dim OutputFile As New R5PostFileObject
    Dim RestartFile As New R5PostFileObject
    
    Relap5Path.Create "C:\SomePath\Relap5.exe"
    Relap5SteamTableOld.Create "C:\SomePath\tpfh2o"
    Relap5SteamTableNew.Create "C:\SomePath\tpfh2onew"
    Inputfile.Create "Case1.i"
    OutputFile.Create "Case1.o"
    RestartFile.Create "Case1.rst"
    
    Proc.Create Relap5Path, Relap5SteamTableNew, Relap5SteamTableOld, Inputfile, OutputFile, RestartFile
    Debug.Print Proc.GetShellCommand(ThisWorkbook.Path & "\")
    Debug.Print "</TESTING PROCESS>"
End Sub


Private Sub TestProcessTHistPlot()
    Debug.Print "<TESTING PROCESS>"

    Dim pwd As String
    pwd = ThisWorkbook.Path

    Dim Proc As New ProcessTHistPlot
    Dim MatlabPath As New R5PostFileObject
    Dim ScriptPath As New R5PostFileObject
    
    Dim Strfile As New R5PostFileObject
    Dim Paramfile As New R5PostFileObject
    Dim Psfile As New R5PostFileObject
    
    MatlabPath.Create "C:\SomePath\Matlab.exe"
    ScriptPath.Create "C:\MyScripts\THistPlot\THistPlot.m"
    
    Strfile.CreateByParts pwd, "Case1\Case1.str"
    Paramfile.CreateByParts "Param.txt"
    Psfile.CreateByParts "Case1\Case1.ps"
    
    Proc.Create MatlabPath, ScriptPath, Strfile, Paramfile, Psfile, "Case1 - Pump start", 10, 15
    Debug.Print Proc.GetShellCommand(Strfile.FolderPath)
    Debug.Print "</TESTING PROCESS>"
End Sub


Private Sub TestAptplotOpen()
    Debug.Print "<TESTING PROCESS>"

    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet

    Dim pwd As String
    pwd = ThisWorkbook.Path

    Dim Proc As New ProcessAptplotOpen
    Dim AptplotPath As New R5PostFileObject
    AptplotPath.CreateByParts sht.Range(APTPLOT_PATH)
    
    Dim Rstfile As New R5PostFileObject
    Rstfile.CreateByParts pwd, "Case1\Case1.rst"
    
    Proc.Create AptplotPath, Rstfile
    Debug.Print Proc.GetShellCommand
    
    Dim Calculate As New MainProcessChain
    Calculate.Add Proc
    
    If Calculate.ProcessChainOK = True Then
        Dim ShellCommand As String
        Dim retval
        ShellCommand = Calculate.GetShellCommand(Rstfile.FolderPath & "\", Rstfile.Basename)
    
        ChDir Rstfile.FolderPath

        retval = Shell("cmd /S /K" & " dir && timeout 1 && " & ShellCommand, 1)
    End If
    
    
    
End Sub


Private Sub TestProcessChain()
    Debug.Print "<TESTING PROCESSCHAIN>"
    
    Dim pwd As String
    pwd = ThisWorkbook.Path
    
    ' Process 1
    Dim ExeFile1 As New R5PostFileObject
    Dim Inputfile As New R5PostFileObject
    Dim OutputFile As New R5PostFileObject
    
    ExeFile1.CreateByParts "C:\Windows\System32\cmd.exe"
    Inputfile.CreateByParts ThisWorkbook.FullName
    OutputFile.CreateByParts pwd, "Output1.dat"
    
    ' Process 2 - uses Outputfile and creates another output file, OutputFile2
    Dim ExeFile2 As New R5PostFileObject
    Dim OutputFile2 As New R5PostFileObject
    
    ExeFile2.CreateByParts "C:\Windows\System32\cmd2.exe"  ' Produces a warning
    OutputFile2.CreateByParts pwd, "Output2.dat"
    
    ' Create filesets for both processes
    Dim FilesRequired1 As New CollectionFileList
    Dim FilesBefore1 As New CollectionFileList
    Dim FilesAfter1 As New CollectionFileList
    FilesRequired1.Add ExeFile1
    FilesBefore1.Add Inputfile
    FilesAfter1.AddMany Inputfile, OutputFile
    
    Dim FilesRequired2 As New CollectionFileList
    Dim FilesAfter2 As New CollectionFileList
    FilesRequired2.Add ExeFile2
    FilesAfter2.AddList FilesAfter1
    FilesAfter2.Add OutputFile2
    
    ' Create processes and process chain
    Dim Proc1 As New ProcessX
    Dim Proc2 As New ProcessX
    Dim ProcessChain As New MainProcessChain
    
    Proc1.Create FilesRequired1, FilesBefore1, FilesAfter1
    Proc2.Create FilesRequired2, FilesAfter1, FilesAfter2   ' Change FilesAfter1 to FilesAfter2 to create an error
    ProcessChain.Add Proc1
    ProcessChain.Add Proc2
    
    
    ' Check process chain
    Debug.Print "CHECKING PROCESS CHAIN"
    ProcessChain.ProcessChainOK
    
    ' Get
    Debug.Print "PROCESS CHAIN COMMAND LINE"
    ProcessChain.GetShellCommand pwd & "\"
    
    
    Debug.Print "</TESTING PROCESSCHAIN>"
End Sub






