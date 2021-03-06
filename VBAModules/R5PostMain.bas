Attribute VB_Name = "R5PostMain"
' R5Post v1.0.0-beta.6
'
'
'
Option Explicit

Const ROW_ZERO = 18   ' The row with the case #

Const CALC_ROW = ROW_ZERO - 6
Const CALC_STRIP_DEMUX_ROW = ROW_ZERO - 5
Const STRIP_DEMUX_ROW = ROW_ZERO - 4
Const POST_ROW = ROW_ZERO - 3
Const CALC_POST_ROW = ROW_ZERO - 2
Const PS2PDF_ROW = ROW_ZERO - 1

Const CASEID_ROW = ROW_ZERO + 1
Const TITLE_ROW = ROW_ZERO + 2
Const TMIN_ROW = ROW_ZERO + 3
Const TMAX_ROW = ROW_ZERO + 4

Public Const CASE_COLUMN_START = 2
Public Const CASE_COLUMN_END = 24

Public Const LOGFILE_ROW = ROW_ZERO + 5
Public Const INPUTFILE_ROW = ROW_ZERO + 6
Public Const OUTPUTFILE_ROW = ROW_ZERO + 7
Public Const RSTFILE_ROW = ROW_ZERO + 8
Public Const DMXFILE_ROW = ROW_ZERO + 9
Public Const STRIPFILE1_ROW = ROW_ZERO + 10
Public Const PARAMFILE_ROW = ROW_ZERO + 11
Public Const STRFILE1_ROW = ROW_ZERO + 12
Public Const CSVFILE_ROW = ROW_ZERO + 13
Public Const PSFILE1_ROW = ROW_ZERO + 14
Public Const PDFFILE1_ROW = ROW_ZERO + 15

Const GLOBAL_STRIPFILE = "B9"
Const GLOBAL_PARAMFILE = "B10"
Const GLOBAL_STRIPFILE_FORCES = "B11"

Public Const R5_PATH = "G3"
Public Const MATLAB_PATH = "G4"
Public Const THISTPLOT_PATH = "G5"
Public Const R2DMX_PATH = "G6"
Public Const GHOSTSCRIPT_PATH = "G7"
Public Const APTPLOT_PATH = "G8"
Public Const STR2CSV_PATH = "G9"

Const CURRENTSHEET = "E1"
Const DEBUG_FLAG = "B7"

Private Enum TypeOfAction
    Calc
    CalcStripDemux
    StripDemux
    Post
    CalcAndPost
    ActionPs2pdf
    AptplotOpenDemux
    AptplotOpenRestart
    AptplotOpenStrip
End Enum

Sub HyperlinkClicked(ByVal cellRow As Long, ByVal cellCol As Long)
' Action: Performs action relating to what link have been clicked
'
    Debug.Print "User clicked " & Chr(64 + cellCol) & cellRow & "   " & CStr(cellCol)
    
    Select Case cellRow
        Case CALC_ROW
            CalculateOrPost cellCol, Calc
        Case CALC_STRIP_DEMUX_ROW
            CalculateOrPost cellCol, CalcStripDemux
        Case STRIP_DEMUX_ROW
            CalculateOrPost cellCol, StripDemux
        Case POST_ROW
            CalculateOrPost cellCol, Post
        Case CALC_POST_ROW
            CalculateOrPost cellCol, CalcAndPost
        Case PS2PDF_ROW
            CalculateOrPost cellCol, ActionPs2pdf
        Case DMXFILE_ROW
            AptplotOpen cellCol, AptplotOpenDemux
        Case RSTFILE_ROW
            AptplotOpen cellCol, AptplotOpenRestart
        Case STRFILE1_ROW
            AptplotOpen cellCol, AptplotOpenStrip
            
    End Select
End Sub

Sub LocateFile()
' Action: Opens a file open dialog box and depending on the type of file, add the path
'
    Dim fileSelected As New R5PostFileObject
    
    If TypeName(Selection) = "Range" Then
        Dim apa As Range
        If Selection.Cells.Count = 1 Or (Selection.Cells.Count = 2 And Selection.Row = ROW_ZERO + 1) Then
            MsgBox "Correct"
        Else
            MsgBox "Select exactly ONE valid cell to add a path", vbCritical, "Add path"
            Exit Sub
        End If
        
    Else
        MsgBox "Select exactly ONE valid cell to add a path", vbCritical, "Add path"
        Exit Sub
    End If
    
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        
        
        .Filters.Clear
        .Filters.Add "All files", "*.*", 1
        
        If Selection.Row = ROW_ZERO + 1 Then
            .Filters.Add "Input files", "*.i", 1
            .InitialFileName = ThisWorkbook.Path & "\"
        Else
            .Filters.Add "All files", "*.*", 1
        End If
        
        If .Show = True Then
            fileSelected.Create .SelectedItems(1)
            MsgBox fileSelected.GetRelativePath(ThisWorkbook.Path & "\")
        End If
        
    End With
    
    

End Sub

Private Sub AptplotOpen(Cellcolumn As Long, Action As TypeOfAction)
' Action: Opens aptplot and automatically reads selected file
'
    Dim answ
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet

    Dim pwd As String
    pwd = ThisWorkbook.Path

    Dim ProcAptplotOpen As New ProcessAptplotOpen
    Dim AptplotPath As New R5PostFileObject
    AptplotPath.CreateByParts sht.Range(APTPLOT_PATH)
    
    Dim FileToOpen As New R5PostFileObject
    
    Select Case Action
        Case AptplotOpenDemux
            FileToOpen.CreateByParts pwd, sht.Cells(DMXFILE_ROW, Cellcolumn)
            ProcAptplotOpen.Create AptplotPath, FileToOpen, "demux"
            
        Case AptplotOpenRestart
            FileToOpen.CreateByParts pwd, sht.Cells(RSTFILE_ROW, Cellcolumn)
            ProcAptplotOpen.Create AptplotPath, FileToOpen, "restart"
        
        Case AptplotOpenStrip
            FileToOpen.CreateByParts pwd, sht.Cells(STRFILE1_ROW, Cellcolumn)
            ProcAptplotOpen.Create AptplotPath, FileToOpen, "strip"
        Case Else
            Exit Sub
    End Select
    
    
    Dim Calculate As New MainProcessChain
    Calculate.Add ProcAptplotOpen
    
    If Calculate.ProcessChainOK = True Then
        Dim ShellCommand As String
        ShellCommand = Calculate.GetShellCommand(FileToOpen.FolderPath & "\", FileToOpen.Basename)
        
        Dim questionString As String
        questionString = "Open file in Aptplot?" & vbNewLine & vbNewLine & SplitShellCommand(ShellCommand)
        answ = MsgBox(questionString, vbQuestion + vbYesNoCancel, "Open file with Aptplot?")
        If answ <> vbYes Then Exit Sub
        
        ' Change drive and current directory on that drive
        ChDir FileToOpen.FolderPath
        ChDrive FileToOpen.DriveName
        Dim retval
        retval = Shell("cmd /S /C" & " dir && timeout 1 && " & ShellCommand, 1)
    End If
End Sub




Private Sub CalculateOrPost(Cellcolumn As Long, Action As TypeOfAction)
' Action: Executes a process chain to perform a simulation or post process the result
'
    ' Misc
    Dim answ

    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName
    
    ' Define paths
    Dim R5Path As New R5PostFileObject
    Dim R5SteamTable As New R5PostFileObject
    Dim R5SteamTableOld As New R5PostFileObject
    Dim MatlabPath As New R5PostFileObject
    Dim THistPlotPath As New R5PostFileObject
    Dim R2DMXPath As New R5PostFileObject
    Dim GhostScriptPath As New R5PostFileObject
    Dim Str2CsvPath As New R5PostFileObject
    
    R5Path.Create sht.Range(R5_PATH)
    R5SteamTable.Create R5Path.FolderPath & "\tpfh2onew"
    R5SteamTableOld.Create R5Path.FolderPath & "\tpfh2o"
    MatlabPath.Create sht.Range(MATLAB_PATH)
    THistPlotPath.Create sht.Range(THISTPLOT_PATH)
    R2DMXPath.Create sht.Range(R2DMX_PATH)
    GhostScriptPath.Create sht.Range(GHOSTSCRIPT_PATH)
    Str2CsvPath.Create sht.Range(STR2CSV_PATH)
    
    ' ---------------DEFINE FILES-----------------------
    '
    ' RELAP5 restart file to start simulation from - existing and used in ProcCopySSRestartFile
    Dim SteadyStateRestartfile As New R5PostFileObject
    SteadyStateRestartfile.CreateByParts ThisWorkbook.Path, sht.Cells(LOGFILE_ROW, Cellcolumn)
    
    ' RELAP5 input file (.i) - existing
    Dim Inputfile As New R5PostFileObject
    Inputfile.CreateByParts ThisWorkbook.Path, sht.Cells(INPUTFILE_ROW, Cellcolumn)
    
    ' RELAP5 output file (.o) - created in ProcR5Calc
    Dim OutputFile As New R5PostFileObject
    OutputFile.CreateByParts ThisWorkbook.Path, sht.Cells(OUTPUTFILE_ROW, Cellcolumn)
    
    ' RELAP5 restart file (.rst) - created in ProcR5Calc
    Dim RestartFile As New R5PostFileObject
    RestartFile.CreateByParts ThisWorkbook.Path, sht.Cells(RSTFILE_ROW, Cellcolumn)
    
    ' Demuxed restart file (.dmx) - created in ProcR2DMX
    Dim Demuxfile As New R5PostFileObject
    Demuxfile.CreateByParts ThisWorkbook.Path, sht.Cells(DMXFILE_ROW, Cellcolumn)
    
    ' Strip request file, global or local - existing and used by ProcR5Strip
    Dim StripRequestfile1 As New R5PostFileObject
    If sht.Cells(STRIPFILE1_ROW, Cellcolumn) = "" Then
        StripRequestfile1.CreateByParts ThisWorkbook.Path, sht.Range(GLOBAL_STRIPFILE)
    Else
        StripRequestfile1.CreateByParts ThisWorkbook.Path, sht.Cells(STRIPFILE1_ROW, Cellcolumn)
    End If
    
    ' Strip file (.str) - created in ProcR5Strip
    Dim Stripfile1 As New R5PostFileObject
    Stripfile1.CreateByParts ThisWorkbook.Path, sht.Cells(STRFILE1_ROW, Cellcolumn)
    
    ' Parameter file, global or local - existing and used by ProcTHistPlot
    Dim Paramfile As New R5PostFileObject
    If sht.Cells(PARAMFILE_ROW, Cellcolumn) = "" Then
        Paramfile.CreateByParts ThisWorkbook.Path, sht.Range(GLOBAL_PARAMFILE)
    Else
        Paramfile.CreateByParts ThisWorkbook.Path, sht.Cells(PARAMFILE_ROW, Cellcolumn)
    End If
    
    ' PostScript file (.ps) - created by ProcThistPlot and used by ProcPs2Pdf
    Dim PSfile1 As New R5PostFileObject
    PSfile1.CreateByParts ThisWorkbook.Path, sht.Cells(PSFILE1_ROW, Cellcolumn)
    
    ' PDF file (.pdf) - created by ProcPs2Pdf
    Dim PDFfile1 As New R5PostFileObject
    PDFfile1.CreateByParts ThisWorkbook.Path, sht.Cells(PDFFILE1_ROW, Cellcolumn)
    
    ' CSV file (.csv) - created by ProcStr2Csv
    Dim Csvfile As New R5PostFileObject
    Csvfile.CreateByParts ThisWorkbook.Path, sht.Cells(CSVFILE_ROW, Cellcolumn)
    
    
    ' Define parameters used in ProcTHistPlot
    Dim PlotTitle As String
    Dim PlotTimeMin As Double
    Dim PlotTimeMax As Double
    
    PlotTitle = sht.Cells(TITLE_ROW, Cellcolumn)
    PlotTimeMin = sht.Cells(TMIN_ROW, Cellcolumn)
    PlotTimeMax = sht.Cells(TMIN_ROW, Cellcolumn + 1)
    
    ' ---------------CREATE PROCESSES-------------------
    '
    ' Create "Copy steady state restart file" process
    Dim ProcCopySSRestartFile As New ProcessManual
    Dim UseBaseRst As Boolean
    Dim s As New ResourceSprintf
    If sht.Cells(LOGFILE_ROW, Cellcolumn) = "" Then
        UseBaseRst = False
        ProcCopySSRestartFile.Create ""
    Else
        UseBaseRst = True
        Dim fBefore As New CollectionFileList
        Dim fAfter As New CollectionFileList
        fBefore.Add SteadyStateRestartfile
        fAfter.AddMany SteadyStateRestartfile, RestartFile
    
        ProcCopySSRestartFile.Create s.sprintf("copy %s %s /Y", SteadyStateRestartfile.GetRelativePath(Inputfile.FolderPath & "\"), RestartFile.Filename), FilesBeforeProcess:=fBefore, FilesAfterProcess:=fAfter
    End If
    
    ' Create Relap5 calculation process
    Dim ProcR5Calc As New ProcessR5Calc
    ProcR5Calc.Create R5Path, R5SteamTable, R5SteamTableOld, Inputfile, OutputFile, RestartFile
    
    ' Create Relap5 strip process
    Dim ProcR5Strip As New ProcessR5Strip
    ProcR5Strip.Create R5Path, StripRequestfile1, RestartFile, Stripfile1
    
    ' Create rst to demux process
    Dim ProcR2DMX As New ProcessR2DMX
    ProcR2DMX.Create R2DMXPath, RestartFile, Demuxfile
    
    ' Create process with the matlab script 'THistPlot'
    Dim ProcTHistPlot As New ProcessTHistPlot
    ProcTHistPlot.Create MatlabPath, ScriptPath:=THistPlotPath, StripFile:=Stripfile1, Paramfile:=Paramfile, _
                         Psfile:=PSfile1, Title:=PlotTitle, tMin:=PlotTimeMin, tMax:=PlotTimeMax
    
    ' Create process that converts a .ps-file to pdf using ghostscript
    Dim ProcPs2Pdf As New ProcessPs2Pdf
    ProcPs2Pdf.Create GhostScriptPath, PSfile1, PDFfile1
    
    ' Create Str2Csv process
    Dim ProcStr2Csv As New ProcessStr2Csv
    ProcStr2Csv.Create Str2CsvPath, Stripfile1, Csvfile
    
    ' Create process chain
    Dim Calculate As New MainProcessChain
    
    Debug.Print Inputfile.FullPath
    
    Dim questionString As String, questionTitle As String
    
    If Action = Calc Then
        If UseBaseRst = True Then Calculate.Add ProcCopySSRestartFile
        Calculate.Add ProcR5Calc
        questionTitle = "Perform calculation?"
        questionString = "Perform RELAP5-calculation on '" & Inputfile.FullPath & "'?"
        
    ElseIf Action = CalcAndPost Then
        If UseBaseRst = True Then Calculate.Add ProcCopySSRestartFile
        Calculate.Add ProcR5Calc
        Calculate.Add ProcR5Strip
        Calculate.Add ProcStr2Csv
        Calculate.Add ProcTHistPlot
        Calculate.Add ProcPs2Pdf
        Calculate.Add ProcR2DMX
        
        questionTitle = "Perform calculation+postprocessing?"
        questionString = "Perform RELAP5-calculation and postprocessing on '" & Inputfile.FullPath & "'?"
        
    ElseIf Action = Post Then
        Calculate.Add ProcR5Strip
        Calculate.Add ProcStr2Csv
        Calculate.Add ProcTHistPlot
        Calculate.Add ProcPs2Pdf
        Calculate.Add ProcR2DMX
        
        questionTitle = "Perform postprocessing?"
        questionString = "Perform RELAP5-postprocessing on '" & RestartFile.FullPath & "'?"
        
    ElseIf Action = CalcStripDemux Then
        If UseBaseRst = True Then Calculate.Add ProcCopySSRestartFile
        Calculate.Add ProcR5Calc
        Calculate.Add ProcR5Strip
        Calculate.Add ProcStr2Csv
        Calculate.Add ProcR2DMX
        questionTitle = "Perform calculation+strip+demux?"
        questionString = "Perform RELAP5-calculation and strip+demux on '" & Inputfile.FullPath & "'?"
        
    ElseIf Action = StripDemux Then
        Calculate.Add ProcR5Strip
        Calculate.Add ProcStr2Csv
        Calculate.Add ProcR2DMX
        questionTitle = "Perform strip+demux?"
        questionString = "Perform Strip and demux on '" & RestartFile.FullPath & "'?"
        
    ElseIf Action = ActionPs2pdf Then
        Calculate.Add ProcPs2Pdf
        questionTitle = "Convert to pdf?"
        questionString = "Convert following ps-files to pdf '" & PSfile1.FullPath & "'?"
    End If
    
    If Calculate.ProcessChainOK = True Then
        Dim ShellCommand As String
        ShellCommand = Calculate.GetShellCommand(Inputfile.FolderPath & "\", Inputfile.Basename)
    
        questionString = questionString & vbNewLine & vbNewLine & SplitShellCommand(ShellCommand)
        answ = MsgBox(questionString, vbQuestion + vbYesNoCancel, questionTitle)
        If answ <> vbYes Then Exit Sub
        
        ' Change drive and current directory on that drive
        ChDir Inputfile.FolderPath
        ChDrive Inputfile.DriveName
        Debug.Print "Current directory is = """ & CurDir & """"
        
        Dim retval
        Dim stayOpen As String
        If Range(DEBUG_FLAG).Value = 1 Then stayOpen = "/K" Else stayOpen = "/C"
        retval = Shell("cmd /S " & stayOpen & " dir && timeout 1 && " & Calculate.GetShellCommand(Inputfile.FolderPath & "\", Inputfile.Basename), 1)
    Else
        MsgBox "Error"
    End If
        
End Sub

Private Function SplitShellCommand(ByVal ShellCommand As String, Optional ByVal Separator As String = "&&") As String
' Action: Returns a shellcommand that is split between
'
    Dim Arr() As String
    Dim outputString As String
    outputString = ""
    Dim i As Integer
    Arr = Split(ShellCommand, Separator)
    
    For i = LBound(Arr) To UBound(Arr)
        outputString = outputString & Arr(i) & IIf(i = UBound(Arr), "", Separator & vbNewLine)
    Next i
    
    SplitShellCommand = outputString
    
End Function


Sub RefreshFileDates()
' Action: Refreshes the file dates
'
    Dim i As Integer, j As Integer
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet

    ToggleAutoCalc False  ' Turns off automatic update

    sht.Range(CURRENTSHEET).Value = "[" & ThisWorkbook.FullName & "]'" & ActiveSheet.Name & "'"

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName
    
    ' Define files
    Dim fileCurr As New R5PostFileObject
    Dim OutputCellCurr As Range
    
    ' Check calculation process files (*.i, *.rst, etc)
    Dim ColumnCurr As Integer
    For i = CASE_COLUMN_START To CASE_COLUMN_END Step 2
        For j = LOGFILE_ROW To PDFFILE1_ROW
            fileCurr.CreateByParts ThisWorkbook.Path, sht.Cells(j, i)
            Set OutputCellCurr = sht.Cells(j, i + 1)
            If fileCurr.FileExists = True Then
                OutputCellCurr.Value = fileCurr.DateLastModified
            ElseIf sht.Cells(j, i) = "" Then
                OutputCellCurr.Value = ""
            Else
                OutputCellCurr.Value = "(missing)"
            End If
        Next j
    Next i
    
    ' Check executables
    Dim ExecPaths As Range
    Set ExecPaths = sht.Range(sht.Range(R5_PATH), sht.Range(STR2CSV_PATH))
    For i = 1 To ExecPaths.Rows.Count
        fileCurr.CreateByParts ExecPaths(i)
        If fileCurr.FileExists = True Then
            sht.Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = "OK"
        ElseIf ExecPaths(i) = "" Then
            sht.Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = ""
        Else
            sht.Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = "(missing)"
        End If
    Next i
    

    ' Check global files
    Dim GlobalFilePaths As Range
    Set GlobalFilePaths = sht.Range(sht.Range(GLOBAL_STRIPFILE), sht.Range(GLOBAL_STRIPFILE_FORCES))
    For i = 1 To GlobalFilePaths.Rows.Count
        fileCurr.CreateByParts ThisWorkbook.Path, GlobalFilePaths(i)
        If fileCurr.FileExists = True Then
            sht.Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = fileCurr.DateLastModified
        ElseIf GlobalFilePaths(i) = "" Then
            sht.Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = ""
        Else
            sht.Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = "(missing)"
        End If
    Next i
    
    ToggleAutoCalc True  ' Turns on automatic update

End Sub

Sub SetPaths()
' Action: Sets the path
'
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet
    
    ' Check executables
    Dim fileCurr As New R5PostFileObject
    Dim ExecPaths As Range, i As Integer, j As Integer
    Set ExecPaths = sht.Range(sht.Cells(sht.Range(R5_PATH).Row, sht.Range(R5_PATH).Column - 1), sht.Range(STR2CSV_PATH))
    For i = 1 To ExecPaths.Rows.Count
        fileCurr.CreateByParts ExecPaths(i, 2)
        If fileCurr.FileExists = False Or ExecPaths(i, 2) = "" Then
            
            
            ' Open the file dialog
            With Application.FileDialog(msoFileDialogOpen)
                .AllowMultiSelect = False
                
                .Title = ExecPaths(i, 1) & " Path"
                On Error Resume Next
                .InitialFileName = Split(ExecPaths(i, 1), "=")(0)
                .Show
                
                ' Display paths of each file selected
                For j = 1 To .SelectedItems.Count
                    ExecPaths(i, 2) = .SelectedItems(j)
                Next j
            
            End With
            
            
        End If
    Next i
    
    
 
    Dim lngCount As Long
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
 
        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            MsgBox .SelectedItems(lngCount)
        Next lngCount
 
    End With
 

End Sub


Sub fixLinks()
    ' Fixar l�nkar. Markera de celler som ska fixas och k�r makrot
    Dim i As Integer, j As Integer
    
    Dim actionLinks As Range
    Set actionLinks = Range("B12:BG17")
    
    
    
    For i = 1 To actionLinks.Columns.Count Step 2
        For j = 1 To actionLinks.Rows.Count
            'actionLinks(j, i).Select
            actionLinks(j, i).Hyperlinks.Delete
            'MsgBox i
            
            actionLinks(j, i).Hyperlinks.Add Anchor:=actionLinks(j, i), Address:="", SubAddress:="'" & ActiveSheet.Name & "'!" & actionLinks(j, i).Address
            'ActiveSheet.Hyperlinks.Add Anchor:=actionLinks(j, i), Address:="", SubAddress:="'" & ActiveSheet.Name & "'!" & actionLinks(j, i).Address
        Next j
    Next i

    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources, Type:=xlExcelLinks
End Sub


Sub PurgeFiles()
' Action: Removes every file except the tracked files
'
    Dim i As Integer, j As Integer, k As Integer
    Dim deleteFile As Boolean
    Dim tmpStr As String
    Dim answ

    answ = MsgBox("This will delete all non-tracked files (files listed below). Continue?", vbExclamation + vbYesNoCancel, "Delete files?")
    If answ <> vbYes Then Exit Sub

    Dim fso As New FileSystemObject
    Dim folderCurr, filesInCurrFolder, fileToCheckForDeletion, fileToDelete
    Dim filesToDelete As New Collection

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName

    Dim InputFileCurr As New R5PostFileObject
    Dim TrackedFileCurr As New R5PostFileObject
    Dim TrackedFileCurrPath As String

    For i = CASE_COLUMN_START To CASE_COLUMN_END Step 2
        If Cells(INPUTFILE_ROW, i) = "" Then GoTo next_case
        
        InputFileCurr.CreateByParts ThisWorkbook.Path, Cells(INPUTFILE_ROW, i)
        
        Set folderCurr = fso.GetFolder(InputFileCurr.FolderPath)
        Set filesInCurrFolder = folderCurr.Files
        Debug.Print InputFileCurr.FolderPath
        For Each fileToCheckForDeletion In filesInCurrFolder
            deleteFile = True
            TrackedFileCurrPath = fileToCheckForDeletion.Path
            For j = CASE_COLUMN_START To CASE_COLUMN_END Step 2
                For k = LOGFILE_ROW To PDFFILE1_ROW
                    TrackedFileCurr.CreateByParts ThisWorkbook.Path, Cells(k, j)
                    If TrackedFileCurrPath = TrackedFileCurr.FullPath Then deleteFile = False
                Next k
            Next j
            
            Debug.Print "   " & fileToCheckForDeletion.Name & " delete = " & CStr(deleteFile)
            If deleteFile = True Then filesToDelete.Add fileToCheckForDeletion
        Next
        
        ' Message
        If filesToDelete.Count > 0 Then
            tmpStr = ""
            For Each fileToDelete In filesToDelete
                tmpStr = tmpStr & vbNewLine & "   " & fileToDelete.Name
            Next
            
            answ = MsgBox("Delete following files" & vbNewLine & folderCurr.Path & tmpStr, vbQuestion + vbYesNoCancel, "Delete files")
            If answ = vbYes Then
                For Each fileToDelete In filesToDelete
                    fileToDelete.Delete
                Next
            ElseIf answ = vbCancel Then
                Exit Sub
            End If
            Set filesToDelete = New Collection
        End If
next_case:
    Next i
    
    
    
    
End Sub


Private Sub ToggleAutoCalc(Optional TurnOn As Boolean = True)
    If TurnOn = True Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    Else
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
End Sub




Sub ShowFileList()

    ' Loopa igenom alla case. Lista alla filer i katalogen och loopa igenom dessa en efter en
    ' Ta bort den aktuella filen om den inte �r en trackad fil i n�got case (loopa igenom en g�ng till)
    '
    ' Loop alla case
    '   Loopa alla filer i aktuellt case katalog
    '       Loopa igenom alla case
    '           Ta bort fil om den inte finns
    '



    Dim folderSpec As String
    folderSpec = "C:\Users\Danne\Downloads\"

    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderSpec)
    Set fc = f.Files
    For Each f1 In fc
        
        s = s & f1.Name
        s = s & vbCrLf
    Next
    MsgBox s
    Debug.Print s
End Sub
