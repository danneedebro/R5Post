Attribute VB_Name = "R5PostMain"
Option Explicit

Const ROW_ZERO = 18   ' The row with the case #

Const CALC_ROW = ROW_ZERO - 6
Const CALC_STRIP_DEMUX_ROW = ROW_ZERO - 5
Const STRIP_DEMUX_ROW = ROW_ZERO - 4
Const POST_ROW = ROW_ZERO - 3
Const CALC_POST_ROW = ROW_ZERO - 2
Const PS2PDF_ROW = ROW_ZERO - 1

Const TITLE_ROW = ROW_ZERO + 2
Const TMIN_ROW = ROW_ZERO + 3
Const TMAX_ROW = ROW_ZERO + 4

Const CASE_COLUMN_START = 2
Const CASE_COLUMN_END = 24

Const LOGFILE_ROW = ROW_ZERO + 5
Const INPUTFILE_ROW = ROW_ZERO + 6
Const OUTPUTFILE_ROW = ROW_ZERO + 7
Const RSTFILE_ROW = ROW_ZERO + 8
Const DMXFILE_ROW = ROW_ZERO + 9
Const STRIPFILE1_ROW = ROW_ZERO + 10
Const PARAMFILE_ROW = ROW_ZERO + 11
Const STRFILE1_ROW = ROW_ZERO + 12
Const PSFILE1_ROW = ROW_ZERO + 13
Const PDFFILE1_ROW = ROW_ZERO + 14

Const GLOBAL_STRIPFILE = "B9"
Const GLOBAL_PARAMFILE = "B10"
Const GLOBAL_STRIPFILE_FORCES = "B11"

Const R5_PATH = "G3"
Const MATLAB_PATH = "G4"
Const PLOTTASTRIPFIL_PATH = "G5"
Const R2DMX_PATH = "G6"
Const GHOSTSCRIPT_PATH = "G7"

Private Enum TypeOfAction
    Calc
    CalcStripDemux
    StripDemux
    Post
    CalcAndPost
    ActionPs2pdf
End Enum

Sub HyperlinkClicked(cellRange As Range)
' Action: Performs action relating to what link have been clicked
'
    MsgBox "User clicked " & cellRange.Address & "   " & CStr(cellRange.Column)
    
    Select Case cellRange.Row
        Case CALC_ROW
            CalculateOrPost cellRange.Column, Calc
        Case CALC_STRIP_DEMUX_ROW
            CalculateOrPost cellRange.Column, CalcStripDemux
        Case STRIP_DEMUX_ROW
            CalculateOrPost cellRange.Column, StripDemux
        Case POST_ROW
            CalculateOrPost cellRange.Column, Post
        Case CALC_POST_ROW
            CalculateOrPost cellRange.Column, CalcAndPost
        Case PS2PDF_ROW
            CalculateOrPost cellRange.Column, ActionPs2pdf
    End Select
        
    
    'If cellRange.Row = CALC_ROW Then
    '    CalculateOrPost cellRange.Column, Calc
    'ElseIf cellRange.Row = POST_ROW Then
    '    CalculateOrPost cellRange.Column, Post
    'ElseIf cellRange.Row = CALC_AND_POST_ROW Then
    '    CalculateOrPost cellRange.Column, CalcAndPost
    'End If
    
    
    

End Sub


Private Sub CalculateOrPost(Cellcolumn As Integer, Action As TypeOfAction)
' Action:
'
    ' Misc
    Dim answ

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName
    
    ' Define paths
    Dim R5Path As New R5PostFileObject
    Dim R5SteamTable As New R5PostFileObject
    Dim R5SteamTableOld As New R5PostFileObject
    Dim MatlabPath As New R5PostFileObject
    Dim PlottaStripfilPath As New R5PostFileObject
    Dim R2DMXPath As New R5PostFileObject
    Dim GhostScriptPath As New R5PostFileObject
    
    R5Path.Create Range(R5_PATH)
    R5SteamTable.Create R5Path.FolderPath & "\tpfh2onew"
    R5SteamTableOld.Create R5Path.FolderPath & "\tpfh2o"
    MatlabPath.Create Range(MATLAB_PATH)
    PlottaStripfilPath.Create Range(PLOTTASTRIPFIL_PATH)
    R2DMXPath.Create Range(R2DMX_PATH)
    GhostScriptPath.Create Range(GHOSTSCRIPT_PATH)
    
    ' Define files
    Dim InputFile As New R5PostFileObject
    Dim Outputfile As New R5PostFileObject
    Dim Restartfile As New R5PostFileObject
    Dim Demuxfile As New R5PostFileObject
    Dim StripRequestfile1 As New R5PostFileObject
    Dim Paramfile As New R5PostFileObject
    Dim Stripfile1 As New R5PostFileObject
    Dim PSfile1 As New R5PostFileObject
    Dim PDFfile1 As New R5PostFileObject
    Dim Logfile As New R5PostFileObject
    
    InputFile.CreateByParts ThisWorkbook.Path, Cells(INPUTFILE_ROW, Cellcolumn)
    Outputfile.CreateByParts ThisWorkbook.Path, Cells(OUTPUTFILE_ROW, Cellcolumn)
    Restartfile.CreateByParts ThisWorkbook.Path, Cells(RSTFILE_ROW, Cellcolumn)
    Demuxfile.CreateByParts ThisWorkbook.Path, Cells(DMXFILE_ROW, Cellcolumn)
    If Cells(STRIPFILE1_ROW, Cellcolumn) = "" Then
        StripRequestfile1.CreateByParts ThisWorkbook.Path, Range(GLOBAL_STRIPFILE)
    Else
        StripRequestfile1.CreateByParts ThisWorkbook.Path, Cells(STRIPFILE1_ROW, Cellcolumn)
    End If
    If Cells(PARAMFILE_ROW, Cellcolumn) = "" Then
        Paramfile.CreateByParts ThisWorkbook.Path, Range(GLOBAL_PARAMFILE)
    Else
        Paramfile.CreateByParts ThisWorkbook.Path, Cells(PARAMFILE_ROW, Cellcolumn)
    End If
    
    Stripfile1.CreateByParts ThisWorkbook.Path, Cells(STRFILE1_ROW, Cellcolumn)
    PSfile1.CreateByParts ThisWorkbook.Path, Cells(PSFILE1_ROW, Cellcolumn)
    PDFfile1.CreateByParts ThisWorkbook.Path, Cells(PDFFILE1_ROW, Cellcolumn)
    Logfile.CreateByParts ThisWorkbook.Path, Cells(LOGFILE_ROW, Cellcolumn)
    
    ' Define parameters
    Dim PlotTitle As String
    Dim PlotTimeMin As Double
    Dim PlotTimeMax As Double
    
    PlotTitle = Cells(TITLE_ROW, Cellcolumn)
    PlotTimeMin = Cells(TMIN_ROW, Cellcolumn)
    PlotTimeMax = Cells(TMAX_ROW, Cellcolumn)
    
    ' Create Relap5 calculation process
    Dim R5Calc As New ProcessR5Calc
    R5Calc.Create R5Path, R5SteamTable, R5SteamTableOld, InputFile, Outputfile, Restartfile
    
    ' Create Relap5 strip process
    Dim R5Strip As New ProcessR5Strip
    R5Strip.Create R5Path, StripRequestfile1, Restartfile, Stripfile1
    
    ' Create rst to demux process
    Dim R2DMX As New ProcessR2DMX
    R2DMX.Create R2DMXPath, Restartfile, Demuxfile
    
    ' Create process with the matlab script 'plottaStripfil'
    Dim MatlabPlottaStripfil As New ProcessPlottaStripfil
    MatlabPlottaStripfil.Create MatlabPath, PlottaStripfilPath, Stripfile1, Paramfile, PSfile1, PlotTitle, PlotTimeMin, PlotTimeMax
    
    ' Create process that converts a .ps-file to pdf using ghostscript
    Dim GhostScript As New ProcessPs2Pdf
    GhostScript.Create GhostScriptPath, PSfile1, PDFfile1
    
    ' Create process chain
    Dim Calculate As New MainProcessChain
    
    If Action = Calc Then
        Calculate.Add R5Calc
        answ = MsgBox("Perform RELAP5-calculation on '" & InputFile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform calculation?")
        If answ <> vbYes Then Exit Sub
    ElseIf Action = CalcAndPost Then
        Calculate.Add R5Calc
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        Calculate.Add MatlabPlottaStripfil
        answ = MsgBox("Perform RELAP5-calculation and postprocessing on '" & InputFile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform calculation+postprocessing?")
        If answ <> vbYes Then Exit Sub
    ElseIf Action = Post Then
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        Calculate.Add MatlabPlottaStripfil
        answ = MsgBox("Perform RELAP5-postprocessing on '" & Restartfile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform postprocessing?")
        If answ <> vbYes Then Exit Sub
    ElseIf Action = CalcStripDemux Then
        Calculate.Add R5Calc
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        answ = MsgBox("Perform RELAP5-calculation and strip+demux on '" & InputFile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform calculation+strip+demux?")
        If answ <> vbYes Then Exit Sub
    ElseIf Action = StripDemux Then
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        answ = MsgBox("Perform Strip and demux on '" & Restartfile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform strip+demux?")
        If answ <> vbYes Then Exit Sub
    ElseIf Action = ActionPs2pdf Then
        Calculate.Add GhostScript
        answ = MsgBox("Convert following ps-files to pdf '" & PSfile1.FullPath & "'?", vbQuestion + vbYesNoCancel, "Convert to pdf?")
        If answ <> vbYes Then Exit Sub
    End If
    
    If Calculate.ProcessChainOK = True Then
        MsgBox "Hej " & Calculate.GetShellCommand(InputFile.FolderPath & "\")
        Dim retval
        ChDir InputFile.FolderPath
        retval = Shell("cmd /S /K dir && timeout 1 && " & Calculate.GetShellCommand(InputFile.FolderPath & "\"), 1)
    Else
        MsgBox "Error"
    End If
        
End Sub


Sub RefreshFileDates()
' Action: Refreshes the file dates
'
    Dim i As Integer, j As Integer

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName
    
    ' Define files
    Dim FileCurr As New R5PostFileObject
    Dim OutputCellCurr As Range
    
    ' Check calculation process files (*.i, *.rst, etc)
    Dim ColumnCurr As Integer
    For i = CASE_COLUMN_START To CASE_COLUMN_END Step 2
        For j = LOGFILE_ROW To PDFFILE1_ROW
            'MsgBox j
            If Cells(j, i) = "" Then Debug.Print "empty"
            FileCurr.CreateByParts ThisWorkbook.Path, Cells(j, i)
            Set OutputCellCurr = Cells(j, i + 1)
            If FileCurr.FileExists = True Then
                'MsgBox "'" & FileCurr.getRelativePath(Workbookfile.FolderPath & "\") & "' " & FileCurr.DateLastModified & " to cell " & OutputCellCurr.Address
                OutputCellCurr.Value = FileCurr.DateLastModified
            ElseIf Cells(j, i) = "" Then
                OutputCellCurr.Value = ""
            Else
                OutputCellCurr.Value = "(missing)"
            End If
        Next j
    Next i
    
    ' Check executables
    Dim ExecPaths As Range
    Set ExecPaths = Range(Range(R5_PATH), Range(GHOSTSCRIPT_PATH))
    For i = 1 To ExecPaths.Rows.Count
        FileCurr.CreateByParts ExecPaths(i)
        If FileCurr.FileExists = True Then
            Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = "OK"
        ElseIf ExecPaths(i) = "" Then
            Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = ""
        Else
            Cells(ExecPaths(i).Row, ExecPaths(i).Column + 4) = "(missing)"
        End If
    Next i
    

    ' Check global files
    Dim GlobalFilePaths As Range
    Set GlobalFilePaths = Range(Range(GLOBAL_STRIPFILE), Range(GLOBAL_STRIPFILE_FORCES))
    For i = 1 To GlobalFilePaths.Rows.Count
        FileCurr.CreateByParts ThisWorkbook.Path, GlobalFilePaths(i)
        If FileCurr.FileExists = True Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = FileCurr.DateLastModified
        ElseIf GlobalFilePaths(i) = "" Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = ""
        Else
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = "(missing)"
        End If
    Next i
    

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







Sub ShowFileList()

    ' Loopa igenom alla case. Lista alla filer i katalogen och loopa igenom dessa en efter en
    ' Ta bort den aktuella filen om den inte är en trackad fil i något case (loopa igenom en gång till)
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
