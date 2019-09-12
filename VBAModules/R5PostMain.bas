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

Const CURRENTSHEET = "E1"
Const DEBUG_FLAG = "B7"

Private Enum TypeOfAction
    Calc
    CalcStripDemux
    StripDemux
    Post
    CalcAndPost
    ActionPs2pdf
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
            MsgBox fileSelected.getRelativePath(ThisWorkbook.Path & "\")
        End If
        
    End With
    
    

End Sub


Private Sub CalculateOrPost(Cellcolumn As Long, Action As TypeOfAction)
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
    Dim matlabPath As New R5PostFileObject
    Dim PlottaStripfilPath As New R5PostFileObject
    Dim R2DMXPath As New R5PostFileObject
    Dim GhostScriptPath As New R5PostFileObject
    
    R5Path.Create Range(R5_PATH)
    R5SteamTable.Create R5Path.FolderPath & "\tpfh2onew"
    R5SteamTableOld.Create R5Path.FolderPath & "\tpfh2o"
    matlabPath.Create Range(MATLAB_PATH)
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
    
    With ThisWorkbook.ActiveSheet
        InputFile.CreateByParts ThisWorkbook.Path, .Cells(INPUTFILE_ROW, Cellcolumn)
        Outputfile.CreateByParts ThisWorkbook.Path, .Cells(OUTPUTFILE_ROW, Cellcolumn)
        Restartfile.CreateByParts ThisWorkbook.Path, .Cells(RSTFILE_ROW, Cellcolumn)
        Demuxfile.CreateByParts ThisWorkbook.Path, .Cells(DMXFILE_ROW, Cellcolumn)
        If .Cells(STRIPFILE1_ROW, Cellcolumn) = "" Then
            StripRequestfile1.CreateByParts ThisWorkbook.Path, Range(GLOBAL_STRIPFILE)
        Else
            StripRequestfile1.CreateByParts ThisWorkbook.Path, .Cells(STRIPFILE1_ROW, Cellcolumn)
        End If
        If .Cells(PARAMFILE_ROW, Cellcolumn) = "" Then
            Paramfile.CreateByParts ThisWorkbook.Path, .Range(GLOBAL_PARAMFILE)
        Else
            Paramfile.CreateByParts ThisWorkbook.Path, .Cells(PARAMFILE_ROW, Cellcolumn)
        End If
        
        Stripfile1.CreateByParts ThisWorkbook.Path, .Cells(STRFILE1_ROW, Cellcolumn)
        PSfile1.CreateByParts ThisWorkbook.Path, .Cells(PSFILE1_ROW, Cellcolumn)
        PDFfile1.CreateByParts ThisWorkbook.Path, .Cells(PDFFILE1_ROW, Cellcolumn)
        Logfile.CreateByParts ThisWorkbook.Path, .Cells(LOGFILE_ROW, Cellcolumn)
        
        ' Define parameters
        Dim PlotTitle As String
        Dim PlotTimeMin As Double
        Dim PlotTimeMax As Double
        
        PlotTitle = .Cells(TITLE_ROW, Cellcolumn)
        PlotTimeMin = .Cells(TMIN_ROW, Cellcolumn)
        PlotTimeMax = .Cells(TMAX_ROW, Cellcolumn)
    End With
    
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
    MatlabPlottaStripfil.Create matlabPath, PlottaStripfilPath, Stripfile1, Paramfile, PSfile1, PlotTitle, PlotTimeMin, PlotTimeMax
    
    ' Create process that converts a .ps-file to pdf using ghostscript
    Dim GhostScript As New ProcessPs2Pdf
    GhostScript.Create GhostScriptPath, PSfile1, PDFfile1
    
    ' Create process chain
    Dim Calculate As New MainProcessChain
    
    Debug.Print InputFile.FullPath
    
    Dim questionString As String, questionTitle As String
    
    If Action = Calc Then
        Calculate.Add R5Calc
        questionTitle = "Perform calculation?"
        questionString = "Perform RELAP5-calculation on '" & InputFile.FullPath & "'?"
        'answ = MsgBox("Perform RELAP5-calculation on '" & InputFile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform calculation?")
        'If answ <> vbYes Then Exit Sub
    ElseIf Action = CalcAndPost Then
        Calculate.Add R5Calc
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        Calculate.Add MatlabPlottaStripfil
        questionTitle = "Perform calculation+postprocessing?"
        questionString = "Perform RELAP5-calculation and postprocessing on '" & InputFile.FullPath & "'?"
        'answ = MsgBox("Perform RELAP5-calculation and postprocessing on '" & InputFile.FullPath & "'?", vbQuestion + vbYesNoCancel, "Perform calculation+postprocessing?")
        'If answ <> vbYes Then Exit Sub
    ElseIf Action = Post Then
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        Calculate.Add MatlabPlottaStripfil
        questionTitle = "Perform postprocessing?"
        questionString = "Perform RELAP5-postprocessing on '" & Restartfile.FullPath & "'?"
    ElseIf Action = CalcStripDemux Then
        Calculate.Add R5Calc
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        questionTitle = "Perform calculation+strip+demux?"
        questionString = "Perform RELAP5-calculation and strip+demux on '" & InputFile.FullPath & "'?"
    ElseIf Action = StripDemux Then
        Calculate.Add R5Strip
        Calculate.Add R2DMX
        questionTitle = "Perform strip+demux?"
        questionString = "Perform Strip and demux on '" & Restartfile.FullPath & "'?"
    ElseIf Action = ActionPs2pdf Then
        Calculate.Add GhostScript
        questionTitle = "Convert to pdf?"
        questionString = "Convert following ps-files to pdf '" & PSfile1.FullPath & "'?"
    End If
    
    If Calculate.ProcessChainOK = True Then
        questionString = questionString & vbNewLine & vbNewLine & Calculate.GetShellCommand(InputFile.FolderPath & "\")
        answ = MsgBox(questionString, vbQuestion + vbYesNoCancel, questionTitle)
        If answ <> vbYes Then Exit Sub
        
        Dim retval
        ChDir InputFile.FolderPath
        Dim stayOpen As String
        If Range(DEBUG_FLAG).Value = 1 Then stayOpen = "/K" Else stayOpen = "/C"
        retval = Shell("cmd /S " & stayOpen & " dir && timeout 1 && " & Calculate.GetShellCommand(InputFile.FolderPath & "\"), 1)
    Else
        MsgBox "Error"
    End If
        
End Sub


Sub RefreshFileDates()
' Action: Refreshes the file dates
'
    Dim i As Integer, j As Integer

    Range(CURRENTSHEET).Value = "[" & ThisWorkbook.FullName & "]'" & ActiveSheet.Name & "'"

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
            'MsgBox j
            If Cells(j, i) = "" Then Debug.Print "empty"
            fileCurr.CreateByParts ThisWorkbook.Path, Cells(j, i)
            Set OutputCellCurr = Cells(j, i + 1)
            If fileCurr.FileExists = True Then
                'MsgBox "'" & FileCurr.getRelativePath(Workbookfile.FolderPath & "\") & "' " & FileCurr.DateLastModified & " to cell " & OutputCellCurr.Address
                OutputCellCurr.Value = fileCurr.DateLastModified
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
        fileCurr.CreateByParts ExecPaths(i)
        If fileCurr.FileExists = True Then
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
        fileCurr.CreateByParts ThisWorkbook.Path, GlobalFilePaths(i)
        If fileCurr.FileExists = True Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = fileCurr.DateLastModified
        ElseIf GlobalFilePaths(i) = "" Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = ""
        Else
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = "(missing)"
        End If
    Next i
    

End Sub

Sub RefreshFileDatesNew()
' Action: Refreshes the file dates
'
    Dim i As Integer, j As Integer
    Dim fso As New FileSystemObject
    Dim folderCurr As Folder, folderStr As String
    Dim fileInFolder As File
    
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet

    Range(CURRENTSHEET).Value = "[" & ThisWorkbook.FullName & "]'" & ActiveSheet.Name & "'"

    ' Workbook path
    Dim Workbookfile As New R5PostFileObject
    Workbookfile.Create ThisWorkbook.FullName
    
    ' Define files
    Dim fileCurr As New R5PostFileObject
    Dim OutputCellCurr As Range
    
    ' Check calculation process files (*.i, *.rst, etc)
    Dim ColumnCurr As Integer
    For i = CASE_COLUMN_START To CASE_COLUMN_END Step 2
        folderStr = fso.GetParentFolderName(fso.BuildPath(ThisWorkbook.Path, sht.Cells(INPUTFILE_ROW, i)))
        folderCurr = fso.GetFolder(folderStr)
        
        
        For Each fileInFolder In folderCurr.Files
            
            For j = LOGFILE_ROW To PDFFILE1_ROW
                
                If fso.GetFileName(sht.Cells(j, i)) = fileInFolder.Name Then
                    sht.Cells(j, i) = fileInFolder.DateLastModified
                    Exit For
                End If
                
                Debug.Print "empty"
            fileCurr.CreateByParts ThisWorkbook.Path, Cells(j, i)
            Set OutputCellCurr = Cells(j, i + 1)
            If fileCurr.FileExists = True Then
                'MsgBox "'" & FileCurr.getRelativePath(Workbookfile.FolderPath & "\") & "' " & FileCurr.DateLastModified & " to cell " & OutputCellCurr.Address
                OutputCellCurr.Value = fileCurr.DateLastModified
            ElseIf Cells(j, i) = "" Then
                OutputCellCurr.Value = ""
            Else
                OutputCellCurr.Value = "(missing)"
            End If
        Next j
        
        Next fileCurr
    
        
    Next i
    
    ' Check executables
    Dim ExecPaths As Range
    Set ExecPaths = Range(Range(R5_PATH), Range(GHOSTSCRIPT_PATH))
    For i = 1 To ExecPaths.Rows.Count
        fileCurr.CreateByParts ExecPaths(i)
        If fileCurr.FileExists = True Then
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
        fileCurr.CreateByParts ThisWorkbook.Path, GlobalFilePaths(i)
        If fileCurr.FileExists = True Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = fileCurr.DateLastModified
        ElseIf GlobalFilePaths(i) = "" Then
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = ""
        Else
            Cells(GlobalFilePaths(i).Row, GlobalFilePaths(i).Column + 1) = "(missing)"
        End If
    Next i
    

End Sub


Sub fixLinks()
    ' Fixar länkar. Markera de celler som ska fixas och kör makrot
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
