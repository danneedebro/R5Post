Attribute VB_Name = "R5PostSheetManipulation"
Option Explicit

Sub ResetFormat()
' Action: Restores the format conditions

    Dim i As Integer, j As Integer
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet
    

    '
    ' Format rows and colums
    Dim CaseRange As Range
    Set CaseRange = sht.Range(sht.Cells(LOGFILE_ROW, CASE_COLUMN_START), sht.Cells(PDFFILE1_ROW, CASE_COLUMN_END))
    CaseRange.FormatConditions.Delete
    With CaseRange
        .FormatConditions.Add Type:=xlExpression, Formula1:="=(RAD()+1)/2=AVRUNDA.NEDÅT((RAD()+1)/2;0)"
         With .FormatConditions(.FormatConditions.Count)
             .SetFirstPriority
             .StopIfTrue = False
             .Interior.Color = RGB(242, 242, 242)
         End With
         
         .FormatConditions.Add Type:=xlExpression, Formula1:="=RAD()/2=AVRUNDA.NEDÅT(RAD()/2;0)"
         With .FormatConditions(.FormatConditions.Count)
             .SetFirstPriority
             .StopIfTrue = False
             .Interior.Color = RGB(255, 255, 255)
         End With
    End With
    
    
    '
    ' Format text red if dates of the files doesn't follow in the right order
    Dim currCell As Range
    Dim formulaBase As String, formula As String, colLetter As String
    
    For i = CASE_COLUMN_START To CASE_COLUMN_END Step 2
        Debug.Print ""
    
        For j = OUTPUTFILE_ROW To PDFFILE1_ROW
            Set currCell = sht.Range(sht.Cells(j, i + 1), sht.Cells(j, i + 1))
'            currCell.FormatConditions.Delete
            
            colLetter = Split(currCell.Address(True, False), "$")(0)
            
            
            Select Case j
                Case OUTPUTFILE_ROW, RSTFILE_ROW
                    formulaBase = "={COL}{JROW}<{COL}{INP_ROW}"
                
                Case DMXFILE_ROW, STRFILE1_ROW
                    formulaBase = "=ELLER({COL}{JROW}<{COL}{INP_ROW};{COL}{JROW}<{COL}{RST_ROW})"
                                    
                Case PSFILE1_ROW
                    formulaBase = "=ELLER({COL}{JROW}<{COL}{INP_ROW};{COL}{JROW}<{COL}{RST_ROW};{COL}{JROW}<{COL}{STR_ROW})"
                
                Case PDFFILE1_ROW
                    formulaBase = "=ELLER({COL}{JROW}<{COL}{INP_ROW};{COL}{JROW}<{COL}{RST_ROW};{COL}{JROW}<{COL}{STR_ROW};{COL}{JROW}<{COL}{PS_ROW})"
                
                Case Else
                    GoTo NextFileRow
            End Select
            
            formula = Replace(formulaBase, "{COL}", colLetter)
            formula = Replace(formula, "{JROW}", CStr(j))
            formula = Replace(formula, "{INP_ROW}", CStr(INPUTFILE_ROW))
            formula = Replace(formula, "{RST_ROW}", CStr(RSTFILE_ROW))
            formula = Replace(formula, "{STR_ROW}", CStr(STRFILE1_ROW))
            formula = Replace(formula, "{PS_ROW}", CStr(PSFILE1_ROW))
            
            With currCell
                .FormatConditions.Add Type:=xlExpression, Formula1:=formula
                With .FormatConditions(.FormatConditions.Count)
                    .SetFirstPriority
                    .StopIfTrue = False
                    .Font.Color = RGB(255, 0, 0)
                End With
            End With
NextFileRow:
        Next j
        
    
        
    Next i

End Sub
