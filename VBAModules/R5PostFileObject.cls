VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "R5PostFileObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pFullPath As String
Private fso As Scripting.FileSystemObject

'
'
' PROPERTIES
Public Property Get Filename() As String
' Action: Returns the filename without the path
'         "C:\Work\File.dat" = "File.dat"
'
    Filename = fso.GetFileName(pFullPath)
End Property

Public Property Get Basename() As String
' Action: Returns the base name, i.e the filename without path and extension
'         "C:\Work\File.dat" = "File"
'
    Basename = fso.GetBaseName(pFullPath)
End Property

Public Property Get FullPath() As String
' Action: Returns the full path of the file including the file and extension
'         "C:\Work\File.dat" = "C:\Work\File.dat"
'
    FullPath = pFullPath
End Property

Public Property Get Extension() As String
' Action: Returns the extension
'         "C:\Work\File.dat" = "dat"
'
    Extension = fso.GetExtensionName(pFullPath)
End Property

Public Property Get FileExists() As Boolean
' Action: Returns true if file exists, false otherwise
'
    FileExists = fso.FileExists(pFullPath)
End Property

Public Property Get FolderExists() As Boolean
' Action: Returns true if folder exists, false otherwise
'
    FolderExists = fso.FolderExists(FolderPath)
End Property

Public Property Get DateLastModified() As Date
    
    If FileExists = True Then
        Dim fo
        Set fo = fso.GetFile(pFullPath)
        DateLastModified = fo.DateLastModified
    Else
        Exit Property
    End If

End Property

Public Property Get FolderPath() As String
' Action: Returns the path (without backslash)
'         "C:\Work\File.dat" = "C:\Work"
'
    FolderPath = fso.GetParentFolderName(pFullPath)
End Property

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set fso = New FileSystemObject
End Sub

Public Sub Create(ByVal AmbigiousPath As String)
    pFullPath = fso.GetAbsolutePathName(AmbigiousPath)
End Sub

Public Sub CreateByParts(ParamArray PartsOfPaths() As Variant)
    Dim NewPath As String
    Dim i As Integer
    For i = LBound(PartsOfPaths) To UBound(PartsOfPaths)
        NewPath = fso.BuildPath(NewPath, CStr(PartsOfPaths(i)))
    Next i
    
    pFullPath = NewPath
End Sub
'
'
' CLASS METHODS



Public Function GetRelativePath(currPath) As String
' Function that returns the relative path of a file compared to the current working folder
'
' INPUT:    filePath    - Full path to file or folder to get a relative path for. If folder it must end with "\"
'           currPath    - Current working folder
'
' OUTPUT:   relativePath
'
' ex)   getRelativePath("c:\work\hello.txt", "c:\work\folder1\")            - returns "..\hello.txt"
'       getRelativePath("c:\work\folder2\hello.txt", "c:\work\folder1\")    - returns "..\folder2\hello.txt"
'       getRelativePath("c:\work\folder1\hello.txt", "c:\work\")            - returns "folder1\hello.txt"
'       getRelativePath("c:\work\folder1\", "c:\work\")                     - returns "folder1\"
'       getRelativePath("h:\script\", "c:\work\")                           - returns "h:\script\"
    Dim filePath As String
    filePath = pFullPath

    Dim FullPath, currFile, relativePath
    
    ' if last sign is a backslash then filePath is a directory
    If Right(FullPath, 1) = "\" Then
        FullPath = filePath
        currFile = ""
    Else
        FullPath = Left(filePath, InStrRev(filePath, "\"))
        currFile = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    End If
    
    ' Split 'fullPath' and 'currPath' to see what level folders start to diverge
    Dim tmp1, tmp2
    tmp1 = Split(FullPath, "\")
    tmp2 = Split(currPath, "\")
    
    Dim N As Integer, i As Integer, j As Integer
    Dim resFullPath As String
    Dim resCurrPath As String
    If UBound(tmp1) < UBound(tmp2) Then N = UBound(tmp1) Else N = UBound(tmp2)
    For i = 0 To N
        ' if words don't match, return the residual paths of 'fullPath' and 'currPath'
        If UCase(tmp1(i)) <> UCase(tmp2(i)) Or i = N Then  '
            resFullPath = ""
            For j = i To UBound(tmp1) - 1
                resFullPath = resFullPath + tmp1(j) + "\"
            Next j
            resCurrPath = ""
            For j = i To UBound(tmp2) - 1
                resCurrPath = resCurrPath + tmp2(j) + "\"
            Next j
            If i = 0 Then  ' Catches error when on diffrent drives
                resCurrPath = ""
                resFullPath = FullPath
            End If
            Exit For
            'i = 99999   ' end loop
        End If
    Next i
    
    Dim tmp
    Dim tmpStr
    If resCurrPath <> "" Then
        tmp = Split(resCurrPath, "\")
        tmpStr = ""
        For i = 0 To UBound(tmp) - 1
            tmpStr = tmpStr + "..\"
        Next i
        resCurrPath = tmpStr
    End If
    
    GetRelativePath = resCurrPath + resFullPath + currFile

End Function

'
'
' STATIC METHODS


