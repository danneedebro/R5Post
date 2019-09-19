VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pFiles As Collection

'
'
' PROPERTIES

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pFiles = New Collection
End Sub

Public Sub Add(newFile As R5PostFileObject)
' Action: Adds file to list
'
    pFiles.Add newFile
End Sub

Public Sub AddList(List As CollectionFileList)
' Action: Adds the files in CollectionFileList to list
'
    Dim i As Integer
    For i = 1 To List.Count
        pFiles.Add List(i)
    Next i
End Sub


Public Sub AddMany(ParamArray Args() As Variant)
' Action: Adds many files att once
'
    Dim i As Integer
    For i = LBound(Args) To UBound(Args)
        pFiles.Add Args(i)
    Next i
End Sub

Public Function Item(index As Variant) As R5PostFileObject
Attribute Item.VB_UserMemId = 0
    Set Item = pFiles(index)
End Function

Public Function Count() As Integer
    Count = pFiles.Count
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

