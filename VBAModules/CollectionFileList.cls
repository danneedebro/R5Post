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
    pFiles.Add newFile
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

