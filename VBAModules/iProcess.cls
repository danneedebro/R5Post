VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES

'
'
' PROPERTIES
Public Property Get UpToDate() As Boolean
End Property

Public Property Get FilesRequired() As CollectionFileList
End Property

Public Property Get FilesBefore() As CollectionFileList
End Property

Public Property Get FilesAfter() As CollectionFileList
End Property
'
'
' CLASS EVENT METHODS

'
'
' CLASS METHODS
Public Function GetShellCommand(Optional WorkingFolder As String = "") As String
End Function

