VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! Interface indicating that an object can be represented as a string.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IStringable", Description:="This class is intended for interfaces."
End Sub

'* Returns the string representation of this object.
'*
'* @details
'* Returns the string representation of this object.
Public Function ToString() As String
End Function
