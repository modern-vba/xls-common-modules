VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IElementTypeProvider"
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
'! Interface indicating that an object can self-report the element type used when it is handled in a collection.
'!
'! A typed element collection reads ElementTypeKey from an object that implements this interface and determines which element type contract to use for that object.
'! The implementing class is responsible for ensuring that ElementTypeKey represents a type that can actually be assigned.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IElementTypeProvider", Description:="This class is intended for interfaces."
End Sub

'* Element type contract name.
'*
'* @return Element type contract name.
'*
'* @details
'* Returns a name of 31 characters or fewer, composed only of characters that can be used in class module names.
'* The first character can be a letter; the second and later characters can be letters, digits, or underscores.
'* The implementing class is responsible for ensuring that this name represents a type that can actually be assigned.
Public Property Get ElementTypeKey() As String
End Property
