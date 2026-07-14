VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ElementTypeProviderStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IElementTypeProvider

' #############################################################################
'!
'! @brief
'! Stub for verifying element type self-reporting.
'!
' #############################################################################

Private pElementTypeKey As String

'* Element type contract key.
Public Property Get ElementTypeKey() As String
    ElementTypeKey = pElementTypeKey
End Property
'* Element type contract key.
Public Property Let ElementTypeKey(ByVal NewValue As String)
    pElementTypeKey = NewValue
End Property
Private Property Get IElementTypeProvider_ElementTypeKey() As String
    IElementTypeProvider_ElementTypeKey = pElementTypeKey
End Property
