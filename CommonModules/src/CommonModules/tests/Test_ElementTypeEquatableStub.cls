VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ElementTypeEquatableStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IElementTypeProvider
Implements IEquatable

' #############################################################################
'!
'! @brief
'! Stub for verifying element type self-reporting and IEquatable.
'!
' #############################################################################

Private pElementTypeKey As String
Private pIdentityKey As String

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

'* Identity check key.
Public Property Get IdentityKey() As String
    IdentityKey = pIdentityKey
End Property
'* Identity check key.
Public Property Let IdentityKey(ByVal NewValue As String)
    pIdentityKey = NewValue
End Property

Private Function IEquatable_Equals(ByVal OtherObject As Object) As Boolean
    If Not TypeOf OtherObject Is Test_ElementTypeEquatableStub Then Exit Function

    Dim other_item As Test_ElementTypeEquatableStub
    Set other_item = OtherObject

    IEquatable_Equals = (pIdentityKey = other_item.IdentityKey)
End Function

Private Function IEquatable_GetIdentityString() As String
    IEquatable_GetIdentityString = pIdentityKey
End Function
