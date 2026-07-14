VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ObjectSetEquatableStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEquatable

' #############################################################################
'!
'! @brief
'! Stub for verifying ObjectSet IEquatable decisions.
'!
' #############################################################################

Private pIdentityKey As String

'* Identity check key.
Public Property Get IdentityKey() As String
    IdentityKey = pIdentityKey
End Property
'* Identity check key.
Public Property Let IdentityKey(ByVal NewValue As String)
    pIdentityKey = NewValue
End Property

Private Function IEquatable_Equals(ByVal OtherObject As Object) As Boolean
    If TypeName(OtherObject) <> TypeName(Me) Then Exit Function

    Dim other_item As Test_ObjectSetEquatableStub
    Set other_item = OtherObject

    IEquatable_Equals = (pIdentityKey = other_item.IdentityKey)
End Function

Private Function IEquatable_GetIdentityString() As String
    IEquatable_GetIdentityString = pIdentityKey
End Function
