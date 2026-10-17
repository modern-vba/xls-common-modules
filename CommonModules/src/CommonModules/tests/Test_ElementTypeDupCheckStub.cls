VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ElementTypeDupCheckStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IElementTypeProvider
Implements IDuplicateCheckable

' #############################################################################
'!
'! @brief
'! Stub for verifying element type self-reporting and IDuplicateCheckable.
'!
' #############################################################################

Private pElementTypeKey As String
Private pDuplicateKey As String

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

'* Duplicate check key.
Public Property Get DuplicateKey() As String
    DuplicateKey = pDuplicateKey
End Property
'* Duplicate check key.
Public Property Let DuplicateKey(ByVal NewValue As String)
    pDuplicateKey = NewValue
End Property

Private Function IDuplicateCheckable_GetKey() As String
    IDuplicateCheckable_GetKey = pDuplicateKey
End Function

Private Function IDuplicateCheckable_IsDuplicateOf(ByVal OtherObject As Object) As Boolean
    If Not TypeOf OtherObject Is Test_ElementTypeDupCheckStub Then Exit Function

    Dim other_item As Test_ElementTypeDupCheckStub
    Set other_item = OtherObject

    IDuplicateCheckable_IsDuplicateOf = (pDuplicateKey = other_item.DuplicateKey)
End Function
