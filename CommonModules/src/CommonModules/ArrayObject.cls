VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Low-level array wrapper that exposes a minimal public API around a Variant array.
'!
' #############################################################################

Private pArray() As Variant

'* The lower bound of the currently held array.
Public Property Get LowerBound() As Long
    LowerBound = LBound(pArray)
End Property

'* The upper bound of the currently held array.
Public Property Get UpperBound() As Long
    UpperBound = UBound(pArray)
End Property

'* Indexer for this collection. It performs no bounds check and exposes the same errors as VBA array access.
Public Property Get Item(ByVal ItemIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(pArray(ItemIndex)) Then
        Set Item = pArray(ItemIndex)
    Else
        Item = pArray(ItemIndex)
    End If
End Property
Public Property Let Item(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call Update(ItemIndex, ItemObject)
End Property
Public Property Set Item(ByVal ItemIndex As Long, ByVal ItemObject As Object)
    Call Update(ItemIndex, ItemObject)
End Property

Private Sub Class_Initialize()
    pArray = EmptyVariantArray()
End Sub

'Private Sub Class_Terminate()
'    On Error Resume Next
'    Debug.Print "Terminate: ArrayObject(0x" & Hex(ObjPtr(Me)) & ")"
'
'    Dim idx As Long
'    For idx = LBound(pArray) To UBound(pArray)
'        pArray(idx) = Empty
'    Next idx
'End Sub

'* Resets the array bounds.
'*
'* @param LowerBound The new lower bound.
'* @param UpperBound The new upper bound.
'* @param PreserveItems Whether to preserve existing elements.
'*
'* @details
'* Immediately after creation, this holds the 0-based empty array from `EmptyVariantArray()` (`LBound = 0` / `UBound = -1`).
'* When `PreserveItems:=False`, this follows the same contract as VBA `ReDim`.
'* When `PreserveItems:=True`, this follows the same contract as VBA `ReDim Preserve`: the lower bound of the existing array cannot be changed, and only the upper bound can be changed.
Public Sub ReDimArray(ByVal LowerBound As Long, ByVal UpperBound As Long, Optional ByVal PreserveItems As Boolean = False)
    If PreserveItems Then
        ReDim Preserve pArray(LowerBound To UpperBound)
    Else
        ReDim pArray(LowerBound To UpperBound)
    End If
End Sub

'* Updates the element at the specified index.
'*
'* @param ItemIndex The element index.
'* @param ItemObject The element to set.
'*
'* @details
'* Performs no bounds check and exposes the same errors as VBA array access.
'* Objects are set by reference; all other values are set by value.
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then
        Set pArray(ItemIndex) = ItemObject
    Else
        pArray(ItemIndex) = ItemObject
    End If
End Sub
