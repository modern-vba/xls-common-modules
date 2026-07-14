VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
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
'! Enumerator interface for iterating over a collection.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IEnumerator", Description:="This class is intended for interfaces."
End Sub

'* Current element.
Public Property Get Current() As Variant
End Property

'* Current index.
Public Property Get Index() As Long
End Property

'* Collection being enumerated.
Public Property Get Target() As Variant
End Property

'* Whether enumeration is in reverse order.
Public Property Get Descending() As Boolean
End Property

'* Whether this enumerator is read-only.
Public Property Get IsReadOnly() As Boolean
End Property

'* Moves to the next element.
'*
'* @return Boolean True if the next element exists; otherwise, False.
'*
'* @details
'* Moves the current position to the next element. Returns False if the end of the list has been reached.
'* The initial value is one position before the start of the collection, so calling this once points to the first element.
Public Function MoveNext() As Boolean
End Function

'* Resets enumeration.
'*
'* @details
'* Returns the current enumeration position to one position before the start of the list.
Public Sub Reset()
End Sub

'* Skips to the specified position.
'*
'* @param TargetIndex Destination index.
'*
'* @details
'* Advances the current position to the specified index. Raises an error if the index is before the current position or outside the list range.
Public Sub SkipTo(ByVal TargetIndex As Long)
End Sub

'* Removes the current element.
'*
'* @return Variant The removed element.
'*
'* @details
'* Removes the current element. Internally, this calls the Remove(ItemIndex) method of the target collection.
'* After this operation, the position indicated by this enumerator moves back by one.
'* When used with a collection that does not guarantee order, behavior of the enumerator after this operation is not guaranteed.
'* Raises an error when IsReadOnly is True.
Public Function Remove() As Variant
End Function

'* Updates the current element.
'*
'* @param ItemObject Value of the element to update.
'*
'* @details
'* Updates the current element. Internally, this calls the Update(ItemIndex, ItemObject) method of the target collection.
'* Raises an error when the implementation does not support writes, such as an array enumerator.
'* When used with a collection that does not guarantee order, behavior of the enumerator after this operation is not guaranteed.
'* Raises an error when IsReadOnly is True.
Public Sub Update(ByVal ItemObject As Variant)
End Sub

