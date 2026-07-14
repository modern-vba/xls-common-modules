VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IEnumerator

' #############################################################################
'!
'! @brief
'! Enumerator class for iterating over a collection.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pList As Object
Private pArray As Variant
Private pDescending As Boolean
Private pIsReadOnly As Boolean
Private pIsArray As Boolean
Private pIndex As Long
Private pMinIndex As Long

Public Property Get Descending() As Boolean
    Call pCheckInit
    Descending = pDescending
End Property
Private Property Get IEnumerator_Descending() As Boolean
    IEnumerator_Descending = Descending
End Property

'* Whether this enumerator is read-only.
Public Property Get IsReadOnly() As Boolean
    Call pCheckInit
    IsReadOnly = pIsReadOnly
End Property
Private Property Get IEnumerator_IsReadOnly() As Boolean
    IEnumerator_IsReadOnly = IsReadOnly
End Property

' Current element.
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    If pIsArray Then
        If IsObject(pArray(pIndex)) Then
            Set Current = pArray(pIndex)
        Else
            Current = pArray(pIndex)
        End If
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set Current = pList.Item(pIndex)
        Else
            Current = pList.Item(pIndex)
        End If
    End If
End Property
Private Property Get IEnumerator_Current() As Variant
    If IsObject(Current) Then
        Set IEnumerator_Current = Current
    Else
        IEnumerator_Current = Current
    End If
End Property

' Current index.
Public Property Get Index() As Long
    Call pCheckInit
    Index = pIndex
End Property
Private Property Get IEnumerator_Index() As Long
    IEnumerator_Index = Index
End Property

' Collection being enumerated.
Public Property Get Target() As Variant
    Call pCheckInit
    If pIsArray Then
        Target = pArray
    Else
        Set Target = pList
    End If
End Property
Private Property Let Target(ByVal NewValue As Variant)
    If Not IsArray(NewValue) Then
        Err.Raise vbObjectError + 1, "Class Enumerator.Target", "The value is not an array."
    End If

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, NewValue)
    If LBound(lbound_arr) <> UBound(lbound_arr) Then
        Err.Raise vbObjectError + 1, "Class Enumerator.Target", "Array enumerators support only one-dimensional arrays."
    End If

    pArray = NewValue
    pMinIndex = lbound_arr(LBound(lbound_arr))
    pIsArray = True
End Property
Private Property Set Target(ByVal NewValue As Object)
    Set pList = NewValue
    pMinIndex = 0
End Property

Private Property Get IEnumerator_Target() As Variant
    If IsObject(Target) Then
        Set IEnumerator_Target = Target
    Else
        IEnumerator_Target = Target
    End If
End Property

'* Initializes this object.
'*
'* @param TargetCollection Collection or one-dimensional array to enumerate. Nothing, uninitialized arrays, and multidimensional arrays are errors.
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @param IsReadOnly [Optional] Whether to make the enumerator read-only.
'*
'* @details
'* Initializes this object. Objects are treated as having Count and Item members.
Public Sub Initialize(ByVal TargetCollection As Variant, Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class Enumerator.Initialize", "The object is already initialized."

    If IsObject(TargetCollection) Then
        If TargetCollection Is Nothing Then Err.Raise vbObjectError + 1, "Class Enumerator.Initialize", "The enumeration target is Nothing."
        Set Target = TargetCollection
    Else
        Target = TargetCollection
    End If

    pDescending = Descending
    pIsReadOnly = IsReadOnly

    pIsInitialized = True

    Call Reset
End Sub

' Moves to the next element.
'
' @return Boolean True if the next element exists; otherwise, False.
'
' @details
' Moves the current position to the next element. Returns False if the end of the list has been reached.
' The initial position is one before the start of the collection, so calling this once points to the first element.
Public Function MoveNext() As Boolean
    Call pCheckInit

    If Descending Then
        If pMinIndex < pIndex Then
            pIndex = pIndex - 1
            MoveNext = True
        Else
            pIndex = pMinIndex - 1
            MoveNext = False
        End If
    Else
        If pIndex < pGetCount() - 1 Then
            pIndex = pIndex + 1
            MoveNext = True
        Else
            pIndex = pGetCount()
            MoveNext = False
        End If
    End If
End Function
Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext()
End Function

' Resets enumeration.
'
' @details
' Returns the current enumeration position to one before the start of the list (or one after the end of the list when descending).
Public Sub Reset()
    Call pCheckInit
    If Descending Then
        pIndex = pGetCount
    Else
        pIndex = pMinIndex - 1
    End If
End Sub
Private Sub IEnumerator_Reset()
    Call Reset
End Sub

' Skips to the specified position.
'
' @param TargetIndex Destination index.
'
' @details
' Advances the current position to the specified index. Raises an error if the index is before the current position or outside the list range.
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < pMinIndex Or pGetCount() <= TargetIndex Then Err.Raise vbObjectError + 1, "Class Enumerator.SkipTo", "The destination is outside the valid range. (min: " & pMinIndex & " / max: " & (pGetCount() - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class Enumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class Enumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

' Removes the current element.
'
' @return Variant The removed element.
'
' @details
' Removes the current element. Internally, the Remove(ItemIndex) method of the target collection is called.
' After this operation, the position indicated by this enumerator moves back by one.
' When used with a collection that does not guarantee order, behavior of the enumerator after this operation is not guaranteed.
Public Function Remove() As Variant
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator.Remove", "This operation cannot be used with arrays."
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set Remove = pList.Remove(pIndex)
        Else
            Remove = pList.Remove(pIndex)
        End If
    End If

    If Not Descending Then
        pIndex = pIndex - 1
    End If
End Function
Private Function IEnumerator_Remove() As Variant
    ' Handled by copy/paste because object and primitive type handling cannot be unified.
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator.IEnumerator_Remove", "This operation cannot be used with arrays."
    Else
        If IsObject(pList.Item(pIndex)) Then
            Set IEnumerator_Remove = pList.Remove(pIndex)
        Else
            IEnumerator_Remove = pList.Remove(pIndex)
        End If
    End If

    If Not Descending Then
        pIndex = pIndex - 1
    End If
End Function

' Updates the current element.
'
' @param ItemObject Value of the element to update.
'
' @details
' Updates the current element. Internally, the Update(ItemIndex, ItemObject) method of the target collection is called.
' This operation cannot be used with array enumerators because updates cannot be reflected back to the caller's array.
' When used with a collection that does not guarantee order, behavior of the enumerator after this operation is not guaranteed.
Public Sub Update(ByVal ItemObject As Variant)
    Call pCheckInit
    Call pCheckEndOfList
    Call pCheckWritable

    If pIsArray Then
        Err.Raise vbObjectError + 1, "Class Enumerator.Update", "This operation cannot be used with arrays."
    Else
        Call pList.Update(pIndex, ItemObject)
    End If
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class Enumerator.pCheckInit", "The object is not initialized."
End Sub

Private Sub pCheckWritable()
    If pIsReadOnly Then Err.Raise vbObjectError + 1, "Class Enumerator.pCheckWritable", "A read-only enumerator cannot update or remove elements."
End Sub

Private Function pGetCount() As Long
    If pIsArray Then
        If IsEmptyArray(pArray) Then
            pGetCount = 0
        Else
            pGetCount = UBound(pArray) + 1
        End If
    Else
        pGetCount = pList.Count
    End If
End Function

Private Sub pCheckEndOfList()
    Dim count_var As Long
    count_var = pGetCount()

    If pIndex < pMinIndex Or count_var <= pIndex Then
        Err.Raise vbObjectError + 1, "Class Enumerator.pCheckEndOfList", "There is no current element. (index: " & pIndex & " / min: " & pMinIndex & " / max: " & (count_var - 1) & ")"
    End If
End Sub
