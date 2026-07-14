VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetVirtualTableEnumerator"
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
'! Enumerator class for enumerating rows of WorksheetVirtualTable.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pDescending As Boolean
Private pTable As WorksheetVirtualTable
Private pLength As Long
Private pIndex As Long

'* Whether to enumerate in reverse order.
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
    IsReadOnly = True
End Property
Private Property Get IEnumerator_IsReadOnly() As Boolean
    IEnumerator_IsReadOnly = IsReadOnly
End Property

'* Row ObjectDictionary at the current position.
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    Set Current = pTable.Item(pIndex)
End Property
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = Current
End Property

'* 0-based index of the current position.
Public Property Get Index() As Long
    Call pCheckInit
    Index = pIndex
End Property
Private Property Get IEnumerator_Index() As Long
    IEnumerator_Index = Index
End Property

'* Virtual table being enumerated.
Public Property Get Target() As Variant
    Call pCheckInit
    Set Target = pTable
End Property
Private Property Get IEnumerator_Target() As Variant
    Set IEnumerator_Target = Target
End Property

'* Initializes this enumerator.
'*
'* @param TargetTable `WorksheetVirtualTable` to enumerate.
'* @param Descending [Optional] Whether to enumerate in descending order.
'*
'* @details
'* Raises an error when `TargetTable` is Nothing or when an initialized enumerator is reinitialized.
Public Sub Initialize( _
        ByVal TargetTable As WorksheetVirtualTable, _
        Optional ByVal Descending As Boolean = False)

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.Initialize", "The object is already initialized."
    If TargetTable Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.Initialize", "The enumeration target is Nothing."

    Set pTable = TargetTable
    pLength = TargetTable.RowCount
    pDescending = Descending
    pIsInitialized = True

    Call Reset
End Sub

'* Moves to the next element.
'*
'* @return True if the next element exists; otherwise, False.
'*
'* @details
'* The initial position is one before the first element. In descending order, starts one after the last element.
Public Function MoveNext() As Boolean
    Call pCheckInit
    If Descending Then
        If 0 < pIndex Then
            pIndex = pIndex - 1
            MoveNext = True
        Else
            pIndex = -1
            MoveNext = False
        End If
    Else
        If pIndex < pLength - 1 Then
            pIndex = pIndex + 1
            MoveNext = True
        Else
            pIndex = pLength
            MoveNext = False
        End If
    End If
End Function
Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext()
End Function

'* Resets enumeration.
'*
'* @details
'* Returns the current position to one before the first element for ascending order, or one after the last element for descending order.
Public Sub Reset()
    Call pCheckInit
    If Descending Then
        pIndex = pLength
    Else
        pIndex = -1
    End If
End Sub
Private Sub IEnumerator_Reset()
    Call Reset
End Sub

'* Skips to the specified 0-based index.
'*
'* @param TargetIndex Destination 0-based index.
'*
'* @details
'* Raises an error when an out-of-range index or an index before the current position is specified.
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < 0 Or pLength <= TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.SkipTo", "The destination is outside the valid range. (min: 0 / max: " & (pLength - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

'* Attempts to remove the current element.
'*
'* @return Removed element.
'*
'* @details
'* `WorksheetVirtualTableEnumerator` is read-only, so this always raises an error.
Public Function Remove() As Variant
    Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.Remove", "This collection is read-only and does not support removing elements."
End Function
Private Function IEnumerator_Remove() As Variant
    Set IEnumerator_Remove = Remove()
End Function

'* Attempts to update the current element.
'*
'* @param ItemObject Element to update.
'*
'* @details
'* `WorksheetVirtualTableEnumerator` is read-only, so this always raises an error.
Public Sub Update(ByVal ItemObject As Variant)
    Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.Update", "This collection is read-only and does not support updating elements."
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.pCheckInit", "The object is not initialized."
End Sub

Private Sub pCheckEndOfList()
    If pIndex < 0 Or pLength <= pIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTableEnumerator.pCheckEndOfList", "There is no current element. (index: " & pIndex & " / min: 0 / max: " & (pLength - 1) & ")"
    End If
End Sub
