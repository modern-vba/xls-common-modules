VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetRangeBoundsEnumerator"
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
'! Enumerator class for enumerating rows, columns, and cells of WorksheetRangeBounds.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pDescending As Boolean
Private pEnumerationMode As Long
Private pRangeBounds As WorksheetRangeBounds
Private pLength As Long
Private pIndex As Long

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

'* Enumeration mode.
Public Property Get EnumerationMode() As Long
    Call pCheckInit
    EnumerationMode = pEnumerationMode
End Property

' Current element.
Public Property Get Current() As Variant
    Call pCheckInit
    Call pCheckEndOfList

    Select Case pEnumerationMode
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL
        Set Current = pRangeBounds.Item(pIndex, ColumnDirection:=False)
     Case G_RANGE_ENUM_MODE_CELLS_VERTICAL
        Set Current = pRangeBounds.Item(pIndex, ColumnDirection:=True)
     Case G_RANGE_ENUM_MODE_COLUMNS
        Set Current = pRangeBounds.GetColumn(pIndex + 1)
     Case Else
        Set Current = pRangeBounds.GetRow(pIndex + 1)
    End Select
End Property
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = Current
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
    Set Target = pRangeBounds
End Property
Private Property Set Target(ByVal NewValue As Object)
    Set pRangeBounds = NewValue

    ' Calculate the length.
    Call pCalculateLength
End Property

Private Property Get IEnumerator_Target() As Variant
    Set IEnumerator_Target = Target
End Property

'* Initializes this object.
'*
'* @param TargetCollection WorksheetRangeBounds to enumerate. Nothing is an error.
'* @param EnumerationMode [Optional] Enumeration mode.
'* @param Descending [Optional] Whether to enumerate in descending order.
'*
'* @details
'* Initializes this object. TargetCollection is expected to be passed from WorksheetRangeBounds.GetEnumerator.
Public Sub Initialize( _
        ByVal TargetCollection As WorksheetRangeBounds, _
        Optional ByVal EnumerationMode As Long = G_RANGE_ENUM_MODE_ROWS, _
        Optional ByVal Descending As Boolean = False)

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.Initialize", "The object is already initialized."
    If TargetCollection Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.Initialize", "The enumeration target is Nothing."

    Set Target = TargetCollection
    Call pSetEnumerationMode(EnumerationMode)
    pDescending = Descending

    pIsInitialized = True

    Call Reset
End Sub

Private Sub pSetEnumerationMode(ByVal NewValue As Long)
    Call pCheckEnumerationMode(NewValue)
    pEnumerationMode = NewValue

    ' Recalculate the length according to the settings.
    Call pCalculateLength
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

' Resets enumeration.
'
' @details
' Returns the current enumeration position to one before the start of the list (or one after the end of the list when descending).
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

' Skips to the specified position.
'
' @param TargetIndex Destination index.
'
' @details
' Advances the current position to the specified index. Raises an error if the index is before the current position or outside the list range.
Public Sub SkipTo(ByVal TargetIndex As Long)
    Call pCheckInit
    If TargetIndex < 0 Or pLength <= TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.SkipTo", "The destination is outside the valid range. (min: 0 / max: " & (pLength - 1) & " / target: " & TargetIndex & ")"
    If Descending Then
        If pIndex < TargetIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    Else
        If TargetIndex < pIndex Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.SkipTo", "Cannot move to a position before the current position. (current: " & pIndex & " / target: " & TargetIndex & ")"
    End If
    pIndex = TargetIndex
End Sub
Private Sub IEnumerator_SkipTo(ByVal TargetIndex As Long)
    Call SkipTo(TargetIndex)
End Sub

' Attempts to remove the current element.
'
' @return Variant The removed element.
'
' @details
' WorksheetRangeBoundsEnumerator is read-only.
' This method always raises an error.
Public Function Remove() As Variant
    Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.Remove", "This collection is read-only and does not support removing elements."
End Function
Private Function IEnumerator_Remove() As Variant
    Set IEnumerator_Remove = Remove()
End Function

' Attempts to update the current element.
'
' @param ItemObject Value of the element to update.
'
' @details
' WorksheetRangeBoundsEnumerator is read-only.
' This method always raises an error.
Public Sub Update(ByVal ItemObject As Variant)
    Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.Update", "This collection is read-only and does not support updating elements."
End Sub
Private Sub IEnumerator_Update(ByVal ItemObject As Variant)
    Call Update(ItemObject)
End Sub

Private Sub pCheckEnumerationMode(ByVal TargetEnumerationMode As Long)
    Select Case TargetEnumerationMode
     Case G_RANGE_ENUM_MODE_ROWS
     Case G_RANGE_ENUM_MODE_COLUMNS
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL
     Case G_RANGE_ENUM_MODE_CELLS_VERTICAL
     Case Else
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.pCheckEnumerationMode", "Invalid enumeration mode. (" & TargetEnumerationMode & ")"
    End Select
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.pCheckInit", "The object is not initialized."
End Sub

Private Sub pCalculateLength()
    If pRangeBounds Is Nothing Then Exit Sub
    If pRangeBounds.IsEmpty Then
        pLength = 0
        Exit Sub
    End If
    Select Case pEnumerationMode
     Case G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, G_RANGE_ENUM_MODE_CELLS_VERTICAL
        pLength = pRangeBounds.Count
     Case G_RANGE_ENUM_MODE_COLUMNS
        pLength = pRangeBounds.ColumnCount
     Case Else
        pLength = pRangeBounds.RowCount
    End Select
End Sub

Private Sub pCheckEndOfList()
    If pIndex < 0 Or pLength <= pIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBoundsEnumerator.pCheckEndOfList", "There is no current element. (index: " & pIndex & " / min: 0 / max: " & (pLength - 1) & ")"
    End If
End Sub
