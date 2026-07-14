VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class for counting.
'!
' #############################################################################

Private pCount As Long
Private pCountStep As Long
Private pStopWhenMax As Boolean
Private pMaxCount As Long

'* Current count (default is 0).
Public Property Get Count() As Long
    Count = pCount
End Property

'* Count interval (default is 1).
Public Property Get CountStep() As Long
    CountStep = pCountStep
End Property

'* Whether to stop counting when MaxCount is reached (that is, whether HasNext returns False).
'* The default is False, so HasNext always returns True.
Public Property Get StopWhenMax() As Boolean
    StopWhenMax = pStopWhenMax
End Property
Public Property Let StopWhenMax(ByVal NewValue As Boolean)
    pStopWhenMax = NewValue
End Property

'* Maximum value at which to stop counting.
Public Property Get MaxCount() As Long
    MaxCount = pMaxCount
End Property
Public Property Let MaxCount(ByVal NewValue As Long)
    pMaxCount = NewValue
End Property

Private Sub Class_Initialize()
    pCountStep = 1
End Sub

'* Initializes the counter.
'*
'* @param InitialCount Initial value (default is 0).
'* @param CountStepNumber Count interval (default is 1).
'*
'* @details
'* Initializes the counter.
Public Sub Initialize(Optional ByVal InitialCount As Long = 0, Optional ByVal CountStepNumber As Long = 1)
    If CountStepNumber = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class Counter.Initialize", Description:="CountStepNumber must be nonzero. (" & CountStepNumber & ")"
        Exit Sub
    End If

    pCount = InitialCount
    pCountStep = CountStepNumber
End Sub

'* Checks whether the count can advance.
'*
'* @return Boolean True if the count can advance; otherwise, False.
'*
'* @details
'* Checks whether the count can advance.
Public Function HasNext() As Boolean
    If Not StopWhenMax Then
        HasNext = True
        Exit Function
    End If

    If 0 < pCountStep Then
        HasNext = CDbl(pCountStep) <= CDbl(pMaxCount) - CDbl(pCount)
    ElseIf pCountStep < 0 Then
        HasNext = CDbl(pMaxCount) - CDbl(pCount) <= CDbl(pCountStep)
    Else
        HasNext = False
    End If
End Function

'* Advances the count and returns the current count value.
'*
'* @return Long The count value after advancing.
'* @details
'* Advances the count and returns the current count value. Raises an error if HasNext is False.
Public Function GoNext() As Long
    If Not HasNext() Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class Counter.GoNext", Description:="There is no next count. (count: " & pCount & ", step: " & pCountStep & ", max: " & pMaxCount & ")"
    End If

    pCount = pCount + pCountStep
    GoNext = pCount
End Function
