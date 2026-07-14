VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class that manages progress status. It visually displays and updates task progress.
'!
' #############################################################################

Private Const C_MSG_COMPLETE_BLOCK As String = "Å°"
Private Const C_MSG_UNCOMPLETE_BLOCK As String = "Å†"

Private pTaskName As String
Private pBlockCount As Long
Private pStartValue As Double
Private pTotalValue As Double
Private pProcessedValue As Double
Private pIsComplete As Boolean
Private pIntervalPercent As Integer
Private pNextUpdate As Integer

'* Task name.
Public Property Get TaskName() As String
    TaskName = pTaskName
End Property
Public Property Let TaskName(ByVal NewValue As String)
    pTaskName = NewValue
End Property

'* Number of blocks. (Default is 10.)
'* Number of progress bar segments.
Public Property Get BlockCount() As Long
    BlockCount = pBlockCount
End Property
Public Property Let BlockCount(ByVal NewValue As Long)
    If NewValue < 1 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus.BlockCount", Description:="BlockCount must be a positive integer. (" & NewValue & ")"
        Exit Property
    End If
    pBlockCount = NewValue
End Property

'* Start value. (Default is 0.)
'* Changing this property also changes the progress value.
Public Property Get StartValue() As Double
    StartValue = pStartValue
End Property
Public Property Let StartValue(ByVal NewValue As Double)
    If TotalValue <= NewValue Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus.StartValue", Description:="StartValue must be less than TotalValue. (new value: " & NewValue & " / TotalValue: " & pTotalValue & ")"
        Exit Property
    End If
    Dim value_diff As Double
    value_diff = pStartValue - NewValue
    pStartValue = NewValue
    pProcessedValue = pProcessedValue - value_diff
    pNextUpdate = 0
End Property

'* Finish value. (Default is 100.)
Public Property Get TotalValue() As Double
    TotalValue = pTotalValue
End Property
Public Property Let TotalValue(ByVal NewValue As Double)
    If NewValue <= StartValue Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus.TotalValue", Description:="TotalValue must be greater than StartValue. (new value: " & NewValue & " / StartValue: " & pStartValue & ")"
        Exit Property
    End If
    pTotalValue = NewValue
End Property

'* Current progress value. (Default is 0.)
Public Property Get ProcessedValue() As Double
    ProcessedValue = pProcessedValue
End Property
Public Property Let ProcessedValue(ByVal NewValue As Double)
    pProcessedValue = NewValue
    Call pUpdateStatus
End Property

'* Whether the task is completed.
Public Property Get IsComplete() As Boolean
    IsComplete = pIsComplete
End Property

Private Sub Class_Initialize()
    Call pInitializeStartAndTotal
    pBlockCount = 10
    pProcessedValue = pStartValue
    pIsComplete = False
    pIntervalPercent = 3
    pNextUpdate = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'* Configures settings for processing a loop over the specified range.
'*
'* @param FinishIndex Loop finish index.
'* @param Optional StartIndex Loop start index. Default is 0.
'*
'* @details
'* Sets StartIndex and FinishIndex, and updates progress by setting the index to ProcessedValue, to manage progress corresponding to loop processing.
Public Sub SetForLoop(ByVal FinishIndex As Long, Optional ByVal StartIndex As Long = 0)
    Dim start_value As Double
    start_value = CDbl(StartIndex) - 1

    Dim total_value As Double
    total_value = CDbl(FinishIndex)

    If total_value <= start_value Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ProgressStatus.SetForLoop", Description:="FinishIndex must be greater than StartIndex. (FinishIndex: " & FinishIndex & " / StartIndex: " & StartIndex & ")"
        Exit Sub
    End If

    pStartValue = start_value
    pTotalValue = total_value
    pProcessedValue = pStartValue
    pIsComplete = False
    pNextUpdate = 0
End Sub

'* Starts progress monitoring.
'*
'* @return Boolean value indicating whether progress monitoring started successfully.
'*
'* @details
'* Sets the initial progress value and starts displaying progress status.
Public Function Start() As Boolean
    If pProcessedValue < pStartValue + 1 Then
        pProcessedValue = pStartValue + 1
        If pTotalValue < pProcessedValue Then pProcessedValue = pTotalValue
        Call pUpdateStatus
        Start = True
    Else
        Start = False
    End If
End Function

'* Marks the task as completed.
'*
'* @details
'* Sets the progress value to the maximum value and marks the task completed.
Public Sub MarkAsComplete()
    pProcessedValue = pTotalValue
    Call pUpdateStatus
End Sub

'* Resets progress status.
'*
'* @param Optional ResetTaskName Whether to also reset the task name. Default is True.
'* @param Optional ResetStartAndTotalValue Whether to also reset the start and finish values. Default is True.
'*
'* @details
'* Returns progress status to its initial state and also resets the task name and values as needed.
Public Sub Reset(Optional ByVal ResetTaskName As Boolean = True, Optional ByVal ResetStartAndTotalValue As Boolean = True)
    If ResetTaskName Then pTaskName = ""
    If ResetStartAndTotalValue Then Call pInitializeStartAndTotal
    pProcessedValue = pStartValue
    pIsComplete = False
    pNextUpdate = 0
    Application.StatusBar = False
End Sub

'* Manually displays the current progress information.
'*
'* @details
'* Forcibly updates progress status and displays it in the status bar.
Public Sub DisplayProgress()
    Call pDisplayProgress(Force:=True)
End Sub

Private Sub pDisplayProgress(Optional ByVal Force As Boolean = False)
    Dim prog_rate As Double
    Dim prog_percent As Integer

    prog_rate = (pProcessedValue - pStartValue) / (pTotalValue - pStartValue)

    If prog_rate < 0 Then
        prog_rate = 0
    ElseIf 1 < prog_rate Then
        prog_rate = 1
    End If

    prog_percent = Int((prog_rate + 0.0001) * 100)

    If Force Or pNextUpdate <= prog_percent Then
        pNextUpdate = pGetNextUpdate(prog_percent)
        Application.StatusBar = pGetStatusString(prog_rate, prog_percent)
        DoEvents
    End If
End Sub

Private Sub pInitializeStartAndTotal()
    pStartValue = 0
    pTotalValue = 100
End Sub

Private Sub pUpdateStatus()
    If pTotalValue <= pProcessedValue Then
        pIsComplete = True
    Else
        pIsComplete = False
    End If


    If pIsComplete Then
        Application.StatusBar = False
        Exit Sub
    End If

    Call pDisplayProgress
End Sub

Private Function pGetNextUpdate(ByVal ProgressPercent As Integer) As Integer
    Dim result_value As Long

    result_value = (ProgressPercent \ pIntervalPercent + 1) * pIntervalPercent

    If 100 < result_value Then
        result_value = 100
    End If

    pGetNextUpdate = result_value
End Function

Private Function pGetStatusString(ByVal ProgressRate As Double, ByVal ProgressPercent As Integer) As String
    Dim result_value As String
    Dim comp_blocks As Long
    Dim item_idx As Long

    result_value = pTaskName & "[" & ProgressPercent & "%]"

    comp_blocks = Int(pBlockCount * (ProgressRate + 0.0001))
    For item_idx = 1 To comp_blocks
        result_value = result_value & C_MSG_COMPLETE_BLOCK
    Next item_idx

    comp_blocks = pBlockCount - comp_blocks
    For item_idx = 1 To comp_blocks
        result_value = result_value & C_MSG_UNCOMPLETE_BLOCK
    Next item_idx

    pGetStatusString = result_value
End Function
