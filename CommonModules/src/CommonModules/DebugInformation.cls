VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class that holds debug information.
'!
' #############################################################################

Private pTaskStack As ObjectList
Private pLocationStack As ObjectList
Private pMessage As String

'* Debug message. This property holds an arbitrary string.
Public Property Get Message() As String
    Message = pMessage
End Property
Public Property Let Message(ByVal NewValue As String)
    pMessage = NewValue
End Property

Private Sub Class_Initialize()
    Set pTaskStack = New ObjectList
    Set pLocationStack = New ObjectList
End Sub

'* Resets the debug information.
'*
'* @details
'* Initializes the currently held task stack, location stack, and debug message.
Public Sub Reset()
    Set pTaskStack = New ObjectList
    Set pLocationStack = New ObjectList
    pMessage = ""
End Sub

'* Starts a task with Excel address information.
'*
'* @param TaskName The task name.
'* @param RangeBounds Worksheet range information (optional).
'*
'* @details
'* Starts a task based on the specified Excel location information.
Public Sub StartTask( _
        ByVal TaskName As String, _
        Optional ByVal RangeBounds As WorksheetRangeBounds = Nothing)

    If RangeBounds Is Nothing Then
        Call StartTaskWithLocationString(TaskName, WbSrv.GetThisWorkbookName())
    Else
        Call StartTaskWithLocationString(TaskName, RangeBounds.ToString())
    End If
End Sub

'* Starts a task with an arbitrary string as location information.
'*
'* @param TaskName The task name.
'* @param LocationName Arbitrary location information.
'*
'* @details
'* Starts a task with location information specified as an arbitrary string.
Public Sub StartTaskWithLocationString( _
        ByVal TaskName As String, _
        ByVal LocationName As String)

    Call pTaskStack.PushItem(TaskName)
    Call pLocationStack.PushItem(LocationName)
End Sub

'* Ends a task.
'*
'* @param TaskName The task name (when omitted, ends the last task on the stack).
'* @param RaiseError Whether to throw an exception when an error occurs (default is False).
'*
'* @return Boolean True if the task ended successfully; otherwise, False.
'*
'* @details
'* Removes the specified task name from the stack. If the task name is omitted, ends the task most recently started.
'* If the matching task does not exist, raises an error when RaiseError is True.
Public Function FinishTask(Optional ByVal TaskName As String = "", Optional ByVal RaiseError As Boolean = False) As Boolean
    If TaskName = "" Then
        If 0 < pTaskStack.Count Then
            Call pFinishTaskCore
            FinishTask = True
        Else
            If RaiseError Then
                Err.Raise vbObjectError + 1, "Class DebugInformation.FinishTask", "The task has not been started."
            Else
                FinishTask = False
            End If
        End If
        Exit Function
    End If

    Dim target_idx As Long

    target_idx = pTaskStack.GetIndexByItem(TaskName, Reverse:=True)

    If target_idx < 0 Then
        If RaiseError Then
            Err.Raise vbObjectError + 1, "Class DebugInformation.FinishTask", "The task to finish does not exist. (" & TaskName & ")"
        Else
            FinishTask = False
        End If
        Exit Function
    End If

    Dim item_idx As Long
    For item_idx = pTaskStack.Count - 1 To target_idx Step -1
        Call pFinishTaskCore
    Next item_idx
    FinishTask = True
End Function

Private Sub pFinishTaskCore()
    Call pTaskStack.PopItem
    Call pLocationStack.PopItem
End Sub

'* Builds a string that summarizes the debug information.
'*
'* @param LineSep Line separator (default is vbCrLf).
'* @param LocationOnly True to include only location information (default is False).
'* @return String A string that summarizes the debug information.
'*
'* @details
'* Returns the currently held tasks and their location information as a formatted string.
Public Function BuildMessageLines(Optional ByVal LineSep As String = vbCrLf, Optional ByVal LocationOnly As Boolean = False) As String
    Dim result_value As String

    result_value = pMessage

    Dim item_idx As Long
    For item_idx = 0 To pTaskStack.Count - 1
        result_value = Strip(result_value)
        result_value = result_value & LineSep & pBuildMessageByIndex(item_idx, LocationOnly)
    Next item_idx

    BuildMessageLines = Strip(result_value)
End Function

'* Builds a message for the most recently started task.
'*
'* @param LocationOnly True to include only location information (default is False).
'* @return String Message for the last task.
'*
'* @details
'* Returns information for the task most recently added to the current task stack as a string.
Public Function BuildCurrentMessage(Optional ByVal LocationOnly As Boolean = False) As String
    If 0 < pTaskStack.Count Then
        BuildCurrentMessage = pBuildMessageByIndex(pTaskStack.Count - 1, LocationOnly)
    Else
        BuildCurrentMessage = ""
    End If
End Function

Private Function pBuildMessageByIndex(ByVal ItemIndex As Long, ByVal LocationOnly As Boolean) As String
    Dim result_value As String

    If Not LocationOnly Then result_value = result_value & pTaskStack.Item(ItemIndex)

    Dim loc_str As String
    loc_str = pLocationStack.Item(ItemIndex)
    If loc_str <> "" Then result_value = result_value & "@" & loc_str

    If result_value <> "" Then result_value = "*" & result_value

    pBuildMessageByIndex = result_value
End Function

