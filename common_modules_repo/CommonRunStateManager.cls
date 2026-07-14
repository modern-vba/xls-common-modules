VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonRunStateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Manages shared run state used during a single GUI run.
'!
' #############################################################################

Private Sub Class_Initialize()
    Set DbgInfo = New DebugInformation
    Set ProgStat = New ProgressStatus
End Sub

'* Explicitly initializes the shared run state.
'*
'* @details
'* Assigns new objects to DbgInfo and ProgStat.
Public Sub Initialize()
    Set DbgInfo = New DebugInformation
    Set ProgStat = New ProgressStatus
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Clear
    On Error GoTo 0
End Sub

'* Explicitly releases the shared run state.
'*
'* @details
'* Sets DbgInfo and ProgStat back to Nothing so state does not carry over to the next GUI run.
Public Sub Clear()
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub
