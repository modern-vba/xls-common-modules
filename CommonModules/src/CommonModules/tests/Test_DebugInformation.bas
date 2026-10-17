Attribute VB_Name = "Test_DebugInformation"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the DebugInformation class.
'!
' #############################################################################

Public Sub Test_FinishTask_LastNamedTask_ReturnsTrueAndRemovesLastTask(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")
    Call dbg_info.StartTaskWithLocationString("TaskB", "LocB")

    ' Act
    Dim actual_value As Boolean
    actual_value = dbg_info.FinishTask("TaskB")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals True, actual_value
    Assert.Equals "*TaskA@LocA", dbg_info.BuildCurrentMessage()
End Sub

Public Sub Test_FinishTask_MiddleNamedTask_ReturnsTrueAndRemovesThroughTarget(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")
    Call dbg_info.StartTaskWithLocationString("TaskB", "LocB")
    Call dbg_info.StartTaskWithLocationString("TaskC", "LocC")

    ' Act
    Dim actual_value As Boolean
    actual_value = dbg_info.FinishTask("TaskB")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals True, actual_value
    Assert.Equals "*TaskA@LocA", dbg_info.BuildCurrentMessage()
End Sub

Public Sub Test_FinishTask_FirstNamedTask_ReturnsTrueAndClearsStack(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")
    Call dbg_info.StartTaskWithLocationString("TaskB", "LocB")

    ' Act
    Dim actual_value As Boolean
    actual_value = dbg_info.FinishTask("TaskA")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals True, actual_value
    Assert.Equals "", dbg_info.BuildCurrentMessage()
End Sub

Public Sub Test_FinishTask_MissingNamedTask_ReturnsFalseAndKeepsStack(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")

    ' Act
    Dim actual_value As Boolean
    actual_value = dbg_info.FinishTask("TaskX")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals False, actual_value
    Assert.Equals "*TaskA@LocA", dbg_info.BuildCurrentMessage()
End Sub

Public Sub Test_FinishTask_EmptyStack_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation

    ' Act
    Dim actual_value As Boolean
    actual_value = dbg_info.FinishTask()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals False, actual_value
    Assert.Equals "", dbg_info.BuildCurrentMessage()
End Sub

Public Sub Test_BuildMessageLines_TaskStack_ReturnsBulletLines(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    dbg_info.Message = "RootMessage"
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")
    Call dbg_info.StartTaskWithLocationString("TaskB", "LocB")

    ' Act
    Dim actual_value As String
    actual_value = dbg_info.BuildMessageLines(LineSep:=vbLf)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "RootMessage" & vbLf & "*TaskA@LocA" & vbLf & "*TaskB@LocB", actual_value
End Sub

Public Sub Test_BuildCurrentMessage_LocationOnly_ReturnsBulletLocation(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim dbg_info As DebugInformation
    Set dbg_info = New DebugInformation
    Call dbg_info.StartTaskWithLocationString("TaskA", "LocA")

    ' Act
    Dim actual_value As String
    actual_value = dbg_info.BuildCurrentMessage(LocationOnly:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "*@LocA", actual_value
End Sub
