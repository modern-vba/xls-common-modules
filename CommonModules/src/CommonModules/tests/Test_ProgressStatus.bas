Attribute VB_Name = "Test_ProgressStatus"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the ProgressStatus class.
'!
' #############################################################################

Public Sub Test_ClassTerminate_ObjectDisposed_StatusBarReset(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Application.StatusBar = "test progress"

    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus
    prog_stat.TaskName = "test"
    Call prog_stat.Start

    ' Act
    Set prog_stat = Nothing

    ' Assert
    Dim actual_value As Boolean
    actual_value = CBool(Application.StatusBar)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        Application.StatusBar = False
        Exit Sub
    End If
    Assert.Equals False, actual_value

    Application.StatusBar = False
End Sub

Public Sub Test_Start_DefaultRange_StartsWithoutCompleting(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Application.StatusBar = False

    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus

    ' Act
    Err.Clear
    Dim actual_started As Boolean
    actual_started = prog_stat.Start

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_processed_value As Double
    Dim actual_is_complete As Boolean
    actual_processed_value = prog_stat.ProcessedValue
    actual_is_complete = prog_stat.IsComplete

    ' Cleanup
    Set prog_stat = Nothing
    Application.StatusBar = False
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_started
    Assert.Equals 1#, actual_processed_value
    Assert.IsFalse actual_is_complete
End Sub

Public Sub Test_Start_ShortRange_CompletesAndClampsToTotal(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Application.StatusBar = False

    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus
    prog_stat.TotalValue = 0.5

    ' Act
    Err.Clear
    Dim actual_started As Boolean
    actual_started = prog_stat.Start

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_processed_value As Double
    Dim actual_is_complete As Boolean
    actual_processed_value = prog_stat.ProcessedValue
    actual_is_complete = prog_stat.IsComplete

    ' Cleanup
    Set prog_stat = Nothing
    Application.StatusBar = False
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_started
    Assert.Equals 0.5, actual_processed_value
    Assert.IsTrue actual_is_complete
End Sub

Public Sub Test_Start_AlreadyStarted_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Application.StatusBar = False

    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus
    Call prog_stat.Start

    ' Act
    Err.Clear
    Dim actual_started As Boolean
    actual_started = prog_stat.Start

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_processed_value As Double
    Dim actual_is_complete As Boolean
    actual_processed_value = prog_stat.ProcessedValue
    actual_is_complete = prog_stat.IsComplete

    ' Cleanup
    Set prog_stat = Nothing
    Application.StatusBar = False
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_started
    Assert.Equals 1#, actual_processed_value
    Assert.IsFalse actual_is_complete
End Sub

Public Sub Test_SetForLoop_StartIndexOverDefaultTotal_IsAccepted(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus

    ' Act
    Call prog_stat.SetForLoop(FinishIndex:=2000, StartIndex:=1000)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals 999#, prog_stat.StartValue
    Assert.Equals 2000#, prog_stat.TotalValue
    Assert.Equals 999#, prog_stat.ProcessedValue
End Sub

Public Sub Test_SetForLoop_StartIndexGreaterThanFinishIndex_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus

    ' Act
    Call prog_stat.SetForLoop(FinishIndex:=5, StartIndex:=10)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ProgressStatus.")
End Sub

Public Sub Test_SetForLoop_LargeStartIndexWithinLargeTotal_DoesNotOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus
    prog_stat.TotalValue = 100000#
    Err.Clear

    ' Act
    Call prog_stat.SetForLoop(FinishIndex:=100000, StartIndex:=50000)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals 49999#, prog_stat.StartValue
    Assert.Equals 100000#, prog_stat.TotalValue
    Assert.Equals 49999#, prog_stat.ProcessedValue
End Sub
