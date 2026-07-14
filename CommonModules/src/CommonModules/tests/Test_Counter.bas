Attribute VB_Name = "Test_Counter"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the Counter class.
'!
' #############################################################################

Private Const C_LONG_MAX As Long = 2147483647#
Private Const C_LONG_MIN As Long = -2147483648#

Public Sub Test_Initialize_CountStepNumberZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_counter As Counter
    Set target_counter = New Counter

    ' Act
    Call target_counter.Initialize(InitialCount:=1, CountStepNumber:=0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Counter.")
End Sub

Public Sub Test_HasNext_PositiveStepAtLongMax_ReturnsFalseWithoutOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_counter As Counter
    Set target_counter = New Counter
    Call target_counter.Initialize(InitialCount:=C_LONG_MAX, CountStepNumber:=1)
    target_counter.StopWhenMax = True
    target_counter.MaxCount = C_LONG_MAX
    Err.Clear

    ' Act
    Dim actual_value As Boolean
    actual_value = target_counter.HasNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals False, actual_value
End Sub

Public Sub Test_HasNext_NegativeStepAtLongMin_ReturnsFalseWithoutOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_counter As Counter
    Set target_counter = New Counter
    Call target_counter.Initialize(InitialCount:=C_LONG_MIN, CountStepNumber:=-1)
    target_counter.StopWhenMax = True
    target_counter.MaxCount = C_LONG_MIN
    Err.Clear

    ' Act
    Dim actual_value As Boolean
    actual_value = target_counter.HasNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals False, actual_value
End Sub

Public Sub Test_HasNext_PositiveStepToLongMax_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_counter As Counter
    Set target_counter = New Counter
    Call target_counter.Initialize(InitialCount:=C_LONG_MAX - 1, CountStepNumber:=1)
    target_counter.StopWhenMax = True
    target_counter.MaxCount = C_LONG_MAX
    Err.Clear

    ' Act
    Dim actual_value As Boolean
    actual_value = target_counter.HasNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals True, actual_value
    Assert.EqualsNumeric C_LONG_MAX, target_counter.GoNext()
End Sub

Public Sub Test_HasNext_NegativeStepToLongMin_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_counter As Counter
    Set target_counter = New Counter
    Call target_counter.Initialize(InitialCount:=C_LONG_MIN + 1, CountStepNumber:=-1)
    target_counter.StopWhenMax = True
    target_counter.MaxCount = C_LONG_MIN
    Err.Clear

    ' Act
    Dim actual_value As Boolean
    actual_value = target_counter.HasNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals True, actual_value
    Assert.EqualsNumeric C_LONG_MIN, target_counter.GoNext()
End Sub
