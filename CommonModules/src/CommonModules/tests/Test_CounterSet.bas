Attribute VB_Name = "Test_CounterSet"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the CounterSet class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_GetCounter_MissingName_RaisesWithoutCreatingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim counter_set As CounterSet
    Set counter_set = New CounterSet

    ' Act
    Dim missing_counter As Counter
    Set missing_counter = counter_set.GetCounter("missing")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class CounterSet.")

    Err.Clear
    Dim added_counter As Counter
    Set added_counter = New Counter
    Call counter_set.AddCounter("missing", added_counter)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim stored_counter As Counter
    Set stored_counter = counter_set.GetCounter("missing")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim is_same As Boolean
    is_same = stored_counter Is added_counter
    Assert.IsTrue is_same
End Sub

Public Sub Test_RemoveCounter_MissingName_RaisesWithoutCreatingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim counter_set As CounterSet
    Set counter_set = New CounterSet

    ' Act
    Dim removed_counter As Counter
    Set removed_counter = counter_set.RemoveCounter("missing")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class CounterSet.")

    Err.Clear
    Dim added_counter As Counter
    Set added_counter = New Counter
    Call counter_set.AddCounter("missing", added_counter)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim stored_counter As Counter
    Set stored_counter = counter_set.GetCounter("missing")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim is_same As Boolean
    is_same = stored_counter Is added_counter
    Assert.IsTrue is_same
End Sub
