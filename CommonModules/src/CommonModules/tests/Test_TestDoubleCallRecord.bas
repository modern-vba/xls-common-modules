Attribute VB_Name = "Test_TestDoubleCallRecord"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the TestDoubleCallRecord class.
'!
' #############################################################################

Public Sub Test_StoreArguments_MultipleArguments_ReturnsCountAndArguments(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord

    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "object-id"

    Dim args_arr(0 To 2) As Variant
    args_arr(0) = "Sheet1"
    args_arr(1) = 2&
    Set args_arr(2) = expected_obj

    Call record_item.StoreArguments(args_arr)

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = record_item.GetArgument(2)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 3, record_item.ArgumentCount
    Assert.Equals "Sheet1", record_item.GetArgument(0)
    Assert.EqualsNumeric 2, record_item.GetArgument(1)
    Assert.IsTrue actual_obj Is expected_obj
End Sub

Public Sub Test_ArgumentCount_NoStoredArguments_ReturnsZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, record_item.ArgumentCount
End Sub

Public Sub Test_StoreArguments_ObjectArgument_KeepsReference(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord

    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "before"

    Dim args_arr(0 To 0) As Variant
    Set args_arr(0) = expected_obj

    Call record_item.StoreArguments(args_arr)
    expected_obj.IdentityKey = "after"

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = record_item.GetArgument(0)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue actual_obj Is expected_obj
    Assert.Equals "after", actual_obj.IdentityKey
End Sub

Public Sub Test_StoreArguments_NotArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord

    Call record_item.StoreArguments("not-array")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleCallRecord.")
    Err.Clear
End Sub

Public Sub Test_GetArgument_OutOfRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim record_item As TestDoubleCallRecord
    Set record_item = New TestDoubleCallRecord

    Dim actual_value As Variant
    actual_value = record_item.GetArgument(0)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleCallRecord.")
    Err.Clear
End Sub
