Attribute VB_Name = "Test_UnitTestAssert"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the UnitTestAssert class.
'!
' #############################################################################

Public Sub Test_Equals_NullValues_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(Null, Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_EmptyValues_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_value As Variant
    Dim actual_value As Variant

    ' Act
    Call target_assert.Equals(expected_value, actual_value)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrSameCode_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrDiv0), CVErr(xlErrDiv0))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrDifferentCode_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrDiv0), CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrAndNull_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrValue), Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_Equals_NullAndEmpty_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim empty_value As Variant

    ' Act
    Call target_assert.Equals(Null, empty_value)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_EqualsNumeric_NullValues_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsNumeric(Null, Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_Equals_ArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("value")

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.Equals(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array arguments are not supported") > 0
End Sub

Public Sub Test_NotEquals_ArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("expected")

    Dim actual_arr As Variant
    actual_arr = Array("actual")

    ' Act
    Call target_assert.NotEquals(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array arguments are not supported") > 0
End Sub

Public Sub Test_EqualsNumeric_NonNumericString_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsNumeric("abc", 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Numeric assertions require") > 0
End Sub

Public Sub Test_NotEqualsNumeric_NonNumericString_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.NotEqualsNumeric("abc", 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Numeric assertions require") > 0
End Sub

Public Sub Test_EqualsArray_EmptyVariantArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array()

    Dim actual_arr As Variant
    actual_arr = Array()

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_UninitializedVariantArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr() As Variant
    Dim actual_arr() As Variant

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_OneEmptyOneNonEmpty_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array()

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_NonArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.EqualsArray("value", actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_DifferentArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("expected")

    Dim actual_arr As Variant
    actual_arr = Array("actual")

    ' Act
    Call target_assert.NotEqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_NotEqualsArray_NonArrayExpected_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.NotEqualsArray("value", actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_NonArrayActual_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("value")

    ' Act
    Call target_assert.NotEqualsArray(expected_arr, "value")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_BothNonArray_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.NotEqualsArray("expected", "actual")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_EqualsArray_WithSpecialPrimitiveElements_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim empty_value As Variant

    Dim expected_arr As Variant
    expected_arr = Array(Null, CVErr(xlErrNA), empty_value)

    Dim actual_arr As Variant
    actual_arr = Array(Null, CVErr(xlErrNA), empty_value)

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_AssertionCount_NewAssert_ReturnsZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Dim actual_count As Long
    actual_count = target_assert.AssertionCount

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, actual_count
End Sub

Public Sub Test_AssertionCount_AfterSuccessfulAssertion_ReturnsOne(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    target_assert.Equals "same", "same"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_AssertionCount_AfterFailedAssertion_ReturnsOne(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    target_assert.Equals "expected", "actual"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_Equals_WithoutCaseName_FailedMessageKeepsExistingFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals("expected", "actual")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: expected<String> | actual: actual<String>", target_assert.ResultMessage
End Sub

Public Sub Test_Equals_WithCaseName_FailedMessageIncludesCaseName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals("expected", "actual", CaseName:="sample case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: expected<String> | actual: actual<String> | case: sample case", target_assert.ResultMessage
End Sub

Public Sub Test_EqualsArray_WithCaseNameAndAdditionalInformation_FailedMessageOrdersAdditionalInformationBeforeCaseName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("expected")

    Dim actual_arr As Variant
    actual_arr = Array("actual")

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr, CaseName:="array case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: expected<String> | actual: actual<String> | @0 | case: array case", target_assert.ResultMessage
End Sub

Public Sub Test_ErrorRaised_WithCaseName_FailedMessageIncludesCaseName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.ErrorRaised(5, 0, "", "", CaseName:="err case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: The error [0x5] is raised. | actual: Any error is not raised. | case: err case", target_assert.ResultMessage
End Sub

Public Sub Test_ErrorRaised_Default_DoesNotClearErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Err.Raise 9, "test source", "test description"

    ' Act
    Dim was_raised As Boolean
    was_raised = target_assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description)

    ' Assert
    Assert.IsTrue was_raised
    Assert.EqualsNumeric 9, Err.Number
    Err.Clear
End Sub

Public Sub Test_ErrorNotRaised_ClearAfterAssertTrue_ClearsErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Err.Raise 9, "test source", "test description"

    ' Act
    Call target_assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description, CaseName:="clear case", ClearAfterAssert:=True)

    ' Assert
    Assert.EqualsNumeric 0, Err.Number
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: Any error is not raised. | actual: The error [0x9]@<test source> ""test description"" is raised. | case: clear case", target_assert.ResultMessage
End Sub

Public Sub Test_EqualsObject_ReferenceModeDifferentInstances_Fails(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "same-id"

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = New Test_ObjectSetEquatableStub
    actual_obj.IdentityKey = "same-id"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, CaseName:="reference case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.IsTrue InStr(target_assert.ResultMessage, "Object@Test_ObjectSetEquatableStub(") > 0
    Assert.IsTrue InStr(target_assert.ResultMessage, "case: reference case") > 0
End Sub

Public Sub Test_EqualsObject_IEquatableSameIdentity_Passes(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "same-id"

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = New Test_ObjectSetEquatableStub
    actual_obj.IdentityKey = "same-id"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_EqualsObject_IEquatableDifferentIdentity_FailsWithIdentityKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "expected-id"

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = New Test_ObjectSetEquatableStub
    actual_obj.IdentityKey = "actual-id"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_I_EQUATABLE, CaseName:="eq case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: IEquatable@Test_ObjectSetEquatableStub(expected-id) | actual: IEquatable@Test_ObjectSetEquatableStub(actual-id) | case: eq case", target_assert.ResultMessage
End Sub

Public Sub Test_EqualsObject_DuplicateCheckableSameKey_Passes(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetDupCheckStub
    Set expected_obj = New Test_ObjectSetDupCheckStub
    expected_obj.DuplicateKey = "same-key"

    Dim actual_obj As Test_ObjectSetDupCheckStub
    Set actual_obj = New Test_ObjectSetDupCheckStub
    actual_obj.DuplicateKey = "same-key"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_EqualsObject_DuplicateCheckableDifferentKey_FailsWithDuplicateKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetDupCheckStub
    Set expected_obj = New Test_ObjectSetDupCheckStub
    expected_obj.DuplicateKey = "expected-key"

    Dim actual_obj As Test_ObjectSetDupCheckStub
    Set actual_obj = New Test_ObjectSetDupCheckStub
    actual_obj.DuplicateKey = "actual-key"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE, CaseName:="dup case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: IDuplicateCheckable@Test_ObjectSetDupCheckStub(expected-key) | actual: IDuplicateCheckable@Test_ObjectSetDupCheckStub(actual-key) | case: dup case", target_assert.ResultMessage
End Sub

Public Sub Test_NotEqualsObject_IEquatableSameIdentity_Fails(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetEquatableStub
    Set expected_obj = New Test_ObjectSetEquatableStub
    expected_obj.IdentityKey = "same-id"

    Dim actual_obj As Test_ObjectSetEquatableStub
    Set actual_obj = New Test_ObjectSetEquatableStub
    actual_obj.IdentityKey = "same-id"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.NotEqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_I_EQUATABLE, CaseName:="not eq case")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.Equals "[0] expected: IEquatable@Test_ObjectSetEquatableStub(same-id) | actual: IEquatable@Test_ObjectSetEquatableStub(same-id) | case: not eq case", target_assert.ResultMessage
End Sub

Public Sub Test_EqualsObject_IEquatableModeNonEquatable_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_obj As Test_ObjectSetDupCheckStub
    Set expected_obj = New Test_ObjectSetDupCheckStub
    expected_obj.DuplicateKey = "same-key"

    Dim actual_obj As Test_ObjectSetDupCheckStub
    Set actual_obj = New Test_ObjectSetDupCheckStub
    actual_obj.DuplicateKey = "same-key"

    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsObject(expected_obj, actual_obj, G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class UnitTestAssert.")
    Assert.IsTrue InStr(Err.Description, "IEquatable") > 0
    Err.Clear
End Sub
