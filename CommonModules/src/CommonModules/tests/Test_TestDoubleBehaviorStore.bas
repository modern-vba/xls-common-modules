Attribute VB_Name = "Test_TestDoubleBehaviorStore"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the TestDoubleBehaviorStore class.
'!
' #############################################################################

Public Sub Test_SetReturn_RegisteredArguments_ReturnsValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetReturn("ReadCell", "value-1", "Sheet1", 1&, 1&, False)

    Dim actual_value As String
    actual_value = store_item.GetReturn("ReadCell", "Sheet1", 1&, 1&, False)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "value-1", actual_value
    Assert.IsTrue store_item.HasReturn("ReadCell", "Sheet1", 1&, 1&, False)
    Assert.IsFalse store_item.HasReturn("ReadCell", "Sheet1", 1&, 2&, False)
End Sub

Public Sub Test_SetReturn_UnregisteredArguments_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Dim actual_value As Variant
    actual_value = store_item.GetReturn("ReadCell", "Sheet1")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_SetOutput_RegisteredArguments_ReturnsByRefOutput(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetOutput("ReadCell", "Expression", "=A1", "Sheet1", 1&, 1&)

    Dim actual_value As String
    actual_value = store_item.GetOutput("ReadCell", "Expression", "Sheet1", 1&, 1&)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "=A1", actual_value
    Assert.IsTrue store_item.HasOutput("ReadCell", "Expression", "Sheet1", 1&, 1&)
    Assert.IsFalse store_item.HasOutput("ReadCell", "NumberFormat", "Sheet1", 1&, 1&)
End Sub

Public Sub Test_SetError_RegisteredArguments_RaisesConfiguredError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetError("CopyFile", vbObjectError + 321, "Test Source", "copy failed", "from.txt", "to.txt", False)
    Call store_item.RaiseIfError("CopyFile", "from.txt", "to.txt", False)

    If Not Assert.ErrorRaised(vbObjectError + 321, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Test Source", Err.Source
    Assert.Equals "copy failed", Err.Description
End Sub

Public Sub Test_RecordCallWithKeyArgs_SameKey_ReturnsLatestCallAndCount(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCallWithKeyArgs("WriteCell", Array("Sheet1", 1&, 1&), "Sheet1", 1&, 1&, "first")
    Call store_item.RecordCallWithKeyArgs("WriteCell", Array("Sheet1", 1&, 1&), "Sheet1", 1&, 1&, "second")

    Dim latest_call As TestDoubleCallRecord
    Set latest_call = store_item.GetLatestCall("WriteCell", "Sheet1", 1&, 1&)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, store_item.Count
    Assert.EqualsNumeric 2, store_item.GetCallCount("WriteCell", "Sheet1", 1&, 1&)
    Assert.Equals "second", latest_call.GetArgument(3)
    Assert.Equals "WriteCell", store_item.GetCall(0).MethodName
    Assert.Equals "WriteCell", store_item.GetCall(1).MethodName
End Sub

Public Sub Test_SetObjectKeyMode_IEquatable_MatchesEquivalentObjects(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore
    Call store_item.SetObjectKeyMode("ReadCell", G_OBJECT_KEY_MODE_I_EQUATABLE)

    Dim stored_arg As Test_ObjectSetEquatableStub
    Set stored_arg = New Test_ObjectSetEquatableStub
    stored_arg.IdentityKey = "same-key"

    Dim lookup_arg As Test_ObjectSetEquatableStub
    Set lookup_arg = New Test_ObjectSetEquatableStub
    lookup_arg.IdentityKey = "same-key"

    Call store_item.SetReturn("ReadCell", "stored", stored_arg)

    Dim actual_value As String
    actual_value = store_item.GetReturn("ReadCell", lookup_arg)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "stored", actual_value
End Sub

Public Sub Test_RecordCall_ArrayArgument_DistinguishesCollisionLikeArrays(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Dim first_arr As Variant
    first_arr = Array("A,B", "C")

    Dim second_arr As Variant
    second_arr = Array("A", "B,C")

    Call store_item.RecordCall("RemoveVBComponents", first_arr, 1&)
    Call store_item.RecordCall("RemoveVBComponents", second_arr, 1&)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, store_item.GetCallCount("RemoveVBComponents", first_arr, 1&)
    Assert.EqualsNumeric 1, store_item.GetCallCount("RemoveVBComponents", second_arr, 1&)
End Sub

Public Sub Test_RecordCall_ConfiguredObjectKeyMode_StoresModeInRecord(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore
    Call store_item.SetObjectKeyMode("ReadCell", G_OBJECT_KEY_MODE_I_EQUATABLE)

    Dim stored_arg As Test_ObjectSetEquatableStub
    Set stored_arg = New Test_ObjectSetEquatableStub
    stored_arg.IdentityKey = "same-key"

    Dim lookup_arg As Test_ObjectSetEquatableStub
    Set lookup_arg = New Test_ObjectSetEquatableStub
    lookup_arg.IdentityKey = "same-key"

    Call store_item.RecordCall("ReadCell", stored_arg)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_OBJECT_KEY_MODE_I_EQUATABLE, store_item.GetCall(0).ObjectKeyMode
    Assert.EqualsNumeric 1, store_item.GetCallCount("ReadCell", lookup_arg)
End Sub

Public Sub Test_GetCall_OutOfRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Dim record_item As TestDoubleCallRecord
    Set record_item = store_item.GetCall(0)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_DefaultObjectKeyMode_AfterKeyGeneration_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.HasReturn("ReadCell")
    store_item.DefaultObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_SetObjectKeyMode_AfterMethodKeyGeneration_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.HasReturn("ReadCell")
    Call store_item.SetObjectKeyMode("ReadCell", G_OBJECT_KEY_MODE_I_EQUATABLE)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_SetObjectKeyMode_DifferentUnlockedMethodAfterKeyGeneration_Succeeds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.HasReturn("ReadCell")
    Call store_item.SetObjectKeyMode("ReadRange", G_OBJECT_KEY_MODE_I_EQUATABLE)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SetReturn_EmptyMethodName_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetReturn("", "value")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_SetOutput_EmptyOutputName_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetOutput("ReadCell", "", "value")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_RecordCall_StoresCallArguments(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCall("WriteCell", "Sheet1", 2&, 3&)

    Dim record_item As TestDoubleCallRecord
    Set record_item = store_item.GetCall(0)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 3, record_item.ArgumentCount
    Assert.Equals "Sheet1", record_item.GetArgument(0)
    Assert.EqualsNumeric 2, record_item.GetArgument(1)
    Assert.EqualsNumeric 3, record_item.GetArgument(2)
End Sub

Public Sub Test_RecordCallWithKeyArgs_KeyArgsNotArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCallWithKeyArgs("WriteCell", "not-array", "Sheet1", 1&, 1&)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_GetCalls_KeyArgs_ReturnsMatchingCallsInOriginalOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCallWithKeyArgs("WriteCell", Array("A1"), "A1", "first")
    Call store_item.RecordCallWithKeyArgs("WriteCell", Array("B1"), "B1", "other")
    Call store_item.RecordCallWithKeyArgs("writecell", Array("A1"), "A1", "second")

    Dim calls As ObjectList
    Set calls = store_item.GetCalls("WRITECELL", "A1")

    Dim first_record As TestDoubleCallRecord
    Set first_record = calls.Item(0)

    Dim second_record As TestDoubleCallRecord
    Set second_record = calls.Item(1)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, calls.Count
    Assert.EqualsNumeric 0, first_record.CallIndex
    Assert.EqualsNumeric 2, second_record.CallIndex
    Assert.Equals "first", first_record.GetArgument(1)
    Assert.Equals "second", second_record.GetArgument(1)
End Sub

Public Sub Test_GetCallsAll_MatchingMethodName_ReturnsMatchingCallsInOriginalOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCall("WriteCell", "A1")
    Call store_item.RecordCall("ReadCell", "B1")
    Call store_item.RecordCall("writecell", "C1")

    Dim calls As ObjectList
    Set calls = store_item.GetCallsAll("WRITECELL")

    Dim first_record As TestDoubleCallRecord
    Set first_record = calls.Item(0)

    Dim second_record As TestDoubleCallRecord
    Set second_record = calls.Item(1)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, calls.Count
    Assert.EqualsNumeric 0, first_record.CallIndex
    Assert.EqualsNumeric 2, second_record.CallIndex
    Assert.Equals "A1", first_record.GetArgument(0)
    Assert.Equals "C1", second_record.GetArgument(0)
End Sub

Public Sub Test_GetCalls_NoMatchingKey_ReturnsEmptyList(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCall("ReadCell", "A1")

    Dim calls As ObjectList
    Set calls = store_item.GetCalls("ReadCell", "B1")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, calls.Count
End Sub

Public Sub Test_GetCallsAll_NoMatchingMethod_ReturnsEmptyList(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.RecordCall("ReadCell", "A1")

    Dim calls As ObjectList
    Set calls = store_item.GetCallsAll("WriteCell")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, calls.Count
End Sub

Public Sub Test_GetCallsAll_SetReturnOnly_ReturnsEmptyList(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Call store_item.SetReturn("ReadCell", "stored", "A1")

    Dim calls As ObjectList
    Set calls = store_item.GetCallsAll("ReadCell")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, calls.Count
End Sub

Public Sub Test_GetLatestCall_NoMatchingKey_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Dim record_item As TestDoubleCallRecord
    Set record_item = store_item.GetLatestCall("WriteCell", "A1")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub

Public Sub Test_GetLatestCallAll_NoMatchingMethod_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim store_item As TestDoubleBehaviorStore
    Set store_item = New TestDoubleBehaviorStore

    Dim record_item As TestDoubleCallRecord
    Set record_item = store_item.GetLatestCallAll("WriteCell")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
End Sub
