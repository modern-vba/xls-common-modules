Attribute VB_Name = "Test_TestDoubleVariantKeyBuilde"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the TestDoubleVariantKeyBuilder class.
'!
' #############################################################################

Public Sub Test_BuildKey_PrimitiveAndSpecialValues_ReturnsTypedKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    Dim actual_value As String
    actual_value = key_builder.BuildKey("A" & vbTab & "B", Null, CVErr(xlErrNA), CLng(1))

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue 0 < InStr(1, actual_value, "String(A\tB)", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_value, "Null()", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_value, "Error(2042)", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_value, "Long(1)", vbBinaryCompare)
End Sub

Public Sub Test_BuildKey_ArrayArguments_DistinguishesCommaContainingItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    Dim first_arr As Variant
    first_arr = Array("A,B", "C")

    Dim second_arr As Variant
    second_arr = Array("A", "B,C")

    Dim first_key As String
    first_key = key_builder.BuildKey(first_arr)

    Dim second_key As String
    second_key = key_builder.BuildKey(second_arr)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.NotEquals first_key, second_key
End Sub

Public Sub Test_BuildKey_IEquatableObjects_UsesIdentityKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder
    key_builder.ObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE

    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub
    first_obj.IdentityKey = "same-key"

    Dim second_obj As Test_ObjectSetEquatableStub
    Set second_obj = New Test_ObjectSetEquatableStub
    second_obj.IdentityKey = "same-key"

    Dim first_key As String
    first_key = key_builder.BuildKey(first_obj)

    Dim second_key As String
    second_key = key_builder.BuildKey(second_obj)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals first_key, second_key
End Sub

Public Sub Test_BuildKeyFromArray_NonArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    Dim actual_value As String
    actual_value = key_builder.BuildKeyFromArray(CLng(1))

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_BuildKey_DefaultObjectKeyMode_UsesReferenceKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub
    first_obj.IdentityKey = "same-key"

    Dim second_obj As Test_ObjectSetEquatableStub
    Set second_obj = New Test_ObjectSetEquatableStub
    second_obj.IdentityKey = "same-key"

    Dim first_key As String
    first_key = key_builder.BuildKey(first_obj)

    Dim second_key As String
    second_key = key_builder.BuildKey(second_obj)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.NotEquals first_key, second_key
End Sub

Public Sub Test_BuildKey_ObjectKeyModeIEquatable_FallsBackToReferenceForUnsupportedObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder
    key_builder.ObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE

    Dim first_obj As Collection
    Set first_obj = New Collection

    Dim second_obj As Collection
    Set second_obj = New Collection

    Dim first_key As String
    first_key = key_builder.BuildKey(first_obj)

    Dim second_key As String
    second_key = key_builder.BuildKey(second_obj)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.NotEquals first_key, second_key
End Sub

Public Sub Test_BuildKey_EmptyArguments_ReturnsEmptyKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    Dim actual_value As String
    actual_value = key_builder.BuildKey()

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", actual_value
End Sub

Public Sub Test_ObjectKeyMode_InvalidValue_RaisesErrorOnSet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim key_builder As TestDoubleVariantKeyBuilder
    Set key_builder = New TestDoubleVariantKeyBuilder

    key_builder.ObjectKeyMode = 9999

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleVariantKeyBuilder.")
End Sub
