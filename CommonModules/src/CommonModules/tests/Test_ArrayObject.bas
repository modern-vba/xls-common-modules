Attribute VB_Name = "Test_ArrayObject"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the ArrayObject class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_ClassInitialize_AfterCreation_KeepsZeroBasedEmptyArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject

    ' --- Act ---
    Dim actual_lbound As Long
    Dim actual_ubound As Long
    actual_lbound = target_obj.LowerBound
    actual_ubound = target_obj.UpperBound

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_lbound
    Assert.EqualsNumeric -1, actual_ubound
End Sub

Public Sub Test_ReDimArray_NamedUpperBoundArgument_SetsBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject

    ' --- Act ---
    Call target_obj.ReDimArray(LowerBound:=2, UpperBound:=4)

    Dim actual_lbound As Long
    Dim actual_ubound As Long
    actual_lbound = target_obj.LowerBound
    actual_ubound = target_obj.UpperBound

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_lbound
    Assert.EqualsNumeric 4, actual_ubound
End Sub

Public Sub Test_UpdateItem_ValueTypeElement_UpdatesValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=0)

    ' --- Act ---
    Call target_obj.Update(0, "alpha")

    Dim actual_value As Variant
    actual_value = target_obj.Item(0)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", actual_value
End Sub

Public Sub Test_UpdateItem_ObjectElement_SetsSameReference(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=0)

    Dim expected_obj As ObjectList
    Set expected_obj = New ObjectList

    ' --- Act ---
    Call target_obj.Update(0, expected_obj)

    Dim actual_obj As ObjectList
    Set actual_obj = target_obj.Item(0)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue (actual_obj Is expected_obj)
End Sub

Public Sub Test_ReDimArray_PreserveFromEmptyArrayToZeroBasedRange_SetsBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject

    ' --- Act ---
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=1, PreserveItems:=True)

    Dim actual_lbound As Long
    Dim actual_ubound As Long
    actual_lbound = target_obj.LowerBound
    actual_ubound = target_obj.UpperBound

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_lbound
    Assert.EqualsNumeric 1, actual_ubound
End Sub

Public Sub Test_ReDimArray_PreserveExpandsUpperBound_KeepsExistingValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=0)
    Call target_obj.Update(0, "alpha")

    ' --- Act ---
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=2, PreserveItems:=True)

    Dim actual_value As Variant
    actual_value = target_obj.Item(0)

    Dim actual_ubound As Long
    actual_ubound = target_obj.UpperBound

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", actual_value
    Assert.EqualsNumeric 2, actual_ubound
End Sub
