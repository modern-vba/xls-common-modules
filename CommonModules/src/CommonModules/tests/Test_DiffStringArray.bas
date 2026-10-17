Attribute VB_Name = "Test_DiffStringArray"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the ObjectList class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

' =============================================================================
'  Template
' =============================================================================
'Sub Test_MethodUnderTest_ArgumentConditions_ExpectedResult(ByVal Assert As UnitTestAssert)
'    On Error Resume Next
'
'    '--- Arrange ---
'    Dim local_xxxx As XXXX
'    Set local_xxxx = New XXXX
'
'    '--- Act ---
'    Dim actual_val as Variant
'    actual_val = ...
'
'    '--- Assert ---
'    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
'
'    Assert.Equals 99, actual_val
'
'    Assert.ErrorNotRaised 0, Err.Number, err.source, Err.Description
'End Sub
' =============================================================================

Public Sub Test_DiffStringArray_WithoutReplacement_ReturnsCorrectDiffResult(ByVal Assert As UnitTestAssert)
    'On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String

    ReDim old_arr(0 To 5)
    old_arr(0) = "aaa"
    old_arr(1) = "bbb"
    old_arr(2) = "ccc"
    old_arr(3) = "aaa"
    old_arr(4) = "bbb"
    old_arr(5) = "ccc"

    ReDim new_arr(0 To 5)
    new_arr(0) = "bbb"
    new_arr(1) = "ccc"
    new_arr(2) = "ddd"
    new_arr(3) = "bbb"
    new_arr(4) = "ccc"
    new_arr(5) = "ddd"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr, EnableReplaceType:=False)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Assert.EqualsNumeric 7, UBound(old_arr)
    If UBound(old_arr) <> 7 Then Exit Sub

    Assert.EqualsNumeric 7, UBound(new_arr)
    If UBound(new_arr) <> 7 Then Exit Sub

    Assert.EqualsNumeric 7, UBound(type_arr)
    If UBound(type_arr) <> 7 Then Exit Sub

    Assert.Equals "aaa", old_arr(0)
    Assert.Equals "", new_arr(0)
    Assert.Equals "DEL", type_arr(0)

    Assert.Equals "bbb", old_arr(1)
    Assert.Equals "bbb", new_arr(1)
    Assert.Equals "", type_arr(1)

    Assert.Equals "ccc", old_arr(2)
    Assert.Equals "ccc", new_arr(2)
    Assert.Equals "", type_arr(2)

    Assert.Equals "aaa", old_arr(3)
    Assert.Equals "", new_arr(3)
    Assert.Equals "DEL", type_arr(3)

    Assert.Equals "", old_arr(4)
    Assert.Equals "ddd", new_arr(4)
    Assert.Equals "ADD", type_arr(4)

    Assert.Equals "bbb", old_arr(5)
    Assert.Equals "bbb", new_arr(5)
    Assert.Equals "", type_arr(5)

    Assert.Equals "ccc", old_arr(6)
    Assert.Equals "ccc", new_arr(6)
    Assert.Equals "", type_arr(6)

    Assert.Equals "", old_arr(7)
    Assert.Equals "ddd", new_arr(7)
    Assert.Equals "ADD", type_arr(7)

    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_DiffStringArray_WithReplacement_ReturnsCorrectDiffResult(ByVal Assert As UnitTestAssert)
    'On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String

    ReDim old_arr(0 To 5)
    old_arr(0) = "aaa"
    old_arr(1) = "bbb"
    old_arr(2) = "ccc"
    old_arr(3) = "aaa"
    old_arr(4) = "bbb"
    old_arr(5) = "ccc"

    ReDim new_arr(0 To 5)
    new_arr(0) = "bbb"
    new_arr(1) = "ccc"
    new_arr(2) = "ddd"
    new_arr(3) = "bbb"
    new_arr(4) = "ccc"
    new_arr(5) = "ddd"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Assert.EqualsNumeric 6, UBound(old_arr)
    If UBound(old_arr) <> 6 Then Exit Sub

    Assert.EqualsNumeric 6, UBound(new_arr)
    If UBound(new_arr) <> 6 Then Exit Sub

    Assert.EqualsNumeric 6, UBound(type_arr)
    If UBound(type_arr) <> 6 Then Exit Sub

    Assert.Equals "aaa", old_arr(0)
    Assert.Equals "", new_arr(0)
    Assert.Equals "DEL", type_arr(0)

    Assert.Equals "bbb", old_arr(1)
    Assert.Equals "bbb", new_arr(1)
    Assert.Equals "", type_arr(1)

    Assert.Equals "ccc", old_arr(2)
    Assert.Equals "ccc", new_arr(2)
    Assert.Equals "", type_arr(2)

    Assert.Equals "aaa", old_arr(3)
    Assert.Equals "ddd", new_arr(3)
    Assert.Equals "MOD", type_arr(3)

    Assert.Equals "bbb", old_arr(4)
    Assert.Equals "bbb", new_arr(4)
    Assert.Equals "", type_arr(4)

    Assert.Equals "ccc", old_arr(5)
    Assert.Equals "ccc", new_arr(5)
    Assert.Equals "", type_arr(5)

    Assert.Equals "", old_arr(6)
    Assert.Equals "ddd", new_arr(6)
    Assert.Equals "ADD", type_arr(6)

    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_DiffStringArray_EmptyOldAndNew_ReturnsEmptyArrays(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(old_arr)
    Assert.IsTrue IsEmptyArray(new_arr)
    Assert.IsTrue IsEmptyArray(type_arr)
End Sub

Public Sub Test_DiffStringArray_EmptyOldAndNonEmptyNew_ReturnsAllAddWithZeroBase(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String
    ReDim new_arr(5 To 6)
    new_arr(5) = "first"
    new_arr(6) = "second"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(old_arr)
    Assert.EqualsNumeric 1, UBound(old_arr)
    Assert.EqualsNumeric 0, LBound(new_arr)
    Assert.EqualsNumeric 1, UBound(new_arr)
    Assert.EqualsNumeric 0, LBound(type_arr)
    Assert.EqualsNumeric 1, UBound(type_arr)

    Assert.Equals "", old_arr(0)
    Assert.Equals "first", new_arr(0)
    Assert.Equals "ADD", type_arr(0)
    Assert.Equals "", old_arr(1)
    Assert.Equals "second", new_arr(1)
    Assert.Equals "ADD", type_arr(1)
End Sub

Public Sub Test_DiffStringArray_NonEmptyOldAndEmptyNew_ReturnsAllDelWithZeroBase(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String
    ReDim old_arr(3 To 4)
    old_arr(3) = "first"
    old_arr(4) = "second"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(old_arr)
    Assert.EqualsNumeric 1, UBound(old_arr)
    Assert.EqualsNumeric 0, LBound(new_arr)
    Assert.EqualsNumeric 1, UBound(new_arr)
    Assert.EqualsNumeric 0, LBound(type_arr)
    Assert.EqualsNumeric 1, UBound(type_arr)

    Assert.Equals "first", old_arr(0)
    Assert.Equals "", new_arr(0)
    Assert.Equals "DEL", type_arr(0)
    Assert.Equals "second", old_arr(1)
    Assert.Equals "", new_arr(1)
    Assert.Equals "DEL", type_arr(1)
End Sub

Public Sub Test_DiffStringArray_NonZeroLowerBounds_ReturnsZeroBaseArrays(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String
    ReDim old_arr(3 To 4)
    old_arr(3) = "same"
    old_arr(4) = "old"

    ReDim new_arr(7 To 8)
    new_arr(7) = "same"
    new_arr(8) = "new"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)

    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(old_arr)
    Assert.EqualsNumeric 1, UBound(old_arr)
    Assert.EqualsNumeric 0, LBound(new_arr)
    Assert.EqualsNumeric 1, UBound(new_arr)
    Assert.EqualsNumeric 0, LBound(type_arr)
    Assert.EqualsNumeric 1, UBound(type_arr)

    Assert.Equals "same", old_arr(0)
    Assert.Equals "same", new_arr(0)
    Assert.Equals "", type_arr(0)
    Assert.Equals "old", old_arr(1)
    Assert.Equals "new", new_arr(1)
    Assert.Equals "MOD", type_arr(1)
End Sub

