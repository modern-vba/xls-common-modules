Attribute VB_Name = "Test_ObjectDictionary"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the ObjectDictionary class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_NewObjectDictionary_AfterCreation_CountIsZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange & Act ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_count
End Sub

Public Sub Test_ObjectDictionary_Add_KeyAndValue_CanRetrieveWithItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Call target_dic.Add("alpha", "value1")

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("alpha")

    Dim actual_missing_exists As Boolean
    actual_missing_exists = target_dic.Exists("missing")

    Dim actual_value As String
    actual_value = target_dic.Item("alpha")

    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_count
    Assert.IsTrue actual_exists
    Assert.IsFalse actual_missing_exists
    Assert.Equals "value1", actual_value
End Sub

Public Sub Test_ObjectDictionary_Item_MissingKeyRaisesErrorAndDoesNotChangeState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")

    ' --- Act ---
    Err.Clear
    Dim actual_value As Variant
    actual_value = target_dic.Item("missing")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("missing")

    Dim actual_existing_value As String
    actual_existing_value = target_dic.Item("alpha")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 1, actual_count
    Assert.IsFalse actual_exists
    Assert.Equals "value1", actual_existing_value
End Sub

Public Sub Test_ObjectDictionary_Add_ExistingKeyRaisesErrorAndKeepsExistingElement(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")

    ' --- Act ---
    Err.Clear
    Call target_dic.Add("alpha", "value2")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_value As String
    actual_value = target_dic.Item("alpha")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 1, actual_count
    Assert.Equals "value1", actual_value
End Sub

Public Sub Test_ObjectDictionary_ExplicitElementTypeContractAllowsSetObjectAndRejectsTypeViolation(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ILeafCondition")

    Dim item_obj As Test_ElementTypeProviderStub
    Set item_obj = New Test_ElementTypeProviderStub
    item_obj.ElementTypeKey = "ILeafCondition"

    ' --- Act ---
    Call target_dic.Add("first", item_obj)

    Dim actual_obj As Test_ElementTypeProviderStub
    Set actual_obj = target_dic.Item("first")

    Err.Clear
    Call target_dic.Add("bad", "primitive")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_bad_exists As Boolean
    actual_bad_exists = target_dic.Exists("bad")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_obj Is item_obj
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 1, actual_count
    Assert.IsFalse actual_bad_exists
End Sub

Public Sub Test_ObjectDictionary_InfersTypeContractFromFirstValueAndRejectsViolation(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Call target_dic.Add("one", CLng(1))

    Err.Clear
    Call target_dic.Add("bad", "string")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_bad_exists As Boolean
    actual_bad_exists = target_dic.Exists("bad")

    Dim actual_value As Long
    actual_value = target_dic.Item("one")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 1, actual_count
    Assert.IsFalse actual_bad_exists
    Assert.EqualsNumeric 1, actual_value
End Sub

Public Sub Test_ObjectDictionary_RequireComparableRejectsUnsupportedObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ILeafCondition", RequireComparable:=True)

    Dim item_obj As Test_ElementTypeProviderStub
    Set item_obj = New Test_ElementTypeProviderStub
    item_obj.ElementTypeKey = "ILeafCondition"

    ' --- Act ---
    Call target_dic.Add("first", item_obj)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("first")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 0, actual_count
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_ObjectDictionary_ObjectKeyCanBeAddedAndRetrieved(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    Dim key_obj As Test_ElementTypeProviderStub
    Set key_obj = New Test_ElementTypeProviderStub
    key_obj.ElementTypeKey = "KeyType"

    ' --- Act ---
    Call target_dic.Add(key_obj, "value1")

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists(key_obj)

    Dim actual_value As String
    actual_value = target_dic.Item(key_obj)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
    Assert.Equals "value1", actual_value
End Sub

Public Sub Test_ObjectDictionary_CompareMode_DefaultIsBinaryAndCanChangeBeforeAdd(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Dim actual_default_mode As Long
    actual_default_mode = target_dic.CompareMode

    target_dic.CompareMode = vbTextCompare

    Dim actual_changed_mode As Long
    actual_changed_mode = target_dic.CompareMode

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric vbBinaryCompare, actual_default_mode
    Assert.EqualsNumeric vbTextCompare, actual_changed_mode
End Sub

Public Sub Test_ObjectDictionary_CompareMode_CannotChangeAfterAdd(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")

    ' --- Act ---
    Err.Clear
    target_dic.CompareMode = vbTextCompare

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_mode As Long
    actual_mode = target_dic.CompareMode

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectDictionary.")
    Assert.EqualsNumeric vbBinaryCompare, actual_mode
End Sub

Public Sub Test_ObjectDictionary_CompareMode_BinaryTreatsStringKeyCaseDifferencesAsDistinctKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Call target_dic.Add("Key", "upper")
    Call target_dic.Add("key", "lower")

    Dim actual_upper As String
    actual_upper = target_dic.Item("Key")

    Dim actual_lower As String
    actual_lower = target_dic.Item("key")

    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_count
    Assert.Equals "upper", actual_upper
    Assert.Equals "lower", actual_lower
End Sub

Public Sub Test_ObjectDictionary_CompareMode_TextTreatsStringKeyCaseDifferencesAsSameKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    target_dic.CompareMode = vbTextCompare

    ' --- Act ---
    Call target_dic.Add("Key", "upper")

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("key")

    Dim actual_value As String
    actual_value = target_dic.Item("key")

    Err.Clear
    Call target_dic.Add("key", "lower")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
    Assert.Equals "upper", actual_value
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.EqualsNumeric 1, actual_count
End Sub

Public Sub Test_ObjectDictionary_InitializeFromOther_InheritsValueContractAndCompareMode(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary("ObjectValue")
    source_dic.CompareMode = vbTextCompare

    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    Dim item_obj As Test_ElementTypeProviderStub
    Set item_obj = New Test_ElementTypeProviderStub
    item_obj.ElementTypeKey = "ObjectValue"

    ' --- Act ---
    Call target_dic.InitializeFromOther(source_dic)
    Call target_dic.Add("KEY", item_obj)

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("key")

    Dim actual_obj As Test_ElementTypeProviderStub
    Set actual_obj = target_dic.Item("key")

    Err.Clear
    Call target_dic.Add("bad", "primitive")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_bad_exists As Boolean
    actual_bad_exists = target_dic.Exists("bad")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_count
    Assert.IsTrue actual_exists
    Assert.IsTrue actual_obj Is item_obj
    Assert.IsFalse actual_bad_exists
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectList.")
End Sub

Public Sub Test_ObjectDictionary_InitializeFromOther_InitializedDictionaryRaisesErrorAndKeepsState(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary("ObjectValue")
    source_dic.CompareMode = vbTextCompare

    Dim added_dic As ObjectDictionary
    Set added_dic = New_ObjectDictionary()
    Call added_dic.Add("alpha", "value1")

    Dim initialized_dic As ObjectDictionary
    Set initialized_dic = New_ObjectDictionary("OtherValue")

    Dim other_obj As Test_ElementTypeProviderStub
    Set other_obj = New Test_ElementTypeProviderStub
    other_obj.ElementTypeKey = "OtherValue"

    ' --- Act ---
    Err.Clear
    Call added_dic.InitializeFromOther(source_dic)

    Dim actual_added_error_number As Long
    actual_added_error_number = Err.Number

    Dim actual_added_error_source As String
    actual_added_error_source = Err.Source

    Dim actual_added_error_description As String
    actual_added_error_description = Err.Description

    Err.Clear
    Call initialized_dic.InitializeFromOther(source_dic)

    Dim actual_initialized_error_number As Long
    actual_initialized_error_number = Err.Number

    Dim actual_initialized_error_source As String
    actual_initialized_error_source = Err.Source

    Dim actual_initialized_error_description As String
    actual_initialized_error_description = Err.Description

    Err.Clear
    Dim actual_added_count As Long
    actual_added_count = added_dic.Count

    Dim actual_added_value As String
    actual_added_value = added_dic.Item("alpha")

    Call initialized_dic.Add("other", other_obj)

    Dim actual_initialized_obj As Test_ElementTypeProviderStub
    Set actual_initialized_obj = initialized_dic.Item("other")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_added_error_number, actual_added_error_source, actual_added_error_description
    Assert.ErrorRaised 0, actual_initialized_error_number, actual_initialized_error_source, actual_initialized_error_description
    Assert.EqualsNumeric 1, actual_added_count
    Assert.Equals "value1", actual_added_value
    Assert.IsTrue actual_initialized_obj Is other_obj
End Sub

Public Sub Test_ObjectDictionary_AddOther_AddsKeyValuesToEmptyDictionaryInOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary("ObjectValue")
    source_dic.CompareMode = vbTextCompare

    Dim first_obj As Test_ElementTypeProviderStub
    Set first_obj = New Test_ElementTypeProviderStub
    first_obj.ElementTypeKey = "ObjectValue"

    Dim second_obj As Test_ElementTypeProviderStub
    Set second_obj = New Test_ElementTypeProviderStub
    second_obj.ElementTypeKey = "ObjectValue"

    Call source_dic.Add("KEY", first_obj)
    Call source_dic.Add("second", second_obj)

    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Call target_dic.AddOther(source_dic)

    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("key")

    Dim actual_first As Test_ElementTypeProviderStub
    Set actual_first = target_dic.Item("key")

    Dim actual_second As Test_ElementTypeProviderStub
    Set actual_second = target_dic.Item("second")

    Err.Clear
    Call target_dic.Add("bad", "primitive")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_count
    Assert.Equals "KEY", actual_keys(0)
    Assert.Equals "second", actual_keys(1)
    Assert.IsTrue actual_exists
    Assert.IsTrue actual_first Is first_obj
    Assert.IsTrue actual_second Is second_obj
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectList.")
End Sub

Public Sub Test_ObjectDictionary_AddOther_ExistingKeyErrorsByDefaultAndKeepsState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "target")
    Call target_dic.Add("omega", "tail")

    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary()
    Call source_dic.Add("beta", "new")
    Call source_dic.Add("alpha", "source")

    ' --- Act ---
    Err.Clear
    Call target_dic.AddOther(source_dic)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_alpha As String
    actual_alpha = target_dic.Item("alpha")

    Dim actual_beta_exists As Boolean
    actual_beta_exists = target_dic.Exists("beta")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectDictionary.AddOther")
    Assert.EqualsNumeric 2, actual_count
    Assert.Equals "alpha", actual_keys(0)
    Assert.Equals "omega", actual_keys(1)
    Assert.Equals "target", actual_alpha
    Assert.IsFalse actual_beta_exists
End Sub

Public Sub Test_ObjectDictionary_AddOther_ErrorIfExistsFalseSkipsExistingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "target")

    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary()
    Call source_dic.Add("alpha", "source")
    Call source_dic.Add("beta", "new")

    ' --- Act ---
    Call target_dic.AddOther(source_dic, ErrorIfExists:=False)

    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_alpha As String
    actual_alpha = target_dic.Item("alpha")

    Dim actual_beta As String
    actual_beta = target_dic.Item("beta")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_count
    Assert.Equals "alpha", actual_keys(0)
    Assert.Equals "beta", actual_keys(1)
    Assert.Equals "target", actual_alpha
    Assert.Equals "new", actual_beta
End Sub

Public Sub Test_ObjectDictionary_AddOther_UpdateIfExistsTrueUpdatesExistingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "target")
    Call target_dic.Add("omega", "tail")

    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary()
    Call source_dic.Add("alpha", "updated")
    Call source_dic.Add("beta", "new")

    ' --- Act ---
    Call target_dic.AddOther(source_dic, UpdateIfExists:=True)

    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_alpha As String
    actual_alpha = target_dic.Item("alpha")

    Dim actual_omega As String
    actual_omega = target_dic.Item("omega")

    Dim actual_beta As String
    actual_beta = target_dic.Item("beta")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, actual_count
    Assert.Equals "alpha", actual_keys(0)
    Assert.Equals "omega", actual_keys(1)
    Assert.Equals "beta", actual_keys(2)
    Assert.Equals "updated", actual_alpha
    Assert.Equals "tail", actual_omega
    Assert.Equals "new", actual_beta
End Sub

Public Sub Test_ObjectDictionary_AddOther_UpdateTypeViolationKeepsExistingValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ObjectValue")

    Dim existing_obj As Test_ElementTypeProviderStub
    Set existing_obj = New Test_ElementTypeProviderStub
    existing_obj.ElementTypeKey = "ObjectValue"

    Call target_dic.Add("alpha", existing_obj)

    Dim source_dic As ObjectDictionary
    Set source_dic = New_ObjectDictionary()
    Call source_dic.Add("alpha", "primitive")

    ' --- Act ---
    Err.Clear
    Call target_dic.AddOther(source_dic, UpdateIfExists:=True)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    Dim actual_obj As Test_ElementTypeProviderStub
    Set actual_obj = target_dic.Item("alpha")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectList.")
    Assert.EqualsNumeric 1, actual_count
    Assert.IsTrue actual_obj Is existing_obj
End Sub

Public Sub Test_ObjectDictionary_Update_UpdatesExistingKeyAndMissingKeyRaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")

    ' --- Act ---
    Call target_dic.Update("alpha", "updated")

    Dim actual_value As String
    actual_value = target_dic.Item("alpha")

    Err.Clear
    Call target_dic.Update("missing", "value2")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "updated", actual_value
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectDictionary.")
    Assert.EqualsNumeric 1, actual_count
End Sub

Public Sub Test_ObjectDictionary_AddOrUpdate_AddsWhenMissingAndUpdatesWhenExisting(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()

    ' --- Act ---
    Call target_dic.AddOrUpdate("alpha", "value1")
    Call target_dic.AddOrUpdate("alpha", "updated")
    Call target_dic.AddOrUpdate("beta", "value2")

    Dim actual_alpha As String
    actual_alpha = target_dic.Item("alpha")

    Dim actual_beta As String
    actual_beta = target_dic.Item("beta")

    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_count
    Assert.Equals "updated", actual_alpha
    Assert.Equals "value2", actual_beta
End Sub

Public Sub Test_ObjectDictionary_ItemAssignmentActsAsAddOrUpdate(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim primitive_dic As ObjectDictionary
    Set primitive_dic = New_ObjectDictionary()

    Dim object_dic As ObjectDictionary
    Set object_dic = New_ObjectDictionary("ObjectValue")

    Dim item_obj As Test_ElementTypeProviderStub
    Set item_obj = New Test_ElementTypeProviderStub
    item_obj.ElementTypeKey = "ObjectValue"

    ' --- Act ---
    primitive_dic.Item("alpha") = "value1"
    primitive_dic.Item("alpha") = "updated"
    Set object_dic.Item("object") = item_obj

    Dim actual_alpha As String
    actual_alpha = primitive_dic.Item("alpha")

    Dim actual_obj As Test_ElementTypeProviderStub
    Set actual_obj = object_dic.Item("object")

    Dim actual_primitive_count As Long
    actual_primitive_count = primitive_dic.Count

    Dim actual_object_count As Long
    actual_object_count = object_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_primitive_count
    Assert.EqualsNumeric 1, actual_object_count
    Assert.Equals "updated", actual_alpha
    Assert.IsTrue actual_obj Is item_obj
End Sub

Public Sub Test_ObjectDictionary_Remove_RemovesExistingKeyAndCompactsRemainingKeyValuePairs(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")
    Call target_dic.Add("beta", "value2")
    Call target_dic.Add("gamma", "value3")

    ' --- Act ---
    Dim removed_value As String
    removed_value = target_dic.Remove("beta")

    Dim actual_alpha As String
    actual_alpha = target_dic.Item("alpha")

    Dim actual_gamma As String
    actual_gamma = target_dic.Item("gamma")

    Dim actual_beta_exists As Boolean
    actual_beta_exists = target_dic.Exists("beta")

    Call target_dic.Add("delta", "value4")

    Dim actual_delta As String
    actual_delta = target_dic.Item("delta")

    Err.Clear
    Call target_dic.Remove("missing")

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "value2", removed_value
    Assert.Equals "value1", actual_alpha
    Assert.Equals "value3", actual_gamma
    Assert.Equals "value4", actual_delta
    Assert.IsFalse actual_beta_exists
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class ObjectDictionary.")
    Assert.EqualsNumeric 3, actual_count
End Sub

Public Sub Test_ObjectDictionary_RemoveAll_RemovesKeyValuesAndKeepsCompareModeAndTypeContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ObjectValue")
    target_dic.CompareMode = vbTextCompare

    Dim first_obj As Test_ElementTypeProviderStub
    Set first_obj = New Test_ElementTypeProviderStub
    first_obj.ElementTypeKey = "ObjectValue"

    Dim second_obj As Test_ElementTypeProviderStub
    Set second_obj = New Test_ElementTypeProviderStub
    second_obj.ElementTypeKey = "ObjectValue"

    Call target_dic.Add("Key", first_obj)

    ' --- Act ---
    Call target_dic.RemoveAll

    Dim actual_count_after_remove_all As Long
    actual_count_after_remove_all = target_dic.Count

    Dim actual_mode As Long
    actual_mode = target_dic.CompareMode

    Call target_dic.Add("KEY", second_obj)

    Dim actual_exists As Boolean
    actual_exists = target_dic.Exists("key")

    Dim actual_obj As Test_ElementTypeProviderStub
    Set actual_obj = target_dic.Item("key")

    Err.Clear
    Call target_dic.Add("bad", "primitive")

    Dim actual_type_error_number As Long
    actual_type_error_number = Err.Number

    Dim actual_type_error_source As String
    actual_type_error_source = Err.Source

    Dim actual_type_error_description As String
    actual_type_error_description = Err.Description

    Err.Clear
    target_dic.CompareMode = vbBinaryCompare

    Dim actual_mode_error_number As Long
    actual_mode_error_number = Err.Number

    Dim actual_mode_error_source As String
    actual_mode_error_source = Err.Source

    Dim actual_mode_error_description As String
    actual_mode_error_description = Err.Description

    Err.Clear
    Dim actual_count As Long
    actual_count = target_dic.Count

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_count_after_remove_all
    Assert.EqualsNumeric vbTextCompare, actual_mode
    Assert.IsTrue actual_exists
    Assert.IsTrue actual_obj Is second_obj
    Assert.ErrorRaised 0, actual_type_error_number, actual_type_error_source, actual_type_error_description
    Assert.ErrorRaised 0, actual_mode_error_number, actual_mode_error_source, actual_mode_error_description
    Assert.IsTrue StartsWith(actual_mode_error_source, "Class ObjectDictionary.")
    Assert.EqualsNumeric 1, actual_count
End Sub

Public Sub Test_ObjectDictionary_Keys_ReturnsAddedOrderZeroBasedArrayAndEmptyArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim empty_dic As ObjectDictionary
    Set empty_dic = New_ObjectDictionary()

    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")
    Call target_dic.Add("beta", "value2")
    Call target_dic.Add("gamma", "value3")

    ' --- Act ---
    Dim empty_keys() As Variant
    empty_keys = empty_dic.Keys

    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(empty_keys)
    Assert.EqualsNumeric -1, UBound(empty_keys)
    Assert.EqualsNumeric 0, LBound(actual_keys)
    Assert.EqualsNumeric 2, UBound(actual_keys)
    Assert.Equals "alpha", actual_keys(0)
    Assert.Equals "beta", actual_keys(1)
    Assert.Equals "gamma", actual_keys(2)
End Sub

Public Sub Test_ObjectDictionary_ItemsAndConvertToArray_ReturnValuesInKeyOrderAndItemsIsCopy(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")
    Call target_dic.Add("beta", "value2")
    Call target_dic.Add("gamma", "value3")

    ' --- Act ---
    Dim item_list As ObjectList
    Set item_list = target_dic.Items
    Call item_list.Update(1, "changed-copy")

    Dim actual_arr() As Variant
    actual_arr = target_dic.ConvertToArray()

    Dim actual_original_beta As String
    actual_original_beta = target_dic.Item("beta")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, item_list.Count
    Assert.Equals "value1", item_list.Item(0)
    Assert.Equals "changed-copy", item_list.Item(1)
    Assert.Equals "value3", item_list.Item(2)
    Assert.Equals "value2", actual_original_beta
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "value1", actual_arr(0)
    Assert.Equals "value2", actual_arr(1)
    Assert.Equals "value3", actual_arr(2)
End Sub

Public Sub Test_ObjectDictionary_ConvertToStringArray_ConvertsValuesToStringArrayInKeyOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim empty_dic As ObjectDictionary
    Set empty_dic = New_ObjectDictionary()

    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", 10&)
    Call target_dic.Add("beta", 20&)
    Call target_dic.Add("gamma", 30&)

    ' --- Act ---
    Dim empty_arr() As String
    empty_arr = empty_dic.ConvertToStringArray()

    Dim actual_arr() As String
    actual_arr = target_dic.ConvertToStringArray()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(empty_arr)
    Assert.EqualsNumeric -1, UBound(empty_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "10", actual_arr(0)
    Assert.Equals "20", actual_arr(1)
    Assert.Equals "30", actual_arr(2)
End Sub

Public Sub Test_ObjectDictionary_ItemsAndConvertToArray_ObjectValueReturnsReferences(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ObjectValue")

    Dim first_obj As Test_ElementTypeProviderStub
    Set first_obj = New Test_ElementTypeProviderStub
    first_obj.ElementTypeKey = "ObjectValue"

    Dim second_obj As Test_ElementTypeProviderStub
    Set second_obj = New Test_ElementTypeProviderStub
    second_obj.ElementTypeKey = "ObjectValue"

    Call target_dic.Add("first", first_obj)
    Call target_dic.Add("second", second_obj)

    ' --- Act ---
    Dim item_list As ObjectList
    Set item_list = target_dic.Items

    Dim actual_arr() As Variant
    actual_arr = target_dic.ConvertToArray()

    Dim actual_item_first As Test_ElementTypeProviderStub
    Set actual_item_first = item_list.Item(0)

    Dim actual_item_second As Test_ElementTypeProviderStub
    Set actual_item_second = item_list.Item(1)

    Dim actual_arr_first As Test_ElementTypeProviderStub
    Set actual_arr_first = actual_arr(0)

    Dim actual_arr_second As Test_ElementTypeProviderStub
    Set actual_arr_second = actual_arr(1)

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_item_first Is first_obj
    Assert.IsTrue actual_item_second Is second_obj
    Assert.IsTrue actual_arr_first Is first_obj
    Assert.IsTrue actual_arr_second Is second_obj
End Sub

Public Sub Test_ObjectDictionary_AfterUpdateAndRemove_KeysItemsConvertToArrayReflectOrderAndCompaction(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")
    Call target_dic.Add("beta", "value2")
    Call target_dic.Add("gamma", "value3")

    ' --- Act ---
    Call target_dic.Update("beta", "updated")
    Call target_dic.Remove("alpha")

    Dim actual_keys() As Variant
    actual_keys = target_dic.Keys

    Dim actual_items As ObjectList
    Set actual_items = target_dic.Items

    Dim actual_arr() As Variant
    actual_arr = target_dic.ConvertToArray()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_keys)
    Assert.EqualsNumeric 1, UBound(actual_keys)
    Assert.Equals "beta", actual_keys(0)
    Assert.Equals "gamma", actual_keys(1)
    Assert.EqualsNumeric 2, actual_items.Count
    Assert.Equals "updated", actual_items.Item(0)
    Assert.Equals "value3", actual_items.Item(1)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "updated", actual_arr(0)
    Assert.Equals "value3", actual_arr(1)
End Sub

Public Sub Test_ObjectDictionary_ForEach_DictionaryEnumeratesKeySnapshotAndItemsEnumeratesValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")
    Call target_dic.Add("beta", "value2")
    Call target_dic.Add("gamma", "value3")

    ' --- Act ---
    Dim actual_keys(0 To 2) As String
    Dim key_idx As Long
    key_idx = 0

    Dim key_item As Variant
    For Each key_item In target_dic
        actual_keys(key_idx) = CStr(key_item)
        key_idx = key_idx + 1
        If key_idx = 1 Then Call target_dic.Add("delta", "value4")
    Next key_item

    Dim actual_values(0 To 3) As String
    Dim value_idx As Long
    value_idx = 0

    Dim value_item As Variant
    For Each value_item In target_dic.Items
        actual_values(value_idx) = CStr(value_item)
        value_idx = value_idx + 1
    Next value_item

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, key_idx
    Assert.Equals "alpha", actual_keys(0)
    Assert.Equals "beta", actual_keys(1)
    Assert.Equals "gamma", actual_keys(2)
    Assert.EqualsNumeric 4, value_idx
    Assert.Equals "value1", actual_values(0)
    Assert.Equals "value2", actual_values(1)
    Assert.Equals "value3", actual_values(2)
    Assert.Equals "value4", actual_values(3)
End Sub

Public Sub Test_ObjectDictionary_PublicGetEnumeratorIsNotExposed(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary()
    Call target_dic.Add("alpha", "value1")

    ' --- Act ---
    Err.Clear
    Dim actual_value As Variant
    actual_value = CallByName(target_dic, "GetEnumerator", VbMethod)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Err.Clear

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
End Sub
