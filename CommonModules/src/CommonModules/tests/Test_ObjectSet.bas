Attribute VB_Name = "Test_ObjectSet"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the ObjectSet class.
'!
' #############################################################################

' -----------------------------------------------------------------------------
' Add/Item/Count
' -----------------------------------------------------------------------------

Public Sub Test_Add_AddStrings_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add("alpha")
    Call obj_set.Add("beta")
    Call obj_set.Add("gamma")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_Add_DuplicateWithErrorIgnored_KeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Call obj_set.Add("alpha", ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_Add_ArrayWithSameTypedValues_DetectsDuplicate(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_arr As Variant
    first_arr = Array(CLng(1), CStr("alpha"))

    Dim duplicate_arr As Variant
    duplicate_arr = Array(CLng(1), CStr("alpha"))

    Call obj_set.Add(first_arr)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_arr, ErrorIfExists:=False)

    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(duplicate_arr)

    Dim actual_item As Variant
    actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.IsTrue actual_exists
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue IsArray(actual_item)
End Sub

Public Sub Test_Add_ArrayWithDifferentElementTypes_AddsDifferentItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_arr As Variant
    first_arr = Array(CLng(1))

    Dim other_arr As Variant
    other_arr = Array(CStr(1))

    Call obj_set.Add(first_arr)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(other_arr, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_InitializedIEquatableElementTypeProvider_DetectsDuplicateByContractKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Initialize( _
            ElementTypeName:="ILeafCondition", _
            ObjectKeyMode:=G_OBJECT_KEY_MODE_I_EQUATABLE)

    Dim first_item As Test_ElementTypeEquatableStub
    Set first_item = New Test_ElementTypeEquatableStub
    first_item.ElementTypeKey = "ILeafCondition"
    first_item.IdentityKey = "same-id"

    Dim duplicate_item As Test_ElementTypeEquatableStub
    Set duplicate_item = New Test_ElementTypeEquatableStub
    duplicate_item.ElementTypeKey = "ILeafCondition"
    duplicate_item.IdentityKey = "same-id"

    ' Act
    Dim first_added As Boolean
    first_added = obj_set.Add(first_item)

    Dim duplicate_added As Boolean
    duplicate_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue first_added
    Assert.IsFalse duplicate_added
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue obj_set.HasItemTypeContract
    Assert.Equals "ILeafCondition", obj_set.ElementTypeName
    Assert.Equals "Object@ILeafCondition", obj_set.ItemTypeName
    Assert.EqualsNumeric G_OBJECT_KEY_MODE_I_EQUATABLE, obj_set.ObjectKeyMode
    Assert.Equals first_item, obj_set.Item(0)
End Sub

Public Sub Test_AddArray_StringArray_AddsEachItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim src_arr(0 To 2) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"
    src_arr(2) = "gamma"

    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Call obj_set.AddArray(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_AddOther_ObjectSet_AddsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    Dim other_set As ObjectSet
    Set other_set = New ObjectSet
    Call other_set.Add("beta")
    Call other_set.Add("gamma")

    ' Act
    Call obj_set.AddOther(other_set)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_AddList_EmptyInitializedList_PropagatesContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim obj_list As ObjectList
    Set obj_list = New_ObjectList( _
            ElementTypeName:="ILeafCondition", _
            ObjectKeyMode:=G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Act
    Call obj_set.AddList(obj_list)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.IsTrue obj_set.HasItemTypeContract
    Assert.Equals "ILeafCondition", obj_set.ElementTypeName
    Assert.EqualsNumeric G_OBJECT_KEY_MODE_I_EQUATABLE, obj_set.ObjectKeyMode
End Sub

Public Sub Test_AddList_ObjectListWithDuplicateIgnored_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("beta")
    Call obj_list.Add("alpha")
    Call obj_list.Add("gamma")

    ' Act
    Call obj_set.AddList(obj_list, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_ConvertToArray_Empty_ReturnsZeroBasedEmptyArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_set.ConvertToArray()

    ' Assert
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ConvertToStringArray_Empty_ReturnsZeroBasedEmptyArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ConvertToArray_WithItems_ReturnsArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_set.ConvertToArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
End Sub

Public Sub Test_ConvertToStringArray_WithItems_ReturnsStringArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
End Sub

Public Sub Test_ConvertToStringArray_ArrayItem_ReturnsTypedValueKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim item_arr As Variant
    item_arr = Array("alpha")
    Call obj_set.Add(item_arr)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:0](String(alpha))", actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_ModernCVErr_ReturnsErrorString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrSpill))

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#SPILL!", actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_Nothing_ReturnsTypedValueKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Object
    Set nothing_item = Nothing
    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals GetTypedValueKey(nothing_item), actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_IStringableObject_ReturnsToString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2)
    Call obj_set.Add(range_bounds)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals range_bounds.ToString(), actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_NonStringableObject_ReturnsTypedValueKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim item_obj As Test_ObjectSetEquatableStub
    Set item_obj = New Test_ObjectSetEquatableStub
    item_obj.IdentityKey = "alpha"
    Call obj_set.Add(item_obj)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals GetTypedValueKey(item_obj), actual_arr(0)
End Sub

Public Sub Test_ConvertToArray_ArrayItem_ReturnsArrayItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim item_arr As Variant
    item_arr = Array(CLng(1), CStr("alpha"))
    Call obj_set.Add(item_arr)

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_set.ConvertToArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 0, UBound(actual_arr)
    Assert.IsTrue IsArray(actual_arr(0))
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(actual_arr(0))
End Sub

Public Sub Test_FindItem_ArrayItem_ReturnsStoredArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_arr As Variant
    stored_arr = Array(CLng(1), CStr("alpha"))
    Call obj_set.Add(stored_arr)

    Dim search_arr As Variant
    search_arr = Array(CLng(1), CStr("alpha"))

    ' Act
    Dim actual_item As Variant
    actual_item = obj_set.FindItem(search_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsArray(actual_item)
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(actual_item)
End Sub

Public Sub Test_RemoveItem_ArrayItem_RemovesMatchingTypedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_arr As Variant
    stored_arr = Array(CLng(1), CStr("alpha"))
    Call obj_set.Add(stored_arr)

    Dim other_arr As Variant
    other_arr = Array(CLng(2), CStr("beta"))
    Call obj_set.Add(other_arr)

    Dim search_arr As Variant
    search_arr = Array(CLng(1), CStr("alpha"))

    ' Act
    Dim actual_item As Variant
    actual_item = obj_set.RemoveItem(search_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsArray(actual_item)
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(actual_item)
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "Variant[0:1](Long(2),String(beta))", GetTypedValueKey(obj_set.Item(0))
End Sub

Public Sub Test_AddList_ArrayItems_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_arr As Variant
    first_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(first_arr)

    Dim duplicate_arr As Variant
    duplicate_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(duplicate_arr)

    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Call obj_set.AddList(obj_list, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue IsArray(obj_set.Item(0))
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(obj_set.Item(0))
End Sub

Public Sub Test_AddOther_ArrayItems_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim src_set As ObjectSet
    Set src_set = New ObjectSet

    Dim first_arr As Variant
    first_arr = Array(CLng(1), CStr("alpha"))
    Call src_set.Add(first_arr)

    Dim dst_set As ObjectSet
    Set dst_set = New ObjectSet

    ' Act
    Call dst_set.AddOther(src_set, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, dst_set.Count
    Assert.IsTrue IsArray(dst_set.Item(0))
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(dst_set.Item(0))
End Sub

Public Sub Test_Add_DuplicateEquatableWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As WorksheetRangeBounds
    Set first_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Dim duplicate_item As WorksheetRangeBounds
    Set duplicate_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Add_Nothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_DuplicateNothingWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub

    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_RemoveItem_Nothing_RemovesNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub

    Call obj_set.Add(nothing_item)

    ' Act
    Dim removed_item As Test_ObjectSetEquatableStub
    Set removed_item = obj_set.RemoveItem(nothing_item)

    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_Add_NothingFirstThenDuplicateCheckable_DetectsDuplicateByGetKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetDupCheckStub

    Dim first_item As Test_ObjectSetDupCheckStub
    Set first_item = New Test_ObjectSetDupCheckStub
    first_item.DuplicateKey = "same-key"

    Dim duplicate_item As Test_ObjectSetDupCheckStub
    Set duplicate_item = New Test_ObjectSetDupCheckStub
    duplicate_item.DuplicateKey = "same-key"

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenDuplicateCheckable_RejectsOtherType(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetDupCheckStub

    Dim first_item As Test_ObjectSetDupCheckStub
    Set first_item = New Test_ObjectSetDupCheckStub
    first_item.DuplicateKey = "first-key"

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Add(other_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenEquatable_DetectsDuplicateByIdentityString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "same-id"

    Dim duplicate_item As Test_ObjectSetEquatableStub
    Set duplicate_item = New Test_ObjectSetEquatableStub
    duplicate_item.IdentityKey = "same-id"

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenEquatable_RejectsOtherType(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "first-id"

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Add(other_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenObjectReference_UsesReferenceIdentity(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Collection

    Dim first_item As Collection
    Set first_item = New Collection

    Dim same_ref_item As Collection
    Set same_ref_item = first_item

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_same_ref_added As Boolean
    actual_same_ref_added = obj_set.Add(same_ref_item, ErrorIfExists:=False)

    Dim actual_other_added As Boolean
    actual_other_added = obj_set.Add(other_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_same_ref_added
    Assert.IsTrue actual_other_added
    Assert.EqualsNumeric 3, obj_set.Count
End Sub

Public Sub Test_Add_ObjectImplementsDuplicateCheckableAndEquatable_UsesDuplicateCheckableKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetKeyPriorityStub
    Set first_item = New Test_ObjectSetKeyPriorityStub
    first_item.DuplicateKey = "same-duplicate-key"
    first_item.IdentityKey = "first-identity"

    Dim duplicate_by_key_item As Test_ObjectSetKeyPriorityStub
    Set duplicate_by_key_item = New Test_ObjectSetKeyPriorityStub
    duplicate_by_key_item.DuplicateKey = "same-duplicate-key"
    duplicate_by_key_item.IdentityKey = "second-identity"

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_by_key_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Add_DuplicateCheckableSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetDupCheckStub
    Set first_item = New Test_ObjectSetDupCheckStub
    first_item.DuplicateKey = "first-key"

    Dim nothing_item As Test_ObjectSetDupCheckStub

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Test_ObjectSetDupCheckStub
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_EquatableSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableStub

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Test_ObjectSetEquatableStub
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_ObjectReferenceSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Collection
    Set first_item = New Collection

    Dim nothing_item As Collection

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Collection
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_EquatableSetThenDuplicateNothingWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableStub

    Call obj_set.Add(first_item)
    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_added As Boolean
    If Err.Number = 0 Then
        actual_added = obj_set.Add(nothing_item, ErrorIfExists:=False)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_RemoveItem_EquatableSetThenNothing_RemovesNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableStub

    Call obj_set.Add(first_item)
    Call obj_set.Add(nothing_item)

    ' Act
    Dim removed_item As Test_ObjectSetEquatableStub
    Dim actual_exists As Boolean
    If Err.Number = 0 Then
        Set removed_item = obj_set.RemoveItem(nothing_item)
        If Err.Number = 0 Then
            actual_exists = obj_set.Exists(nothing_item)
        End If
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsFalse actual_exists
End Sub

' -----------------------------------------------------------------------------
' Update / FindItem
' -----------------------------------------------------------------------------

Public Sub Test_Update_EquatableSameIdentity_ReplacesStoredObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "same-id"

    Dim replacement_item As Test_ObjectSetEquatableStub
    Set replacement_item = New Test_ObjectSetEquatableStub
    replacement_item.IdentityKey = "same-id"

    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Update(0, replacement_item)

    Dim actual_item As Test_ObjectSetEquatableStub
    If Err.Number = 0 Then
        Set actual_item = obj_set.Item(0)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals replacement_item, actual_item
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue obj_set.Exists(first_item)
End Sub

Public Sub Test_Update_ReferenceKeyMismatch_RaisesObjectSetKeyMismatchError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As ObjectSet
    Set first_item = New ObjectSet

    Dim replacement_item As ObjectSet
    Set replacement_item = New ObjectSet

    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Update(0, replacement_item)

    Dim actual_source As String
    actual_source = Err.Source

    Dim actual_description As String
    actual_description = Err.Description

    Dim actual_item As ObjectSet
    Set actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue StartsWith(actual_source, "Class ObjectSet.")
    Assert.IsTrue 0 < InStr(1, actual_description, "object key", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "old_key: Object@ObjectSet(", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, " / new_key: Object@ObjectSet(", vbBinaryCompare)
    Assert.Equals first_item, actual_item
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Update_EquatableKeyMismatch_RaisesObjectSetKeyMismatchError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "first-id"

    Dim replacement_item As Test_ObjectSetEquatableStub
    Set replacement_item = New Test_ObjectSetEquatableStub
    replacement_item.IdentityKey = "other-id"

    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Update(0, replacement_item)

    Dim actual_source As String
    actual_source = Err.Source

    Dim actual_description As String
    actual_description = Err.Description

    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue StartsWith(actual_source, "Class ObjectSet.")
    Assert.IsTrue 0 < InStr(1, actual_description, "object key", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "old_key: IEquatable@Test_ObjectSetEquatableStub(first-id)", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "new_key: IEquatable@Test_ObjectSetEquatableStub(other-id)", vbBinaryCompare)
    Assert.Equals first_item, actual_item
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Update_DuplicateCheckableKeyMismatch_RaisesObjectSetKeyMismatchError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetDupCheckStub
    Set first_item = New Test_ObjectSetDupCheckStub
    first_item.DuplicateKey = "first-key"

    Dim replacement_item As Test_ObjectSetDupCheckStub
    Set replacement_item = New Test_ObjectSetDupCheckStub
    replacement_item.DuplicateKey = "other-key"

    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Update(0, replacement_item)

    Dim actual_source As String
    actual_source = Err.Source

    Dim actual_description As String
    actual_description = Err.Description

    Dim actual_item As Test_ObjectSetDupCheckStub
    Set actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue StartsWith(actual_source, "Class ObjectSet.")
    Assert.IsTrue 0 < InStr(1, actual_description, "object key", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "old_key: IDuplicateCheckable@Test_ObjectSetDupCheckStub(first-key)", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "new_key: IDuplicateCheckable@Test_ObjectSetDupCheckStub(other-key)", vbBinaryCompare)
    Assert.Equals first_item, actual_item
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Update_CVErrKeyMismatch_RaisesObjectSetKeyMismatchError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))

    ' Act
    Call obj_set.Update(0, CVErr(xlErrValue))

    Dim actual_source As String
    actual_source = Err.Source

    Dim actual_description As String
    actual_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue StartsWith(actual_source, "Class ObjectSet.")
    Assert.IsTrue 0 < InStr(1, actual_description, "object key", vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "old_key: " & GetTypedValueKey(CVErr(xlErrNA)), vbBinaryCompare)
    Assert.IsTrue 0 < InStr(1, actual_description, "new_key: " & GetTypedValueKey(CVErr(xlErrValue)), vbBinaryCompare)
    Assert.Equals CVErr(xlErrNA), obj_set.Item(0)
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_FindItem_Nothing_ReturnsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing actual_item
End Sub

Public Sub Test_FindItem_Equatable_ReturnsStoredObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_item As Test_ObjectSetEquatableStub
    Set stored_item = New Test_ObjectSetEquatableStub
    stored_item.IdentityKey = "same-id"

    Dim search_item As Test_ObjectSetEquatableStub
    Set search_item = New Test_ObjectSetEquatableStub
    search_item.IdentityKey = "same-id"

    Call obj_set.Add(stored_item)

    ' Act
    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals stored_item, actual_item
End Sub

Public Sub Test_FindItem_EquatableTypeMismatch_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_item As Test_ObjectSetEquatableStub
    Set stored_item = New Test_ObjectSetEquatableStub
    stored_item.IdentityKey = "same-id"
    Call obj_set.Add(stored_item)

    Dim search_item As Collection
    Set search_item = New Collection

    ' Act
    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(search_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals stored_item, obj_set.Item(0)
End Sub

' -----------------------------------------------------------------------------
' Exists/Remove
' -----------------------------------------------------------------------------

Public Sub Test_Exists_PresentAndMissing_ReturnsExpectedResult(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_present As Boolean
    Dim actual_missing As Boolean
    actual_present = obj_set.Exists("alpha")
    actual_missing = obj_set.Exists("gamma")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_present
    Assert.IsFalse actual_missing
End Sub

Public Sub Test_Exists_EmptySet_DoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists("1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_exists

    Call obj_set.Add(CLng(1))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_ExistsAndFindItem_EmptySetNothing_DoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_exists
    Assert.IsNothing actual_item

    Call obj_set.Add(CLng(1))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_ExistsFindItemAndRemoveItem_NothingOnlyNothing_FindsAndRemovesNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(nothing_item)

    Dim removed_item As Test_ObjectSetEquatableStub
    Set removed_item = obj_set.RemoveItem(nothing_item)

    Dim actual_exists_after_remove As Boolean
    actual_exists_after_remove = obj_set.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.IsFalse actual_exists_after_remove
End Sub

Public Sub Test_RemoveItem_NothingOnlyNonNothing_ReturnsMissingAndDoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableStub
    Call obj_set.Add(nothing_item)

    Dim search_item As Test_ObjectSetEquatableStub
    Set search_item = New Test_ObjectSetEquatableStub
    search_item.IdentityKey = "search-id"

    ' Act
    Dim removed_item As Test_ObjectSetEquatableStub
    Set removed_item = obj_set.RemoveItem(search_item, IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 1, obj_set.Count

    Call obj_set.Add(search_item)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.Equals search_item, obj_set.Item(1)
End Sub

Public Sub Test_Exists_PrimitiveTypeMismatch_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CLng(1))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(CStr(1))

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_Exists_PrimitiveSetThenNothing_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CLng(1))

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_FindItem_PrimitiveSetThenNothing_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CLng(1))

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    ' Act
    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.FindItem(nothing_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_PrimitiveSetThenNothing_RaisesTypeErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CLng(1))

    Dim nothing_item As Test_ObjectSetEquatableStub
    Set nothing_item = Nothing

    ' Act
    Dim actual_item As Test_ObjectSetEquatableStub
    Set actual_item = obj_set.RemoveItem(nothing_item, IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 1, obj_set.Item(0)
End Sub

Public Sub Test_Exists_EquatableTypeMismatch_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_item As Test_ObjectSetEquatableStub
    Set stored_item = New Test_ObjectSetEquatableStub
    stored_item.IdentityKey = "same-id"
    Call obj_set.Add(stored_item)

    Dim search_item As Collection
    Set search_item = New Collection

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(search_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals stored_item, obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_EquatableTypeMismatch_RaisesTypeErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_item As Test_ObjectSetEquatableStub
    Set stored_item = New Test_ObjectSetEquatableStub
    stored_item.IdentityKey = "same-id"
    Call obj_set.Add(stored_item)

    Dim search_item As Collection
    Set search_item = New Collection

    ' Act
    Dim removed_item As Test_ObjectSetEquatableStub
    Set removed_item = obj_set.RemoveItem(search_item, IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class ObjectSet.")
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals stored_item, obj_set.Item(0)
End Sub

Public Sub Test_Remove_RemoveByIndex_RemovesItemAndReturnsValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim removed_value As String
    removed_value = obj_set.Remove(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", removed_value
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "beta", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingWithIgnoreNotExists_DoesNotRaiseAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Dim removed_value As Variant
    removed_value = obj_set.RemoveItem("missing", IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingWithErrorNotIgnored_RaisesErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Dim removed_value As Variant
    removed_value = obj_set.RemoveItem("missing", IgnoreNotExists:=False)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingObjectWithIgnoreNotExists_ReturnsNothingAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim existing_item As WorksheetRangeBounds
    Set existing_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Dim missing_item As WorksheetRangeBounds
    Set missing_item = New_RangeBounds(Row:=2, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Call obj_set.Add(existing_item)

    ' Act
    Dim removed_item As WorksheetRangeBounds
    Set removed_item = obj_set.RemoveItem(missing_item, IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_RemoveItem_LastItem_KeepsTypeState(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(10&)

    ' Act
    Dim removed_value As Variant
    removed_value = obj_set.RemoveItem(10&)

    Err.Clear
    Call obj_set.Add("alpha")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.Equals GetTypeString(10&), obj_set.ItemTypeName

    Call obj_set.Add(20&)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_RemoveAll_AfterItems_KeepsTypeState(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(10&)

    ' Act
    Call obj_set.RemoveAll
    Call obj_set.Add("alpha")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.Equals GetTypeString(10&), obj_set.ItemTypeName

    Call obj_set.Add(20&)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

' -----------------------------------------------------------------------------
' CopySet
' -----------------------------------------------------------------------------

Public Sub Test_CopySet_WithItems_ReturnsCopiedSet(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_set As ObjectSet
    Set actual_set = obj_set.CopySet()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_set.Count
    Assert.Equals "alpha", actual_set.Item(0)
    Assert.Equals "beta", actual_set.Item(1)
End Sub

' -----------------------------------------------------------------------------
' Sort
' -----------------------------------------------------------------------------

Public Sub Test_Sort_WhenCalled_SortsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("gamma")
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Call obj_set.Sort

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

' -----------------------------------------------------------------------------
' Sort
' -----------------------------------------------------------------------------



Public Sub Test_Sort_ComparableObjectsWithNothing_SortsUsingObjectListContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Dim nothing_item As Test_ObjectListComparableStub
    Set nothing_item = Nothing
    Call obj_set.Add(pNewObjectSetComparableStub(2, "second"))
    Call obj_set.Add(nothing_item)
    Call obj_set.Add(pNewObjectSetComparableStub(1, "first"))

    ' Act
    Call obj_set.Sort(False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count

    Dim actual_nothing As Test_ObjectListComparableStub
    Set actual_nothing = obj_set.Item(0)
    Assert.IsNothing actual_nothing

    Dim actual_item As Test_ObjectListComparableStub
    Set actual_item = obj_set.Item(1)
    Assert.EqualsNumeric 1, actual_item.SortKey
    Set actual_item = obj_set.Item(2)
    Assert.EqualsNumeric 2, actual_item.SortKey
End Sub

Private Function pNewObjectSetComparableStub(ByVal SortKey As Long, ByVal ItemName As String) As Test_ObjectListComparableStub
    Dim result_value As Test_ObjectListComparableStub
    Set result_value = New Test_ObjectListComparableStub
    result_value.SortKey = SortKey
    result_value.ItemName = ItemName

    Set pNewObjectSetComparableStub = result_value
End Function

' -----------------------------------------------------------------------------
' GetEnumerator
' -----------------------------------------------------------------------------

Public Sub Test_GetEnumerator_WhenCalled_IteratesOverItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")
    Call obj_set.Add("gamma")

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator()

    Dim actual_value As String
    Do While enum_obj.MoveNext()
        If actual_value = "" Then
            actual_value = enum_obj.Current
        Else
            actual_value = actual_value & "," & enum_obj.Current
        End If
    Loop

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

Public Sub Test_ForEach_ObjectSet_IteratesOverSnapshot(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_value As String
    actual_value = ""
    Dim Count As Long
    Count = 0

    Dim set_item As Variant
    For Each set_item In obj_set
        If actual_value <> "" Then actual_value = actual_value & ","
        actual_value = actual_value & CStr(set_item)
        Count = Count + 1
    Next set_item

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta", actual_value
    Assert.EqualsNumeric 2, Count
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_ForEach_ObjectSet_AfterAdd_RefreshesSnapshot(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    Dim first_value As String
    first_value = ""
    Dim set_item As Variant
    For Each set_item In obj_set
        If first_value <> "" Then first_value = first_value & ","
        first_value = first_value & CStr(set_item)
    Next set_item

    ' Act
    Call obj_set.Add("gamma")

    Dim second_value As String
    second_value = ""
    For Each set_item In obj_set
        If second_value <> "" Then second_value = second_value & ","
        second_value = second_value & CStr(set_item)
    Next set_item

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta", first_value
    Assert.Equals "alpha,beta,gamma", second_value
End Sub
' -----------------------------------------------------------------------------
' Special Variant values
' -----------------------------------------------------------------------------

Public Sub Test_Add_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Call obj_set.Add(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_Add_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Call obj_set.Add(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_Update_CVErrWithSameError_DoesNotRaise(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))

    ' Act
    Call obj_set.Update(0, CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue IsError(obj_set.Item(0))
End Sub

Public Sub Test_Exists_CVErr_ReturnsTrueForSameError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_Exists_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_exists As Boolean
    actual_exists = False
    actual_exists = obj_set.Exists(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_exists
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_Exists_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_exists As Boolean
    actual_exists = False
    actual_exists = obj_set.Exists(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_exists
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_FindItem_CVErr_ReturnsStoredError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))

    ' Act
    Dim actual_item As Variant
    actual_item = obj_set.FindItem(CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals CVErr(xlErrNA), actual_item
End Sub

Public Sub Test_RemoveItem_CVErr_RemovesSameError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))
    Call obj_set.Add(CVErr(xlErrValue))

    ' Act
    Dim actual_item As Variant
    actual_item = obj_set.RemoveItem(CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals CVErr(xlErrNA), actual_item
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals CVErr(xlErrValue), obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_item As Variant
    actual_item = Empty
    actual_item = obj_set.RemoveItem(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue IsEmpty(actual_item)
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_RemoveItem_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_item As Variant
    actual_item = Empty
    actual_item = obj_set.RemoveItem(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue IsEmpty(actual_item)
    Assert.EqualsNumeric 0, obj_set.Count
End Sub

Public Sub Test_ConvertToStringArray_CVErr_ReturnsErrorString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add(CVErr(xlErrNA))

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#N/A", actual_arr(0)
End Sub
