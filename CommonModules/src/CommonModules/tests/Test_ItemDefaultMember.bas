Attribute VB_Name = "Test_ItemDefaultMember"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for Item default member access.
'!
' #############################################################################

Public Sub Test_ObjectList_DefaultMemberGet_ReturnsSameValueAsItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    ' Act
    Dim actual_value As String
    actual_value = CStr(obj_list(0))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals obj_list.Item(0), actual_value
End Sub
Public Sub Test_ObjectList_DefaultMemberLet_UpdatesExistingValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    ' Act
    obj_list(0) = "updated"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "updated", obj_list.Item(0)
    Assert.EqualsNumeric 1, obj_list.Count
End Sub

Public Sub Test_ObjectList_DefaultMemberLet_OutOfRangeDoesNotAddItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    ' Act
    obj_list(1) = "beta"

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
End Sub

Public Sub Test_ObjectList_DefaultMemberSet_UpdatesExistingObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(New ObjectList)

    Dim expected_obj As ObjectList
    Set expected_obj = New ObjectList

    ' Act
    Set obj_list(0) = expected_obj

    Dim actual_obj As ObjectList
    Set actual_obj = obj_list.Item(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue (actual_obj Is expected_obj)
End Sub

Public Sub Test_ArrayObject_DefaultMemberGetAndLet_UsesItemContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=0)
    Call target_obj.Update(0, "alpha")

    ' Act
    Dim actual_before As String
    actual_before = CStr(target_obj(0))
    target_obj(0) = "updated"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", actual_before
    Assert.Equals "updated", target_obj.Item(0)
End Sub

Public Sub Test_ArrayObject_DefaultMemberSet_UpdatesExistingObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_obj As ArrayObject
    Set target_obj = New ArrayObject
    Call target_obj.ReDimArray(LowerBound:=0, UpperBound:=0)
    Call target_obj.Update(0, New ObjectList)

    Dim expected_obj As ObjectList
    Set expected_obj = New ObjectList

    ' Act
    Set target_obj(0) = expected_obj

    Dim actual_obj As ObjectList
    Set actual_obj = target_obj.Item(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue (actual_obj Is expected_obj)
End Sub

Public Sub Test_ObjectSet_DefaultMemberGet_ReturnsSameValueAsItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Dim actual_value As String
    actual_value = CStr(obj_set(0))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals obj_set.Item(0), actual_value
End Sub

Public Sub Test_ObjectDictionary_DefaultMemberGetAndLet_UsesItemContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim target_dic As ObjectDictionary
    Set target_dic = New ObjectDictionary
    Call target_dic.Add("alpha", "value1")

    ' Act
    Dim actual_before As String
    actual_before = CStr(target_dic("alpha"))
    target_dic("alpha") = "updated"
    target_dic("beta") = "added"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "value1", actual_before
    Assert.Equals "updated", target_dic.Item("alpha")
    Assert.Equals "added", target_dic.Item("beta")
End Sub

Public Sub Test_ObjectDictionary_DefaultMemberSet_UsesItemContract(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim target_dic As ObjectDictionary
    Set target_dic = New_ObjectDictionary("ObjectList")

    Dim expected_obj As ObjectList
    Set expected_obj = New ObjectList

    ' Act
    Set target_dic("object") = expected_obj

    Dim actual_obj As ObjectList
    Set actual_obj = target_dic.Item("object")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue (actual_obj Is expected_obj)
End Sub

Public Sub Test_WorksheetRangeBounds_DefaultMemberGet_ReturnsSameCellAsItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=3, FinishColumn:=4, Sheet:="S", Book:="B.xlsm")

    ' Act
    Dim expected_cell As WorksheetRangeBounds
    Set expected_cell = rng_bds.Item(1)

    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds(1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_cell.GetIdentityString(), actual_cell.GetIdentityString()
End Sub

Public Sub Test_WorksheetVirtualTable_DefaultMemberGet_ReturnsSameRowAsItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=3, Sheet:="S", Book:="B.xlsm"))

    Dim headers(0 To 0) As String
    headers(0) = "Name"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, headers)

    ' Act
    Dim expected_row As ObjectDictionary
    Set expected_row = table.Item(0)

    Dim actual_row As ObjectDictionary
    Set actual_row = table(0)

    Dim expected_bounds As WorksheetRangeBounds
    Set expected_bounds = expected_row.Item("Name")

    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = actual_row.Item("Name")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_bounds.GetIdentityString(), actual_bounds.GetIdentityString()
End Sub
