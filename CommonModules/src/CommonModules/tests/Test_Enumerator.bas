Attribute VB_Name = "Test_Enumerator"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the Enumerator class.
'!
' #############################################################################

' -----------------------------------------------------------------------------
' Array
' -----------------------------------------------------------------------------

Public Sub Test_MoveNext_Array_IteratesEachItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 2) As Long
    src_arr(0) = 10
    src_arr(1) = 20
    src_arr(2) = 30

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Act
    Dim actual_arr(0 To 2) As Long
    Dim idx As Long
    Do While enum_obj.MoveNext()
        actual_arr(idx) = enum_obj.Current
        idx = idx + 1
    Loop

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, idx
    Assert.EqualsNumeric 10, actual_arr(0)
    Assert.EqualsNumeric 20, actual_arr(1)
    Assert.EqualsNumeric 30, actual_arr(2)
End Sub

Public Sub Test_MoveNext_DescendingArray_IteratesFromLastItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 2) As Long
    src_arr(0) = 10
    src_arr(1) = 20
    src_arr(2) = 30

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, Descending:=True)

    ' Act
    Dim actual_arr(0 To 2) As Long
    Dim idx As Long
    Do While enum_obj.MoveNext()
        actual_arr(idx) = enum_obj.Current
        idx = idx + 1
    Loop

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, idx
    Assert.EqualsNumeric 30, actual_arr(0)
    Assert.EqualsNumeric 20, actual_arr(1)
    Assert.EqualsNumeric 10, actual_arr(2)
End Sub

Public Sub Test_Reset_AfterMoveNext_ReturnsToInitialPosition(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As Long
    src_arr(0) = 10
    src_arr(1) = 20

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext
    ignored = enum_obj.MoveNext

    ' Act
    Call enum_obj.Reset
    Dim moved As Boolean
    moved = enum_obj.MoveNext

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue moved
    Assert.EqualsNumeric 0, enum_obj.Index
    Assert.EqualsNumeric 10, enum_obj.Current
End Sub

Public Sub Test_Target_Array_ReturnsTargetArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 2) As Long
    src_arr(0) = 10
    src_arr(1) = 20
    src_arr(2) = 30

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Act
    Dim actual_arr As Variant
    actual_arr = enum_obj.Target

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsArray(actual_arr)
    Assert.EqualsNumeric 10, actual_arr(0)
    Assert.EqualsNumeric 20, actual_arr(1)
    Assert.EqualsNumeric 30, actual_arr(2)
End Sub

Public Sub Test_MoveNext_EmptyArray_ReturnsFalseAtEndIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array()

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Act
    Dim first_moved As Boolean
    first_moved = enum_obj.MoveNext()
    Dim second_moved As Boolean
    second_moved = enum_obj.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse first_moved
    Assert.IsFalse second_moved
    Assert.EqualsNumeric 0, enum_obj.Index
End Sub

Public Sub Test_MoveNext_DescendingEmptyArray_ReturnsFalseAtBeforeStartIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array()

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, Descending:=True)

    ' Act
    Dim first_moved As Boolean
    first_moved = enum_obj.MoveNext()
    Dim second_moved As Boolean
    second_moved = enum_obj.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse first_moved
    Assert.IsFalse second_moved
    Assert.EqualsNumeric -1, enum_obj.Index
End Sub

Public Sub Test_Current_EmptyArray_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array()

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim actual_value As Variant
    actual_value = enum_obj.Current

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetArrayEnumerator_UninitializedArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr() As Variant

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetArrayEnumerator_MultidimensionalArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1, 0 To 1) As Variant

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_Initialize_NothingObject_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_obj As ObjectList

    Dim enum_obj As Enumerator
    Set enum_obj = New Enumerator

    ' Act
    Call enum_obj.Initialize(target_obj)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
End Sub

Public Sub Test_Update_Array_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
End Sub

Public Sub Test_Update_Array_RaisesWithoutUpdatingSourceArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Err.Clear

    Assert.Equals "alpha", src_arr(0)
    Assert.Equals "beta", src_arr(1)
End Sub

Public Sub Test_Target_AfterFailedArrayUpdate_ReturnsOriginalArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()
    Call enum_obj.Update("updated")

    ' Act
    Err.Clear
    Dim actual_arr As Variant
    actual_arr = enum_obj.Target

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsArray(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
End Sub

Public Sub Test_Remove_Array_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
End Sub
' -----------------------------------------------------------------------------
' ObjectList
' -----------------------------------------------------------------------------

Public Sub Test_Target_ObjectList_ReturnsTargetObjectList(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = enum_obj.Target

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals obj_list, actual_list
End Sub

Public Sub Test_Reset_DescendingObjectList_ReturnsToLastItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(Descending:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Reset
    Dim moved As Boolean
    moved = enum_obj.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue moved
    Assert.EqualsNumeric 2, enum_obj.Index
    Assert.Equals "gamma", enum_obj.Current
End Sub

Public Sub Test_Remove_ObjectList_RemovesCurrentItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext
    ignored = enum_obj.MoveNext

    ' Act
    Dim removed_value As String
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "beta", removed_value
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
    Assert.Equals "gamma", obj_list.Item(1)
End Sub

Public Sub Test_Update_ObjectList_UpdatesCurrentItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext
    ignored = enum_obj.MoveNext

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", obj_list.Item(0)
    Assert.Equals "updated", obj_list.Item(1)
End Sub

Public Sub Test_Remove_DescendingObjectList_ContinuesWithNextLowerIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")
    Call obj_list.Add("delta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(Descending:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext

    ' Act
    Dim removed_value As String
    removed_value = enum_obj.Remove()
    Dim moved As Boolean
    moved = enum_obj.MoveNext

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "delta", removed_value
    Assert.IsTrue moved
    Assert.Equals "gamma", enum_obj.Current
End Sub

Public Sub Test_SkipTo_Count_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()

    ' Act
    Call enum_obj.SkipTo(obj_list.Count)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SkipTo_DescendingObjectList_AllowsForwardLowerIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(Descending:=True)

    ' Act
    Call enum_obj.SkipTo(2)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, enum_obj.Index
    Assert.Equals "gamma", enum_obj.Current
End Sub

Public Sub Test_Current_BeforeMoveNext_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()

    ' Act
    Dim actual_value As String
    actual_value = enum_obj.Current

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_Current_AfterMoveNextFalse_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator()
    Dim ignored As Boolean
    Do While enum_obj.MoveNext()
        ignored = True
    Loop

    ' Act
    Dim actual_value As String
    actual_value = enum_obj.Current

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_Current_DescendingBeforeMoveNext_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(Descending:=True)

    ' Act
    Dim actual_value As String
    actual_value = enum_obj.Current

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_Current_DescendingAfterMoveNextFalse_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(Descending:=True)
    Dim ignored As Boolean
    Do While enum_obj.MoveNext()
        ignored = True
    Loop

    ' Act
    Dim actual_value As String
    actual_value = enum_obj.Current

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub
' -----------------------------------------------------------------------------
' ObjectSet
' -----------------------------------------------------------------------------

Public Sub Test_MoveNext_ObjectSet_IteratesEachItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator()

    ' Act
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
    Assert.Equals "alpha,beta", actual_value
End Sub

Public Sub Test_Update_ObjectSet_UpdatesCurrentItemWithSameKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

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

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator()
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update(replacement_item)

    Dim actual_item As Test_ObjectSetEquatableStub
    If Err.Number = 0 Then
        Set actual_item = enum_obj.Current
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals replacement_item, actual_item
    Assert.Equals replacement_item, obj_set.Item(0)
End Sub

Public Sub Test_Remove_ObjectSet_RemovesCurrentItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator()
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As String
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", removed_value
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "beta", obj_set.Item(0)
End Sub
' -----------------------------------------------------------------------------
' IsReadOnly
' -----------------------------------------------------------------------------

Public Sub Test_IsReadOnly_Default_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 0) As String
    src_arr(0) = "alpha"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr)

    ' Act
    Dim actual_value As Boolean
    actual_value = enum_obj.IsReadOnly

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_value
End Sub

Public Sub Test_IsReadOnly_InitializeReadOnly_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 0) As String
    src_arr(0) = "alpha"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, IsReadOnly:=True)

    ' Act
    Dim actual_value As Boolean
    actual_value = enum_obj.IsReadOnly

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_value
End Sub

Public Sub Test_Update_ReadOnlyEnumerator_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As Enumerator
    Set enum_obj = New Enumerator
    Call enum_obj.Initialize(src_arr, IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Dim actual_arr As Variant
    actual_arr = enum_obj.Target
    Assert.Equals "alpha", actual_arr(0)
End Sub

Public Sub Test_Remove_ReadOnlyEnumerator_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As Enumerator
    Set enum_obj = New Enumerator
    Call enum_obj.Initialize(obj_list, IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
End Sub

Public Sub Test_Update_ReadOnlyArray_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Dim actual_arr As Variant
    actual_arr = enum_obj.Target
    Assert.Equals "alpha", actual_arr(0)
End Sub

Public Sub Test_Remove_ReadOnlyArray_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 1) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.Equals "alpha", enum_obj.Current
End Sub

Public Sub Test_Update_ReadOnlyObjectList_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update("updated")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.Equals "alpha", obj_list.Item(0)
End Sub

Public Sub Test_Remove_ReadOnlyObjectList_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_list.GetEnumerator(IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
End Sub

Public Sub Test_Update_ReadOnlyObjectSet_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

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

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator(IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Call enum_obj.Update(replacement_item)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.Equals first_item, obj_set.Item(0)
End Sub

Public Sub Test_Remove_ReadOnlyObjectSet_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator(IsReadOnly:=True)
    Dim ignored As Boolean
    ignored = enum_obj.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = enum_obj.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class Enumerator.")
    Assert.IsTrue (0 < InStr(Err.Description, "read-only"))
    Err.Clear

    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue obj_set.Exists("alpha")
End Sub
