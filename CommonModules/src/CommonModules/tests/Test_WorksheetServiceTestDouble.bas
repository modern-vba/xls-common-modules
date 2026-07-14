Attribute VB_Name = "Test_WorksheetServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the WorksheetServiceTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_WriteCell_WithSimpleValues_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 2, Sheet:="TestSheet") ' Row=1,Col=2

    ' Act
    sheet_srv.WriteCell range_bounds, "Hello", "m/d", True, True, False, True

    ' Assert
    ' 1) Whether one call is recorded.
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("WriteCell", range_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("WriteCell", range_bounds)
    Dim stored_range As WorksheetRangeBounds
    Set stored_range = stored_value.GetArgument(0)

    Assert.EqualsNumeric 7, stored_value.ArgumentCount
    Assert.IsTrue stored_range Is range_bounds
    Assert.Equals "Hello", stored_value.GetArgument(1)
    Assert.Equals "m/d", stored_value.GetArgument(2)
    Assert.IsTrue stored_value.GetArgument(3)          ' AsFormula = True
    Assert.IsTrue stored_value.GetArgument(4)          ' ClearWhenEmpty = True
    Assert.IsFalse stored_value.GetArgument(5)         ' IgnoreEmpty = False
    Assert.IsTrue stored_value.GetArgument(6)          ' TypeConvert = True
End Sub

Public Sub Test_WriteRange_With2DArray_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(2, 3, Sheet:="TestSheet")

    Dim arr_values(1 To 2, 1 To 2) As Variant
    arr_values(1, 1) = "A1"
    arr_values(1, 2) = "B1"
    arr_values(2, 1) = "A2"
    arr_values(2, 2) = "B2"

    ' Act
    sheet_srv.WriteRange range_bounds, arr_values

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("WriteRange", range_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("WriteRange", range_bounds)
    Dim copied_arr As Variant
    copied_arr = stored_value.GetArgument(1)

    Assert.Equals "A1", copied_arr(1, 1)
    Assert.Equals "B2", copied_arr(2, 2)
End Sub

Public Sub Test_WriteArrayFormula_WithSimpleFormula_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 3, Sheet:="TestSheet")

    ' Act
    sheet_srv.WriteArrayFormula range_bounds, "=TRANSPOSE({1,2,3})"

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("WriteArrayFormula", range_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("WriteArrayFormula", range_bounds)
    Assert.Equals "=TRANSPOSE({1,2,3})", stored_value.GetArgument(1)
End Sub

Public Sub Test_ReadCell_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(4, 1, Sheet:="TestSheet")

    Call test_dbl.Store.SetReturn("ReadCell", "StubValue", range_bounds, True)
    Call test_dbl.Store.SetOutput("ReadCell", "Expression", "StubValue", range_bounds, True)
    Call test_dbl.Store.SetOutput("ReadCell", "NumberFormat", "m/d", range_bounds, True)

    ' Act
    Dim actual_expr As String
    Dim actual_fmt As String
    sheet_srv.ReadCell range_bounds, actual_expr, actual_fmt, True

    ' Assert: whether the read value becomes StubValue.
    Assert.Equals "StubValue", actual_expr
    Assert.Equals "m/d", actual_fmt
End Sub

Public Sub Test_ReadCell_WithEquivalentRangeBounds_UsesIEquatableKey(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl

    Dim stored_range As WorksheetRangeBounds
    Set stored_range = New_RangeBounds(4, 1, Sheet:="TestSheet")

    Dim lookup_range As WorksheetRangeBounds
    Set lookup_range = New_RangeBounds(4, 1, Sheet:="TestSheet")

    Call test_dbl.Store.SetReturn("ReadCell", "StubValue", stored_range, True)
    Call test_dbl.Store.SetOutput("ReadCell", "Expression", "StubValue", stored_range, True)
    Call test_dbl.Store.SetOutput("ReadCell", "NumberFormat", "m/d", stored_range, True)

    ' Act
    Dim actual_expr As String
    Dim actual_fmt As String
    sheet_srv.ReadCell lookup_range, actual_expr, actual_fmt, True

    ' Assert
    Assert.Equals "StubValue", actual_expr
    Assert.Equals "m/d", actual_fmt
End Sub

Public Sub Test_CopyCell_WithAsValueTrue_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(1, 1, Sheet:="SrcSheet")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(2, 2, Sheet:="DstSheet")

    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds, True, True

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("CopyCell", dst_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("CopyCell", dst_bounds)
    ' Call arguments: SourceRangeBounds, DestinationRangeBounds, AsValue, CopyNumberFormat.

    Dim as_value As Boolean: as_value = stored_value.GetArgument(2)
    Dim copy_num_fmt As Boolean: copy_num_fmt = stored_value.GetArgument(3)

    Assert.IsTrue as_value
    Assert.IsTrue copy_num_fmt
End Sub

Public Sub Test_CopyRange_WithDefaultOptions_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(1, 2, Sheet:="SrcSheet")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(5, 5, Sheet:="DstSheet")

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("CopyRange", dst_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("CopyRange", dst_bounds)
    ' Call arguments: SourceRangeBounds, DestinationRangeBounds, AsValue, CopyNumberFormat.
    Assert.IsFalse stored_value.GetArgument(2) ' AsValue=False (default).
    Assert.IsFalse stored_value.GetArgument(3) ' CopyNumberFormat=False (default).
End Sub

Public Sub Test_IsEmptyCell_WithIgnoreEmptyFalse_ReturnsStub(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(10, 2, Sheet:="TestSheet")

    Call test_dbl.Store.SetReturn("IsEmptyCell", True, range_bounds, False)

    ' Act
    Dim actual_result As Boolean
    actual_result = test_dbl.IsEmptyCell(range_bounds, False)

    ' Assert
    Assert.IsTrue actual_result
End Sub

Public Sub Test_HasFormula_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="Sheet1")
    Call test_dbl.Store.SetReturn("HasFormula", True, range_bounds)

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.HasFormula(range_bounds)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_ClearRange_WithDefaults_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 3, Sheet:="TestSheet")

    ' Act
    sheet_srv.ClearRange range_bounds

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("ClearRange", range_bounds)

    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("ClearRange", range_bounds)
    ' Call arguments: RangeBounds, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks.
    Assert.IsFalse stored_value.GetArgument(1) ' ClearAll=False
    Assert.IsTrue stored_value.GetArgument(2)  ' ClearContents=True
    Assert.IsFalse stored_value.GetArgument(3) ' ClearNumberFormats=False
End Sub

Public Sub Test_RemoveDuplicates_WithColumns_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 1, 10, 4, Sheet:="TestSheet")
    Dim duplicate_columns As Variant
    duplicate_columns = Array(1, 3)

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, duplicate_columns)

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("RemoveDuplicates", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("RemoveDuplicates", range_bounds)
    Dim duplicate_columns_result As Variant
    duplicate_columns_result = stored_value.GetArgument(1)
    Assert.EqualsNumeric 1, duplicate_columns_result(0)
    Assert.EqualsNumeric 3, duplicate_columns_result(1)
End Sub

Public Sub Test_InsertRows_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 1, 8, 1, Sheet:="TestSheet")
    ' Act
    Call sheet_srv.InsertRows(range_bounds)
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("InsertRows", range_bounds)
    Assert.EqualsNumeric 1, test_dbl.Store.GetLatestCall("InsertRows", range_bounds).ArgumentCount
End Sub

Public Sub Test_DeleteRows_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 1, 8, 1, Sheet:="TestSheet")

    ' Act
    Call sheet_srv.DeleteRows(range_bounds)

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("DeleteRows", range_bounds)
    Assert.EqualsNumeric 1, test_dbl.Store.GetLatestCall("DeleteRows", range_bounds).ArgumentCount
End Sub

Public Sub Test_SetAllDataVisible_StoresCallArguments(ByVal Assert As UnitTestAssert)    ' Arrange
    Call InitializeCommonService
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")

    ' Act
    sheet_srv.SetAllDataVisible range_bounds

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetAllDataVisible", range_bounds)
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = test_dbl.Store.GetLatestCall("SetAllDataVisible", range_bounds)
    Assert.EqualsNumeric 1, stored_val.ArgumentCount
End Sub

Public Sub Test_SetSheetOutlineLevel_WithRowAndColumn_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="TestSheet")

    ' Act
    sheet_srv.SetSheetOutlineLevel range_bounds, 2, 3

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetSheetOutlineLevel", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetSheetOutlineLevel", range_bounds)
    ' Call arguments: RangeBounds, RowLevels, ColumnLevels.
    Assert.EqualsNumeric 2, stored_value.GetArgument(1)
    Assert.EqualsNumeric 3, stored_value.GetArgument(2)
End Sub

Public Sub Test_SetSheetTabColor_WithColorIndex_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")

    ' Act
    sheet_srv.SetSheetTabColor range_bounds, 5, -1

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetSheetTabColor", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetSheetTabColor", range_bounds)
    ' Call arguments: RangeBounds, TabColorIndex, TabColor.
    Assert.EqualsNumeric 5, stored_value.GetArgument(1)
    Assert.EqualsNumeric -1, stored_value.GetArgument(2)
End Sub

Public Sub Test_SetRangeColor_WithColorIndex_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(2, 2, Sheet:="Sheet2")

    ' Act
    sheet_srv.SetRangeColor range_bounds, 3, -1, 6, -1

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetRangeColor", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetRangeColor", range_bounds)
    ' Call arguments: RangeBounds, FontColorIndex, FontColor, InteriorColorIndex, InteriorColor.
    Assert.EqualsNumeric 3, stored_value.GetArgument(1)
    Assert.EqualsNumeric -1, stored_value.GetArgument(2)
    Assert.EqualsNumeric 6, stored_value.GetArgument(3)
End Sub

Public Sub Test_SetWrapText_True_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 10, Sheet:="WrapSheet")

    ' Act
    sheet_srv.SetWrapText range_bounds, True

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetWrapText", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetWrapText", range_bounds)
    ' Call arguments: RangeBounds, WrapText.
    Assert.IsTrue stored_value.GetArgument(1)
End Sub

Public Sub Test_SetShrinkToFit_True_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(9, 9, Sheet:="ShrinkSheet")

    ' Act
    sheet_srv.SetShrinkToFit range_bounds, True

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetShrinkToFit", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetShrinkToFit", range_bounds)
    ' Call arguments: RangeBounds, ShrinkToFit.
    Assert.IsTrue stored_value.GetArgument(1)
End Sub

Public Sub Test_SetAlignment_CustomOrientation_StoresCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 5, Sheet:="AlignSheet")

    ' Act
    sheet_srv.SetAlignment range_bounds, xlCenter, xlTop, 45, 1

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetAlignment", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetAlignment", range_bounds)
    ' Call arguments: RangeBounds, HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel.
    Assert.EqualsNumeric xlCenter, stored_value.GetArgument(1)
    Assert.EqualsNumeric xlTop, stored_value.GetArgument(2)
    Assert.EqualsNumeric 45, stored_value.GetArgument(3)
    Assert.EqualsNumeric 1, stored_value.GetArgument(4)
End Sub

Public Sub Test_SetAlignment_OmittedArguments_StoresNoChangeValuesInCallArguments(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 5, Sheet:="AlignSheet")

    ' Act
    sheet_srv.SetAlignment range_bounds

    ' Assert
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("SetAlignment", range_bounds)
    Dim stored_value As TestDoubleCallRecord
    Set stored_value = test_dbl.Store.GetLatestCall("SetAlignment", range_bounds)
    ' Call arguments: RangeBounds, HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel.
    Assert.EqualsNumeric G_ALIGNMENT_NO_CHANGE, stored_value.GetArgument(1)
    Assert.EqualsNumeric G_ALIGNMENT_NO_CHANGE, stored_value.GetArgument(2)
    Assert.EqualsNumeric G_ALIGNMENT_NO_CHANGE, stored_value.GetArgument(3)
    Assert.EqualsNumeric G_ALIGNMENT_NO_CHANGE, stored_value.GetArgument(4)
End Sub

Public Sub Test_EvaluateFormula_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="EvalSheet")

    Dim expected_value As Variant
    expected_value = 42
    Call test_dbl.Store.SetReturn("EvaluateFormula", expected_value, range_bounds, "=SUM(40,2)")

    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.EvaluateFormula(range_bounds, "=SUM(40,2)")

    ' Assert
    Assert.EqualsNumeric 42, actual_value
End Sub

Public Sub Test_XLookup_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl

    Dim lookup_bounds As WorksheetRangeBounds
    Set lookup_bounds = New_RangeBounds(2, 1, FinishRow:=5, FinishColumn:=1, Sheet:="LookupSheet")

    Dim return_bounds As WorksheetRangeBounds
    Set return_bounds = New_RangeBounds(2, 2, FinishRow:=5, FinishColumn:=2, Sheet:="LookupSheet")

    Dim expected_value As Variant
    expected_value = "Beta"
    Call test_dbl.Store.SetReturn("XLookup", expected_value, "B", lookup_bounds, return_bounds, "missing", 0&, 1&)

    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.XLookup("B", lookup_bounds, return_bounds, "missing")

    ' Assert
    Assert.Equals "Beta", CStr(actual_value)
End Sub

Public Sub Test_WriteCell_SameRangeTwice_RecordsBothCalls(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble
    Set test_dbl = New WorksheetServiceTestDouble

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = test_dbl

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "First")
    Call sheet_srv.WriteCell(range_bounds, "Second")

    ' Assert
    Assert.EqualsNumeric 2, test_dbl.Store.GetCallCount("WriteCell", range_bounds)

    Dim first_value As TestDoubleCallRecord
    Set first_value = test_dbl.Store.GetCall(0)

    Dim second_value As TestDoubleCallRecord
    Set second_value = test_dbl.Store.GetCall(1)

    Assert.Equals "First", first_value.GetArgument(1)
    Assert.Equals "Second", second_value.GetArgument(1)
End Sub

Public Sub Test_Find_RegisteredEmptyObjectList_ReturnsCountZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    Dim test_dbl As WorksheetServiceTestDouble
    Set test_dbl = New WorksheetServiceTestDouble

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")

    Dim expected_list As ObjectList
    Set expected_list = New ObjectList
    Call test_dbl.Store.SetReturn("Find", expected_list, "missing", range_bounds, True, True, True, True)

    Dim actual_list As ObjectList
    Set actual_list = test_dbl.Find("missing", range_bounds)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_list.Count
    Assert.EqualsNumeric 1, test_dbl.Store.GetCallCount("Find", "missing", range_bounds, True, True, True, True)
End Sub
