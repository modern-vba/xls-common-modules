Attribute VB_Name = "Test_WorksheetRangeBounds"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the WorksheetRangeBounds class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

' -----------------------------------------------------------------------------
' ToString
' -----------------------------------------------------------------------------

Public Sub Test_ToString_CellIndex_ReturnsCorrectCellAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2)

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[" & ThisWorkbook.Name & "]Sheet1!B1", actual_value
End Sub

Public Sub Test_ToString_RangeIndex_ReturnsCorrectRangeAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=3, FinishColumn:=4)

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[" & ThisWorkbook.Name & "]Sheet1!B1:D3", actual_value
End Sub

Public Sub Test_ToString_RowIndex_ReturnsCorrectRangeAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, FinishRow:=3)

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[" & ThisWorkbook.Name & "]Sheet1!1:3", actual_value
End Sub

Public Sub Test_ToString_ColumnIndex_ReturnsCorrectRangeAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Column:=2, FinishColumn:=4)

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[" & ThisWorkbook.Name & "]Sheet1!B:D", actual_value
End Sub

Public Sub Test_ToString_WithWorkbookName_ReturnsCorrectRangeAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[MyBook.xlsm]Data!A1", actual_value
End Sub

Public Sub Test_ToString_EmptyRangeBounds_ReturnsCorrectRangeAddressString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=0, FinishColumn:=0, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.Equals "[MyBook.xlsm]Data!A1:EMPTY(R0,C0)", actual_value
End Sub

Public Sub Test_ToString_EmptyEntireRow_ReturnsSingleRowEmptyAddress(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, FinishRow:=2, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.IsTrue range_bounds.IsEmpty
    Assert.IsTrue range_bounds.IsEntireRow
    Assert.Equals "[MyBook.xlsm]Data!3:3:EMPTY", actual_value
End Sub

Public Sub Test_ToString_EmptyEntireColumn_ReturnsSingleColumnEmptyAddress(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Column:=4, FinishColumn:=3, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString()

    ' Assert
    Assert.IsTrue range_bounds.IsEmpty
    Assert.IsTrue range_bounds.IsEntireColumn
    Assert.Equals "[MyBook.xlsm]Data!D:D:EMPTY", actual_value
End Sub

Public Sub Test_ToString_CellOnlyTrue_ReturnsCellReferenceOnly(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=3, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="AnySheet", Book:="AnyBook.xlsm")

    ' Act
    Dim cell_str As String
    cell_str = rng_bds.ToString(CellOnly:=True)

    ' Assert
    ' No book or sheet name; only a bare "D3".
    Assert.Equals "D3", cell_str
End Sub

Public Sub Test_ToString_CellOnlyTrue_ForAreaRange(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=5, Sheet:="S", Book:="B.xlsm")

    ' Act
    Dim address_str As String
    address_str = rng_bds.ToString(CellOnly:=True)

    ' Assert
    ' No book or sheet; only "B2:E4" (default AddressType="A1").
    Assert.Equals "B2:E4", address_str
End Sub

Public Sub Test_ToString_Uninitialized_ReturnsUninitializedString(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds ' Uninitialized.

    ' Act
    Dim str_val As String
    str_val = rng_bds.ToString

    ' Assert
    ' Returns a string like "UNINITIALIZED(ObjPtr)".
    Assert.IsTrue InStr(str_val, "UNINITIALIZED") > 0
End Sub

Public Sub Test_TransformAbsolute_Uninitialized_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.TransformAbsolute()

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The object is not initialized.", Err.Description
    Err.Clear
End Sub

Public Sub Test_Equals_Uninitialized_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    Dim other_bounds As WorksheetRangeBounds
    Set other_bounds = New_RangeBounds(Row:=1, Column:=1)

    ' Act
    Dim actual_value As Boolean
    actual_value = rng_bds.Equals(other_bounds)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The object is not initialized.", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetIdentityString_Uninitialized_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    Dim actual_value As String
    actual_value = rng_bds.GetIdentityString()

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The object is not initialized.", Err.Description
    Err.Clear
End Sub

Public Sub Test_ToString_AbsoluteCell_ReturnsAbsoluteA1Address(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString( _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            IsAbsoluteFinishRow:=True, _
            IsAbsoluteFinishColumn:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "[MyBook.xlsm]Data!$A$1", actual_value
End Sub

Public Sub Test_ToString_AbsoluteRowRange_ReturnsAbsoluteA1Address(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, FinishRow:=3, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString(IsAbsoluteStartRow:=True, IsAbsoluteFinishRow:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "[MyBook.xlsm]Data!$1:$3", actual_value
End Sub

Public Sub Test_ToString_AbsoluteColumnRange_ReturnsAbsoluteA1Address(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Column:=2, FinishColumn:=4, Book:="MyBook.xlsm", Sheet:="Data")

    ' Act
    Dim actual_value As String
    actual_value = range_bounds.ToString(IsAbsoluteStartColumn:=True, IsAbsoluteFinishColumn:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "[MyBook.xlsm]Data!$B:$D", actual_value
End Sub

' -----------------------------------------------------------------------------
' Initialize
' -----------------------------------------------------------------------------

Public Sub Test_Initialize_StartRow2FinishRow1_BecomesEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    ' row=2, finishRow=1 => 1 < 2 ¨ finishRow=0 => empty
    rng_bds.Initialize Row:=2, Column:=3, FinishRow:=1, FinishColumn:=4

    ' Assert
    Assert.IsTrue rng_bds.IsEmpty
    Assert.EqualsNumeric 2, rng_bds.Row
    Assert.EqualsNumeric 3, rng_bds.Column
    Assert.EqualsNumeric 0, rng_bds.FinishRow
    Assert.EqualsNumeric 4, rng_bds.FinishColumn
End Sub

Public Sub Test_Initialize_StartColumn5FinishColumn3_BecomesEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    ' col=5, finishCol=3 => 3<5 => finishCol=0 => empty
    rng_bds.Initialize Row:=10, Column:=5, FinishRow:=12, FinishColumn:=3

    ' Assert
    Assert.IsTrue rng_bds.IsEmpty
    Assert.EqualsNumeric 10, rng_bds.Row
    Assert.EqualsNumeric 5, rng_bds.Column
    Assert.EqualsNumeric 12, rng_bds.FinishRow
    Assert.EqualsNumeric 0, rng_bds.FinishColumn
End Sub

Public Sub Test_Initialize_StartRowNegative_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    rng_bds.Initialize Row:=-1, Column:=2, FinishRow:=5, FinishColumn:=6

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Initialize_FinishRowTooBig_ClampedToG_ROW_MAX(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    ' Act
    rng_bds.Initialize Row:=1, FinishRow:=999999999, Sheet:="S"

    ' Assert
    ' Confirm that finishRow is clamped to G_ROW_MAX.
    Assert.EqualsNumeric G_ROW_MAX, rng_bds.FinishRow
End Sub

Public Sub Test_Initialize_Twice_RaisesError(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New WorksheetRangeBounds

    rng_bds.Initialize 1, 1, 2, 2

    On Error Resume Next
    rng_bds.Initialize 3, 3, 4, 4 ' Second initialization raises an error.
    Dim err_num As Long: err_num = Err.Number
    On Error GoTo 0

    Assert.IsTrue (err_num <> 0)
End Sub

Public Sub Test_Initialize_StartRowGreaterThanFinishRow_BecomesEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=5, FinishRow:=3, Column:=2, FinishColumn:=4

    ' Assert
    ' => pWellFormBounds sets finishRow=0, making the range empty.
    Assert.IsTrue rng.IsEmpty
    Assert.EqualsNumeric 5, rng.Row
    Assert.EqualsNumeric 0, rng.FinishRow
    Assert.EqualsNumeric 2, rng.Column
    Assert.EqualsNumeric 4, rng.FinishColumn
End Sub

Public Sub Test_Initialize_StartColumnZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=1, Column:=0

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Initialize_FinishRowNegative_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=1, Column:=1, FinishRow:=-1, FinishColumn:=1

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Initialize_FinishColumnNegative_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=1, Column:=1, FinishRow:=1, FinishColumn:=-1

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Initialize_StartRowTooBig_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=G_ROW_MAX + 1, Column:=1

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Initialize_StartColumnTooBig_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng As WorksheetRangeBounds
    Set rng = New WorksheetRangeBounds

    ' Act
    rng.Initialize Row:=1, Column:=G_COL_MAX + 1

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' Transform
' -----------------------------------------------------------------------------

Public Sub Test_Transform_WithSomeValues_ReturnsCorrentBounds(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=5, FinishColumn:=5)

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = src_bds.TransformAbsolute(Row:=2, Column:=3, FinishRow:=6, FinishColumn:=7)

    ' Assert
    Assert.EqualsNumeric 2, tf_bds.Row
    Assert.EqualsNumeric 3, tf_bds.Column
    Assert.EqualsNumeric 6, tf_bds.FinishRow
    Assert.EqualsNumeric 7, tf_bds.FinishColumn
End Sub

Public Sub Test_Transform_WhenRowColumnAreNotSpecified_KeepRowColumn(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    ' Row=2,Col=3,FinishRow=10,FinishColumn=12
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=10, FinishColumn:=12)

    ' Act
    ' Row and Column are omitted => keep old values.
    ' FinishRow=15 => changed
    ' FinishColumn=17 => changed
    Dim transform_bds As WorksheetRangeBounds
    Set transform_bds = rng_bds.TransformAbsolute(FinishRow:=15, FinishColumn:=17)

    ' Assert
    ' Row,Col ¨ 2,3
    Assert.EqualsNumeric 2, transform_bds.Row
    Assert.EqualsNumeric 3, transform_bds.Column
    ' FinishRow,FinishColumn ¨ 15,17
    Assert.EqualsNumeric 15, transform_bds.FinishRow
    Assert.EqualsNumeric 17, transform_bds.FinishColumn
End Sub

Public Sub Test_Transform_WhenFinishNotSpecified_KeepFinishRowFinishColumn(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=3, Column:=4, FinishRow:=8, FinishColumn:=9)

    ' Act
    ' Row=5 => changed
    ' Column=6 => changed
    ' FinishRow,FinishColumn => keep old
    Dim transform_bds As WorksheetRangeBounds
    Set transform_bds = rng_bds.TransformAbsolute(Row:=5, Column:=6)

    ' Assert
    Assert.EqualsNumeric 5, transform_bds.Row
    Assert.EqualsNumeric 6, transform_bds.Column
    Assert.EqualsNumeric 8, transform_bds.FinishRow
    Assert.EqualsNumeric 9, transform_bds.FinishColumn
End Sub

Public Sub Test_Transform_WithOmittedIndexes_KeepsOriginal(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = rng_bds.TransformAbsolute( _
            Row:=G_OMIT_CELL_INDEX, _
            Column:=G_OMIT_CELL_INDEX, _
            FinishRow:=G_OMIT_CELL_INDEX, _
            FinishColumn:=G_OMIT_CELL_INDEX)

    ' Assert
    Assert.EqualsNumeric 2, tf_bds.Row
    Assert.EqualsNumeric 2, tf_bds.Column
    Assert.EqualsNumeric 5, tf_bds.FinishRow
    Assert.EqualsNumeric 7, tf_bds.FinishColumn
End Sub

Public Sub Test_Transform_RowZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = rng_bds.TransformAbsolute(Row:=0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Transform_ColumnNegative_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = rng_bds.TransformAbsolute(Column:=-1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Transform_FinishRowNegative_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = rng_bds.TransformAbsolute(FinishRow:=-1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Transform_TooLargeAddRow_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.Transform(AddRow:=2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Transform_TooLargeAddColumn_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.Transform(AddColumn:=2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Transform_RowEmptyRange_TooSmallAddRow_StaysEmptyBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=0, FinishColumn:=6)

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.Transform(AddRow:=G_OMIT_CELL_INDEX)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_bounds.IsEmpty
    Assert.EqualsNumeric 5, actual_bounds.Row
    Assert.EqualsNumeric 0, actual_bounds.FinishRow
End Sub

Public Sub Test_Transform_EmptyRange_RemainsOrReformed(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    ' Row=2,Column=3,FinishRow=0,FinishColumn=0 => empty
    Set empty_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=0, FinishColumn:=0)

    ' Act
    ' Row=4 => changed?
    ' Column=5 => changed?
    ' This logic = if Row <=0 => keep pRow, else row=4...
    ' So presumably => row=4, col=5, finishRow=0, finishColumn=0 => still empty
    Dim new_bds As WorksheetRangeBounds
    Set new_bds = empty_bds.TransformAbsolute(Row:=4, Column:=5)

    ' Assert
    Assert.IsTrue new_bds.IsEmpty
    Assert.EqualsNumeric 4, new_bds.Row
    Assert.EqualsNumeric 5, new_bds.Column
    Assert.EqualsNumeric 0, new_bds.FinishRow
    Assert.EqualsNumeric 0, new_bds.FinishColumn
End Sub

Public Sub Test_Transform_InheritsWorksheetBook(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="Data", Book:="BookA.xlsm")

    ' Act
    Dim transform_bds As WorksheetRangeBounds
    Set transform_bds = rng_bds.TransformAbsolute(Row:=5, FinishRow:=6)

    ' Assert
    Assert.Equals "Data", transform_bds.WorksheetName
    Assert.Equals "BookA.xlsm", transform_bds.WorkbookName
    Assert.EqualsNumeric 5, transform_bds.Row
    Assert.EqualsNumeric 6, transform_bds.FinishRow
    Assert.EqualsNumeric rng_bds.Column, transform_bds.Column
    Assert.EqualsNumeric rng_bds.FinishColumn, transform_bds.FinishColumn
End Sub

Public Sub Test_Transform_UpdatesStartRowColumnButKeepsFinish(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src As WorksheetRangeBounds
    ' row=2,col=3,finishRow=5,finishCol=6
    Set src = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=6)

    ' Act
    ' Transform(Row:=4, Column:=5)
    ' => row=4, col=5, finishRow=5, finishCol=6
    Dim tf As WorksheetRangeBounds
    Set tf = src.TransformAbsolute(Row:=4, Column:=5)

    ' Assert
    Assert.EqualsNumeric 4, tf.Row
    Assert.EqualsNumeric 5, tf.Column
    Assert.EqualsNumeric 5, tf.FinishRow
    Assert.EqualsNumeric 6, tf.FinishColumn
End Sub

Public Sub Test_Transform_UpdatesFinishRowColumnButKeepsStart(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src As WorksheetRangeBounds
    ' row=2,col=3,finishRow=5,finishCol=6
    Set src = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=6)

    ' Act
    ' Transform(FinishRow:=10, FinishColumn:=12)
    ' => row=2,col=3, finishRow=10,finishCol=12
    Dim tf As WorksheetRangeBounds
    Set tf = src.TransformAbsolute(FinishRow:=10, FinishColumn:=12)

    ' Assert
    Assert.EqualsNumeric 2, tf.Row
    Assert.EqualsNumeric 3, tf.Column
    Assert.EqualsNumeric 10, tf.FinishRow
    Assert.EqualsNumeric 12, tf.FinishColumn
End Sub

Public Sub Test_Transform_EmptyRangeWithNewFinish_MightBecomeNonEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    ' row=3,col=4,finish=0 => empty
    Set empty_bds = New_RangeBounds(Row:=3, Column:=4, FinishRow:=0, FinishColumn:=0)

    ' Act
    ' Transform(FinishRow:=5,FinishColumn:=6)
    Dim new_bds As WorksheetRangeBounds
    Set new_bds = empty_bds.TransformAbsolute(FinishRow:=5, FinishColumn:=6)

    ' Assert
    ' row=3,col=4, finishRow=5,finishColumn=6 => now not empty
    Assert.IsFalse new_bds.IsEmpty
    Assert.EqualsNumeric 3, new_bds.Row
    Assert.EqualsNumeric 4, new_bds.Column
    Assert.EqualsNumeric 5, new_bds.FinishRow
    Assert.EqualsNumeric 6, new_bds.FinishColumn
End Sub

Public Sub Test_Transform_InheritsSheetAndBook(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=3, FinishColumn:=3, Sheet:="DataSheet", Book:="DataBook.xlsm")

    ' Act
    Dim tf_bds As WorksheetRangeBounds
    Set tf_bds = src_bds.TransformAbsolute(Row:=5, FinishRow:=10)

    ' Assert
    Assert.Equals "DataSheet", tf_bds.WorksheetName
    Assert.Equals "DataBook.xlsm", tf_bds.WorkbookName
    Assert.EqualsNumeric 5, tf_bds.Row
    Assert.EqualsNumeric 10, tf_bds.FinishRow
    ' Column,FinishColumn => same as src
    Assert.EqualsNumeric 1, tf_bds.Column
    Assert.EqualsNumeric 3, tf_bds.FinishColumn
End Sub

' -----------------------------------------------------------------------------
' Empty range bounds
' -----------------------------------------------------------------------------

Public Sub Test_Initialize_RowEmptyRange_PreservesColumnBounds(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange / Act
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=4, FinishColumn:=6)

    ' Assert
    Assert.IsTrue rng_bds.IsEmpty
    Assert.EqualsNumeric 5, rng_bds.Row
    Assert.EqualsNumeric 2, rng_bds.Column
    Assert.EqualsNumeric 0, rng_bds.FinishRow
    Assert.EqualsNumeric 6, rng_bds.FinishColumn
    Assert.EqualsNumeric 0, rng_bds.Count
End Sub

Public Sub Test_Initialize_ColumnEmptyRange_PreservesRowBounds(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange / Act
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=6)

    ' Assert
    Assert.IsTrue rng_bds.IsEmpty
    Assert.EqualsNumeric 3, rng_bds.Row
    Assert.EqualsNumeric 7, rng_bds.Column
    Assert.EqualsNumeric 8, rng_bds.FinishRow
    Assert.EqualsNumeric 0, rng_bds.FinishColumn
    Assert.EqualsNumeric 0, rng_bds.Count
End Sub

Public Sub Test_Intersect_RowEmptyRange_PreservesColumnOverlap(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=2, FinishRow:=3, FinishColumn:=5, Sheet:="S")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=4, Column:=3, FinishRow:=6, FinishColumn:=4, Sheet:="S")

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = rng1.Intersect(rng2)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 4, actual_bds.Row
    Assert.EqualsNumeric 3, actual_bds.Column
    Assert.EqualsNumeric 0, actual_bds.FinishRow
    Assert.EqualsNumeric 4, actual_bds.FinishColumn
End Sub

Public Sub Test_Intersect_ColumnEmptyRange_PreservesRowOverlap(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=2, Column:=1, FinishRow:=5, FinishColumn:=2, Sheet:="S")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=3, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="S")

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = rng1.Intersect(rng2)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 3, actual_bds.Column
    Assert.EqualsNumeric 4, actual_bds.FinishRow
    Assert.EqualsNumeric 0, actual_bds.FinishColumn
End Sub

' -----------------------------------------------------------------------------
' Transform size
' -----------------------------------------------------------------------------

Public Sub Test_Transform_AddRowColumn_ExpandsFinishOnly(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = src_bds.Transform(AddRow:=2, AddColumn:=1)

    ' Assert
    Assert.IsFalse actual_bds.IsEmpty
    Assert.EqualsNumeric 2, actual_bds.Row
    Assert.EqualsNumeric 3, actual_bds.Column
    Assert.EqualsNumeric 6, actual_bds.FinishRow
    Assert.EqualsNumeric 6, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_RowEmptyRange_AddRow_ReformsRowsPreservingColumns(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=4, FinishColumn:=6)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Transform(AddRow:=1)

    ' Assert
    Assert.IsFalse actual_bds.IsEmpty
    Assert.EqualsNumeric 5, actual_bds.Row
    Assert.EqualsNumeric 2, actual_bds.Column
    Assert.EqualsNumeric 5, actual_bds.FinishRow
    Assert.EqualsNumeric 6, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_ColumnEmptyRange_AddColumn_ReformsColumnsPreservingRows(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=6)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Transform(AddColumn:=1)

    ' Assert
    Assert.IsFalse actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 7, actual_bds.Column
    Assert.EqualsNumeric 8, actual_bds.FinishRow
    Assert.EqualsNumeric 7, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_EmptyRange_AddRowColumn_ReformsRange(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=5, Column:=7, FinishRow:=0, FinishColumn:=0)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Transform(AddRow:=2, AddColumn:=3)

    ' Assert
    Assert.IsFalse actual_bds.IsEmpty
    Assert.EqualsNumeric 5, actual_bds.Row
    Assert.EqualsNumeric 7, actual_bds.Column
    Assert.EqualsNumeric 6, actual_bds.FinishRow
    Assert.EqualsNumeric 9, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_ShrinkRowsToEmpty_PreservesColumnBounds(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=6)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = src_bds.Transform(AddRow:=-4)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 2, actual_bds.Row
    Assert.EqualsNumeric 3, actual_bds.Column
    Assert.EqualsNumeric 0, actual_bds.FinishRow
    Assert.EqualsNumeric 6, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_ShrinkRowsPastZero_BecomesEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=6, FinishColumn:=4)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = src_bds.Transform(AddRow:=-100)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 5, actual_bds.Row
    Assert.EqualsNumeric 2, actual_bds.Column
    Assert.EqualsNumeric 0, actual_bds.FinishRow
    Assert.EqualsNumeric 4, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_ShrinkColumnsPastZero_BecomesEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim src_bds As WorksheetRangeBounds
    Set src_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=9)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = src_bds.Transform(AddColumn:=-100)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 7, actual_bds.Column
    Assert.EqualsNumeric 8, actual_bds.FinishRow
    Assert.EqualsNumeric 0, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_RowEmptyRange_NegativeAddRow_StaysEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=0, FinishColumn:=4)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Transform(AddRow:=-100)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 5, actual_bds.Row
    Assert.EqualsNumeric 2, actual_bds.Column
    Assert.EqualsNumeric 0, actual_bds.FinishRow
    Assert.EqualsNumeric 4, actual_bds.FinishColumn
End Sub

Public Sub Test_Transform_ColumnEmptyRange_NegativeAddColumn_StaysEmpty(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=0)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Transform(AddColumn:=-100)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 7, actual_bds.Column
    Assert.EqualsNumeric 8, actual_bds.FinishRow
    Assert.EqualsNumeric 0, actual_bds.FinishColumn
End Sub

' -----------------------------------------------------------------------------
' Shift empty range bounds
' -----------------------------------------------------------------------------

Public Sub Test_Shift_RowEmptyRange_NegativeRowShift_KeepsFinishRowZero(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=0, FinishColumn:=6)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Shift(Row:=-2)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 2, actual_bds.Column
    Assert.EqualsNumeric 0, actual_bds.FinishRow
    Assert.EqualsNumeric 6, actual_bds.FinishColumn
End Sub

Public Sub Test_Shift_ColumnEmptyRange_NegativeColumnShift_KeepsFinishColumnZero(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim empty_bds As WorksheetRangeBounds
    Set empty_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=0)

    ' Act
    Dim actual_bds As WorksheetRangeBounds
    Set actual_bds = empty_bds.Shift(Column:=-2)

    ' Assert
    Assert.IsTrue actual_bds.IsEmpty
    Assert.EqualsNumeric 3, actual_bds.Row
    Assert.EqualsNumeric 5, actual_bds.Column
    Assert.EqualsNumeric 8, actual_bds.FinishRow
    Assert.EqualsNumeric 0, actual_bds.FinishColumn
End Sub

Public Sub Test_Shift_TooLargeRow_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.Shift(Row:=2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Shift_TooLargeColumn_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng_bds.Shift(Column:=2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Err.Clear
End Sub

Public Sub Test_Equals_BookAndSheetCaseDiff_DefaultReturnsTrue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="Data", Book:="BookA.xlsm")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="data", Book:="booka.xlsm")

    ' Act
    Dim actual_value As Boolean
    actual_value = rng1.Equals(rng2)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_GetIdentityString_BookAndSheetCaseDiff_DefaultMatches(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="Data", Book:="BookA.xlsm")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="data", Book:="booka.xlsm")

    ' Act
    Dim actual_value As Boolean
    actual_value = (rng1.GetIdentityString() = rng2.GetIdentityString())

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_Intersect_BookAndSheetCaseDiff_DefaultReturnsOverlap(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=4, FinishColumn:=4, Sheet:="Data", Book:="BookA.xlsm")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=5, Sheet:="data", Book:="booka.xlsm")

    ' Act
    Err.Clear
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = rng1.Intersect(rng2)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_bounds.Row
    Assert.EqualsNumeric 2, actual_bounds.Column
    Assert.EqualsNumeric 4, actual_bounds.FinishRow
    Assert.EqualsNumeric 4, actual_bounds.FinishColumn
End Sub

' -----------------------------------------------------------------------------
' Intersect
' -----------------------------------------------------------------------------

Public Sub Test_Intersect_BecomesEmpty_WhenOverlapIsNegative(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' rng1 => Row=1..3, col=1..3
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=3, FinishColumn:=3, Sheet:="S")

    ' rng2 => Row=4..5, col=4..5 => no overlap
    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=4, Column:=4, FinishRow:=5, FinishColumn:=5, Sheet:="S")

    ' Act
    Dim inter_bds As WorksheetRangeBounds
    Set inter_bds = rng1.Intersect(rng2)

    ' Assert
    Assert.IsTrue inter_bds.IsEmpty
    ' Finishing row/col = 0?
    Assert.EqualsNumeric 0, inter_bds.FinishRow
    Assert.EqualsNumeric 0, inter_bds.FinishColumn
End Sub

Public Sub Test_Intersect_BookOrSheetDiff_RaisesError(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rngA As WorksheetRangeBounds
    Set rngA = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="S", Book:="BookA.xlsm")

    Dim rngB As WorksheetRangeBounds
    Set rngB = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="S", Book:="BookB.xlsm")

    ' Act
    On Error Resume Next
    Dim tmp As WorksheetRangeBounds
    Set tmp = rngA.Intersect(rngB)
    Dim err_num As Long: err_num = Err.Number
    On Error GoTo 0

    ' Assert
    Assert.IsTrue (err_num <> 0)
End Sub

Public Sub Test_Intersect_PartialOverlap_ReturnsOverlapRange(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=2, Column:=2, FinishRow:=6, FinishColumn:=6, Sheet:="SheetX")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=4, Column:=1, FinishRow:=8, FinishColumn:=7, Sheet:="SheetX")

    ' overlap => Row=4..6, Col=2..6
    ' Act
    Dim over_bds As WorksheetRangeBounds
    Set over_bds = rng1.Intersect(rng2)

    ' Assert
    Assert.EqualsNumeric 4, over_bds.Row
    Assert.EqualsNumeric 2, over_bds.Column
    Assert.EqualsNumeric 6, over_bds.FinishRow
    Assert.EqualsNumeric 6, over_bds.FinishColumn
    Assert.IsFalse over_bds.IsEmpty
End Sub

Public Sub Test_Intersect_NoOverlap_BecomesEmptyRange(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rngA As WorksheetRangeBounds
    Set rngA = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="S")

    Dim rngB As WorksheetRangeBounds
    Set rngB = New_RangeBounds(Row:=3, Column:=3, FinishRow:=4, FinishColumn:=4, Sheet:="S")

    ' Act
    Dim intr As WorksheetRangeBounds
    Set intr = rngA.Intersect(rngB)

    ' Assert
    ' No overlap => if the code is updated to handle non-overlap, the result becomes empty.
    ' Example: (finishRow=0, finishCol=0)
    Assert.IsTrue intr.IsEmpty
End Sub

Public Sub Test_Intersect_DiffWorkbook_RaisesError(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng1 As WorksheetRangeBounds
    Set rng1 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=3, Sheet:="A", Book:="BookA.xlsm")

    Dim rng2 As WorksheetRangeBounds
    Set rng2 = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=3, Sheet:="A", Book:="BookB.xlsm")

    On Error Resume Next
    Dim tmp As WorksheetRangeBounds
    Set tmp = rng1.Intersect(rng2)
    Dim err_num As Long
    err_num = Err.Number
    On Error GoTo 0

    Assert.IsTrue (err_num <> 0)
End Sub

Public Sub Test_GetEnumerator_InvalidEnumerationMode_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5)

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = rng_bds.GetEnumerator(EnumerationMode:=999)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

' -----------------------------------------------------------------------------
' GetRows / GetColumns
' -----------------------------------------------------------------------------

Public Sub Test_GetRows_ReturnsTypedRowBoundsList(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=6)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = rng_bds.GetRows()

    ' Assert
    Assert.IsTrue actual_list.HasItemTypeContract
    Assert.Equals "WorksheetRangeBounds", actual_list.ElementTypeName
    Assert.EqualsNumeric 3, actual_list.Count

    Dim first_row As WorksheetRangeBounds
    Set first_row = actual_list.Item(0)
    Assert.EqualsNumeric 2, first_row.Row
    Assert.EqualsNumeric 2, first_row.FinishRow
    Assert.EqualsNumeric 3, first_row.Column
    Assert.EqualsNumeric 6, first_row.FinishColumn

    Dim last_row As WorksheetRangeBounds
    Set last_row = actual_list.Item(2)
    Assert.EqualsNumeric 4, last_row.Row
    Assert.EqualsNumeric 4, last_row.FinishRow
    Assert.EqualsNumeric 3, last_row.Column
    Assert.EqualsNumeric 6, last_row.FinishColumn
End Sub

Public Sub Test_GetColumns_ReturnsTypedColumnBoundsList(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=10, Column:=4, FinishRow:=11, FinishColumn:=6)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = rng_bds.GetColumns()

    ' Assert
    Assert.IsTrue actual_list.HasItemTypeContract
    Assert.Equals "WorksheetRangeBounds", actual_list.ElementTypeName
    Assert.EqualsNumeric 3, actual_list.Count

    Dim first_col As WorksheetRangeBounds
    Set first_col = actual_list.Item(0)
    Assert.EqualsNumeric 10, first_col.Row
    Assert.EqualsNumeric 11, first_col.FinishRow
    Assert.EqualsNumeric 4, first_col.Column
    Assert.EqualsNumeric 4, first_col.FinishColumn

    Dim last_col As WorksheetRangeBounds
    Set last_col = actual_list.Item(2)
    Assert.EqualsNumeric 10, last_col.Row
    Assert.EqualsNumeric 11, last_col.FinishRow
    Assert.EqualsNumeric 6, last_col.Column
    Assert.EqualsNumeric 6, last_col.FinishColumn
End Sub

Public Sub Test_GetRows_Descending_ReturnsRowsFromBottom(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=6)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = rng_bds.GetRows(Descending:=True)

    ' Assert
    Assert.EqualsNumeric 3, actual_list.Count

    Dim first_row As WorksheetRangeBounds
    Set first_row = actual_list.Item(0)
    Assert.EqualsNumeric 4, first_row.Row
    Assert.EqualsNumeric 4, first_row.FinishRow

    Dim last_row As WorksheetRangeBounds
    Set last_row = actual_list.Item(2)
    Assert.EqualsNumeric 2, last_row.Row
    Assert.EqualsNumeric 2, last_row.FinishRow
End Sub

Public Sub Test_GetColumns_Descending_ReturnsColumnsFromRight(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=10, Column:=4, FinishRow:=11, FinishColumn:=6)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = rng_bds.GetColumns(Descending:=True)

    ' Assert
    Assert.EqualsNumeric 3, actual_list.Count

    Dim first_col As WorksheetRangeBounds
    Set first_col = actual_list.Item(0)
    Assert.EqualsNumeric 6, first_col.Column
    Assert.EqualsNumeric 6, first_col.FinishColumn

    Dim last_col As WorksheetRangeBounds
    Set last_col = actual_list.Item(2)
    Assert.EqualsNumeric 4, last_col.Column
    Assert.EqualsNumeric 4, last_col.FinishColumn
End Sub

Public Sub Test_GetRowsAndColumns_EmptyRange_ReturnTypedEmptyLists(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=0, FinishColumn:=0)

    ' Act
    Dim row_list As ObjectList
    Set row_list = rng_bds.GetRows()

    Dim col_list As ObjectList
    Set col_list = rng_bds.GetColumns()

    ' Assert
    Assert.IsTrue row_list.HasItemTypeContract
    Assert.Equals "WorksheetRangeBounds", row_list.ElementTypeName
    Assert.EqualsNumeric 0, row_list.Count

    Assert.IsTrue col_list.HasItemTypeContract
    Assert.Equals "WorksheetRangeBounds", col_list.ElementTypeName
    Assert.EqualsNumeric 0, col_list.Count
End Sub

' -----------------------------------------------------------------------------
' GetEnumerator
' -----------------------------------------------------------------------------

Public Sub Test_GetEnumerator_Rows_EnumeratesEachRow(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' rows=2..4 => total 3 rows
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=6)

    Dim enum_obj As IEnumerator
    Set enum_obj = rng_bds.GetEnumerator()

    Dim row_idx As Long
    row_idx = 2

    ' Act
    Do While enum_obj.MoveNext
        Dim row_bds As WorksheetRangeBounds
        Set row_bds = enum_obj.Current

        ' Assert inside loop
        Assert.EqualsNumeric row_idx, row_bds.Row
        Assert.EqualsNumeric row_idx, row_bds.FinishRow
        Assert.EqualsNumeric 3, row_bds.Column
        Assert.EqualsNumeric 6, row_bds.FinishColumn

        row_idx = row_idx + 1
    Loop

    ' Rows 2..4 => MoveNext=True three times => row_idx=5 after loop.
    Assert.EqualsNumeric 5, row_idx
End Sub

Public Sub Test_GetEnumerator_Columns_EnumeratesEachColumn(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' col=4..6 => total3 => enumerates column by column
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=10, Column:=4, FinishRow:=11, FinishColumn:=6)

    Dim enum_obj As IEnumerator
    Set enum_obj = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS)

    Dim col_idx As Long
    col_idx = 4

    ' Act
    Do While enum_obj.MoveNext
        Dim col_bds As WorksheetRangeBounds
        Set col_bds = enum_obj.Current

        ' Check
        Assert.EqualsNumeric 10, col_bds.Row
        Assert.EqualsNumeric 11, col_bds.FinishRow
        Assert.EqualsNumeric col_idx, col_bds.Column
        Assert.EqualsNumeric col_idx, col_bds.FinishColumn

        col_idx = col_idx + 1
    Loop

    ' 3 columns => after loop col_idx=7
    Assert.EqualsNumeric 7, col_idx
End Sub

Public Sub Test_GetEnumerator_Cells_HorizontalDirection(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' 2x2 => row=2..3, col=5..6 => total4cells
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    Dim enum_obj As IEnumerator
    ' "Cells" + ColumnDirection=False => horizontal => (2,5)->(2,6)->(3,5)->(3,6)
    Set enum_obj = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL)

    ' Act
    Dim idx As Long
    idx = 0
    Dim expected_row() As Variant
    expected_row = Array(2, 2, 3, 3)
    Dim expected_col() As Variant
    expected_col = Array(5, 6, 5, 6)

    Do While enum_obj.MoveNext
        Dim cell_bds As WorksheetRangeBounds
        Set cell_bds = enum_obj.Current

        ' Assert inside loop
        Assert.EqualsNumeric expected_row(idx), cell_bds.Row
        Assert.EqualsNumeric expected_col(idx), cell_bds.Column
        Assert.IsTrue cell_bds.IsCell

        idx = idx + 1
    Loop

    Assert.EqualsNumeric 4, idx ' total4cells
End Sub

Public Sub Test_GetEnumerator_Cells_VerticalDirection(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' 2x2 => row=2..3, col=5..6 => total4cells
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    Dim enum_obj As IEnumerator
    ' "Cells" + ColumnDirection:=True => vertical => (2,5)->(3,5)->(2,6)->(3,6)
    Set enum_obj = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_VERTICAL)

    ' Act
    Dim idx As Long
    idx = 0
    Dim expected_row() As Variant
    expected_row = Array(2, 3, 2, 3)
    Dim expected_col() As Variant
    expected_col = Array(5, 5, 6, 6)

    Do While enum_obj.MoveNext
        Dim cell_bds As WorksheetRangeBounds
        Set cell_bds = enum_obj.Current

        ' Assert
        Assert.EqualsNumeric expected_row(idx), cell_bds.Row
        Assert.EqualsNumeric expected_col(idx), cell_bds.Column
        Assert.IsTrue cell_bds.IsCell

        idx = idx + 1
    Loop

    Assert.EqualsNumeric 4, idx
End Sub

Public Sub Test_ForEach_CellsHorizontal_EnumeratesSnapshot(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    ' Act
    Dim actual_value As String
    actual_value = ""
    Dim idx As Long
    idx = 0

    Dim cell_bds As WorksheetRangeBounds
    For Each cell_bds In rng_bds
        If actual_value <> "" Then actual_value = actual_value & ";"
        actual_value = actual_value & CStr(cell_bds.Row) & "," & CStr(cell_bds.Column)
        idx = idx + 1
    Next cell_bds

    ' Assert
    Assert.Equals "2,5;2,6;3,5;3,6", actual_value
    Assert.EqualsNumeric 4, idx
End Sub

Public Sub Test_ForEach_GetEnumeratorModeDoesNotAffectForEach(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    Dim explicit_value As String
    explicit_value = ""

    Dim enum_obj As IEnumerator
    Set enum_obj = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS)
    Do While enum_obj.MoveNext()
        Dim part_bds As WorksheetRangeBounds
        Set part_bds = enum_obj.Current
        If explicit_value <> "" Then explicit_value = explicit_value & ";"
        explicit_value = explicit_value & CStr(part_bds.Row) & ":" & CStr(part_bds.FinishRow) & "," & CStr(part_bds.Column) & ":" & CStr(part_bds.FinishColumn)
    Loop

    ' Act
    Dim for_each_value As String
    for_each_value = ""
    For Each part_bds In rng_bds
        If for_each_value <> "" Then for_each_value = for_each_value & ";"
        for_each_value = for_each_value & CStr(part_bds.Row) & "," & CStr(part_bds.Column)
    Next part_bds

    ' Assert
    Assert.Equals "2:3,5:5;2:3,6:6", explicit_value
    Assert.Equals "2,5;2,6;3,5;3,6", for_each_value
End Sub

' -----------------------------------------------------------------------------
' IsEntireSheet / IsEntireRow / IsEntireColumn
' -----------------------------------------------------------------------------

Public Sub Test_IsEntireSheet_WhenSheetFullyCovers_ReturnsTrue(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' Rows 1..G_ROW_MAX, columns 1..G_COL_MAX => entire sheet.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=G_COL_MAX, _
                                  Sheet:="MySheet", Book:="MyBook.xlsm")

    ' Act
    Dim actual As Boolean
    actual = rng_bds.IsEntireSheet

    ' Assert
    Assert.IsTrue actual
End Sub

Public Sub Test_IsEntireSheet_WhenRowPartOnlyButNotColumn_ReturnsFalse(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' Rows 1..G_ROW_MAX, but col=2..G_COL_MAX => all rows, but missing one column.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=G_ROW_MAX, FinishColumn:=G_COL_MAX)

    ' Act
    Dim actual As Boolean
    actual = rng_bds.IsEntireSheet

    ' Assert
    Assert.IsFalse actual
End Sub

Public Sub Test_IsEntireRow_WhenEmpty_ReturnsFalse(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' Row=2, col=2, finishRow=0 => empty => IsEntireRow=False.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=0, FinishColumn:=0)

    ' Act
    Dim actual As Boolean
    actual = rng_bds.IsEntireRow

    ' Assert
    Assert.IsFalse actual
End Sub

Public Sub Test_IsEntireRow_WhenPartOfRowOnly_ReturnsFalse(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' Row=5..5 => one row, but col=2..4 => not an "entire row"; entire row = col=1..G_COL_MAX.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=5, FinishColumn:=4)

    ' Act
    Dim actual As Boolean
    actual = rng_bds.IsEntireRow

    ' Assert
    Assert.IsFalse actual
End Sub

Public Sub Test_IsEntireColumn_WhenFinishColumnIsMaxButNotRow1_ReturnsFalse(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' col=1..G_COL_MAX, but row=2..G_ROW_MAX => not entire column.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=G_COL_MAX)

    ' Act
    Dim actual As Boolean
    actual = rng_bds.IsEntireColumn

    ' Assert
    Assert.IsFalse actual
End Sub


' -----------------------------------------------------------------------------
' GetRow / GetColumn / GetCell
' -----------------------------------------------------------------------------

Public Sub Test_GetRow_NegativeIndex_ReturnsRowFromBottom(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_last_row As WorksheetRangeBounds
    Set actual_last_row = rng_bds.GetRow(-1)

    Dim actual_second_last_row As WorksheetRangeBounds
    Set actual_second_last_row = rng_bds.GetRow(-2)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 5, actual_last_row.Row
    Assert.EqualsNumeric 5, actual_last_row.FinishRow
    Assert.EqualsNumeric 3, actual_last_row.Column
    Assert.EqualsNumeric 7, actual_last_row.FinishColumn
    Assert.EqualsNumeric 4, actual_second_last_row.Row
    Assert.EqualsNumeric 4, actual_second_last_row.FinishRow
End Sub

Public Sub Test_GetColumn_NegativeIndex_ReturnsColumnFromRight(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_last_column As WorksheetRangeBounds
    Set actual_last_column = rng_bds.GetColumn(-1)

    Dim actual_second_last_column As WorksheetRangeBounds
    Set actual_second_last_column = rng_bds.GetColumn(-2)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 7, actual_last_column.Column
    Assert.EqualsNumeric 7, actual_last_column.FinishColumn
    Assert.EqualsNumeric 2, actual_last_column.Row
    Assert.EqualsNumeric 5, actual_last_column.FinishRow
    Assert.EqualsNumeric 6, actual_second_last_column.Column
    Assert.EqualsNumeric 6, actual_second_last_column.FinishColumn
End Sub

Public Sub Test_GetCell_NegativeIndex_ReturnsCellFromBottomRight(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_bottom_right As WorksheetRangeBounds
    Set actual_bottom_right = rng_bds.GetCell(-1, -1)

    Dim actual_offset_cell As WorksheetRangeBounds
    Set actual_offset_cell = rng_bds.GetCell(-2, -3)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 5, actual_bottom_right.Row
    Assert.EqualsNumeric 7, actual_bottom_right.Column
    Assert.EqualsNumeric 5, actual_bottom_right.FinishRow
    Assert.EqualsNumeric 7, actual_bottom_right.FinishColumn
    Assert.EqualsNumeric 4, actual_offset_cell.Row
    Assert.EqualsNumeric 5, actual_offset_cell.Column
End Sub

Public Sub Test_GetRow_ZeroIndex_RaisesRangeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_row As WorksheetRangeBounds
    Set actual_row = rng_bds.GetRow(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The row index is outside the row range. (index:0, start: 2, finish: 5)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetColumn_ZeroIndex_RaisesRangeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_column As WorksheetRangeBounds
    Set actual_column = rng_bds.GetColumn(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The column index is outside the column range. (index:0, start: 3, finish: 7)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetCell_ZeroRowIndex_RaisesRangeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.GetCell(0, -1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The row index is outside the row range. (index:0, start: 2, finish: 5)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetCell_ZeroColumnIndex_RaisesRangeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.GetCell(-1, 0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The column index is outside the column range. (index:0, start: 3, finish: 7)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetRow_TooSmallNegativeIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_row As WorksheetRangeBounds
    Set actual_row = rng_bds.GetRow(-2147483648#)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The row index is outside the row range. (index:-2147483648, start: 2, finish: 5)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetColumn_TooSmallNegativeIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_column As WorksheetRangeBounds
    Set actual_column = rng_bds.GetColumn(-2147483648#)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The column index is outside the column range. (index:-2147483648, start: 3, finish: 7)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetRow_TooLargeIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_row As WorksheetRangeBounds
    Set actual_row = rng_bds.GetRow(2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The row index is outside the row range. (index:2147483647, start: 2, finish: 5)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetColumn_TooLargeIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_column As WorksheetRangeBounds
    Set actual_column = rng_bds.GetColumn(2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The column index is outside the column range. (index:2147483647, start: 3, finish: 7)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetCell_TooLargeRowIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.GetCell(2147483647, 1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The row index is outside the row range. (index:2147483647, start: 2, finish: 5)", Err.Description
    Err.Clear
End Sub

Public Sub Test_GetCell_TooLargeColumnIndex_RaisesRangeErrorBeforeOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=5, FinishColumn:=7)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.GetCell(1, 2147483647)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "The column index is outside the column range. (index:2147483647, start: 3, finish: 7)", Err.Description
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' Item
' -----------------------------------------------------------------------------

Public Sub Test_Item_NegativeIndexHorizontal_RaisesIndexOutOfRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.Item(-1, ColumnDirection:=False)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "Index is out of range.", Err.Description
    Err.Clear
End Sub

Public Sub Test_Item_NegativeIndexVertical_RaisesIndexOutOfRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=6)

    ' Act
    Dim actual_cell As WorksheetRangeBounds
    Set actual_cell = rng_bds.Item(-1, ColumnDirection:=True)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBounds.")
    Assert.Equals "Index is out of range.", Err.Description
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' RowCount / ColumnCount / Count
' -----------------------------------------------------------------------------

Public Sub Test_RowCount_EmptyDueToFinishRow0_Returns0(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' row=2, col=3 => finishRow=0 => empty
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=0, FinishColumn:=0)

    ' Act
    Dim actual_rc As Long
    actual_rc = rng_bds.RowCount

    ' Assert
    Assert.EqualsNumeric 0, actual_rc
End Sub

Public Sub Test_ColumnCount_EmptyDueToFinishColumn0_Returns0(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=0)

    ' Act
    Dim actual_cc As Long
    actual_cc = rng_bds.ColumnCount

    ' Assert
    Assert.EqualsNumeric 0, actual_cc
End Sub

Public Sub Test_Count_WhenEmpty_Returns0(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' row=10,col=10,finish=0 => empty
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=10, Column:=10, FinishRow:=0, FinishColumn:=0)

    ' Act
    Dim actual_c As Long
    actual_c = rng_bds.Count

    ' Assert
    Assert.EqualsNumeric 0, actual_c
End Sub

Public Sub Test_Count_When3x4Cells_Returns12(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    ' row=2..4 => 3 rows, col=3..6 => 4 columns => cell count=3*4=12.
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=6)

    ' Act
    Dim actual_c As Long
    actual_c = rng_bds.Count

    ' Assert
    Assert.EqualsNumeric 12, actual_c
End Sub

' -----------------------------------------------------------------------------
' New_RangeBoundsFromAddress / shape properties
' -----------------------------------------------------------------------------

Public Sub Test_New_RangeBoundsFromAddress_CellAddress_InitializesIndexes(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress("'[Book.xlsm]Data Sheet'!$B$2:$C$3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, range_bounds.Row
    Assert.EqualsNumeric 2, range_bounds.Column
    Assert.EqualsNumeric 3, range_bounds.FinishRow
    Assert.EqualsNumeric 3, range_bounds.FinishColumn
    Assert.Equals "Book.xlsm", range_bounds.WorkbookName
    Assert.Equals "Data Sheet", range_bounds.WorksheetName
End Sub

Public Sub Test_New_RangeBoundsFromAddress_RowRange_InitializesEntireRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress("1:3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue range_bounds.IsEntireRow
    Assert.EqualsNumeric 1, range_bounds.Row
    Assert.EqualsNumeric 1, range_bounds.Column
    Assert.EqualsNumeric 3, range_bounds.FinishRow
    Assert.EqualsNumeric G_COL_MAX, range_bounds.FinishColumn
End Sub

Public Sub Test_New_RangeBoundsFromAddress_ColumnRange_InitializesEntireColumn(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress("A:C")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue range_bounds.IsEntireColumn
    Assert.EqualsNumeric 1, range_bounds.Row
    Assert.EqualsNumeric 1, range_bounds.Column
    Assert.EqualsNumeric G_ROW_MAX, range_bounds.FinishRow
    Assert.EqualsNumeric 3, range_bounds.FinishColumn
End Sub

Public Sub Test_New_RangeBoundsFromAddress_MultiRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress("A1,B2")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsArea_WhenEmpty_ReturnsFalse(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=0, FinishColumn:=0)

    ' Assert
    Assert.IsTrue range_bounds.IsEmpty
    Assert.IsFalse range_bounds.IsCell
    Assert.IsFalse range_bounds.IsArea
End Sub

Public Sub Test_ShapeProperties_Cell_ReturnExpectedValues(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1)

    ' Assert
    Assert.IsTrue range_bounds.IsCell
    Assert.IsFalse range_bounds.IsArea
    Assert.IsTrue range_bounds.IsOneRow
    Assert.IsTrue range_bounds.IsOneColumn
    Assert.IsFalse range_bounds.IsOneRowArea
    Assert.IsFalse range_bounds.IsOneColumnArea
End Sub

Public Sub Test_ShapeProperties_OneRowArea_ReturnExpectedValues(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=1, FinishColumn:=2)

    ' Assert
    Assert.IsFalse range_bounds.IsCell
    Assert.IsTrue range_bounds.IsArea
    Assert.IsTrue range_bounds.IsOneRow
    Assert.IsFalse range_bounds.IsOneColumn
    Assert.IsTrue range_bounds.IsOneRowArea
    Assert.IsFalse range_bounds.IsOneColumnArea
End Sub

Public Sub Test_ShapeProperties_OneColumnArea_ReturnExpectedValues(ByVal Assert As UnitTestAssert)
    Call InitializeCommonService
    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=1)

    ' Assert
    Assert.IsFalse range_bounds.IsCell
    Assert.IsTrue range_bounds.IsArea
    Assert.IsFalse range_bounds.IsOneRow
    Assert.IsTrue range_bounds.IsOneColumn
    Assert.IsFalse range_bounds.IsOneRowArea
    Assert.IsTrue range_bounds.IsOneColumnArea
End Sub

Public Sub Test_Initialize_BookOmittedWithUninitializedWbSrv_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New WorksheetRangeBounds

    ' Act
    Call range_bounds.Initialize(Row:=1, Column:=1, Sheet:="Sheet1")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class WorksheetRangeBounds.pGetDefaultWorkbookName", actual_err_source
    Assert.Equals "InitializeCommonService has not been called.", actual_err_desc
End Sub

Public Sub Test_Initialize_BookExplicitWithUninitializedWbSrv_CanInitialize(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New WorksheetRangeBounds

    ' Act
    Call range_bounds.Initialize(Row:=1, Column:=1, Sheet:="Sheet1", Book:="Book1.xlsm")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Book1.xlsm", range_bounds.WorkbookName
End Sub
