Attribute VB_Name = "Test_WorksheetRangeBoundsEnumer"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the WorksheetRangeBoundsEnumerator class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

' --------------------------------------------------------------------------
' Initialize method.
' --------------------------------------------------------------------------
Public Sub Test_Initialize_EnumerationModeOmitted_UsesRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator

    ' Act
    Call rng_enum.Initialize(TargetCollection:=rng_bds)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_RANGE_ENUM_MODE_ROWS, rng_enum.EnumerationMode
    Assert.EqualsNumeric -1, rng_enum.Index
End Sub

Public Sub Test_Initialize_RowsMode_ExposesBasicProperties(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator

    ' Act
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_RANGE_ENUM_MODE_ROWS, rng_enum.EnumerationMode
    Assert.IsFalse rng_enum.Descending
    Assert.EqualsNumeric -1, rng_enum.Index
End Sub

Public Sub Test_Initialize_CellsVerticalDescending_ExposesBasicProperties(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=3, FinishColumn:=4, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator

    ' Act
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_VERTICAL, Descending:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_RANGE_ENUM_MODE_CELLS_VERTICAL, rng_enum.EnumerationMode
    Assert.IsTrue rng_enum.Descending
    Assert.EqualsNumeric 4, rng_enum.Index
End Sub

Public Sub Test_Initialize_NothingRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim target_range As WorksheetRangeBounds

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator

    ' Act
    Call rng_enum.Initialize(target_range)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

Public Sub Test_GetEnumerator_DescendingRows_EnumeratesFromFinishRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    ' Act
    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator(Descending:=True)
    Dim moved As Boolean
    moved = rng_enum.MoveNext()
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue rng_enum.Descending
    Assert.EqualsNumeric 2, rng_enum.Index
    Assert.IsTrue moved
    Assert.EqualsNumeric 4, current_bds.Row
    Assert.EqualsNumeric 4, current_bds.FinishRow
    Assert.EqualsNumeric 3, current_bds.Column
    Assert.EqualsNumeric 5, current_bds.FinishColumn
End Sub

' --------------------------------------------------------------------------
' Target property.
' --------------------------------------------------------------------------
Public Sub Test_Target_Initialized_ReturnsSourceRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS)

    ' Act
    Dim actual_target As WorksheetRangeBounds
    Set actual_target = rng_enum.Target

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue (rng_bds Is actual_target)
End Sub

Public Sub Test_Target_BeforeInitialize_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator

    ' Act
    Dim actual_target As WorksheetRangeBounds
    Set actual_target = rng_enum.Target

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

' --------------------------------------------------------------------------
' Current property.
' --------------------------------------------------------------------------
Public Sub Test_Current_BeforeMoveNext_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()

    ' Act
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

Public Sub Test_Current_AfterMoveNextFalse_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()
    Do While rng_enum.MoveNext()
    Loop

    ' Act
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

Public Sub Test_Current_DescendingBeforeMoveNext_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=True)

    ' Act
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

Public Sub Test_Current_DescendingAfterMoveNextFalse_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=True)
    Do While rng_enum.MoveNext()
    Loop

    ' Act
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

' --------------------------------------------------------------------------
' MoveNext method.
' --------------------------------------------------------------------------
Public Sub Test_MoveNext_RowEmptyColumnsMode_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=4, FinishColumn:=6, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS)

    ' Act
    Dim is_succeed As Boolean
    is_succeed = rng_enum.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
End Sub

Public Sub Test_MoveNext_ColumnEmptyRowsMode_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=3, Column:=7, FinishRow:=8, FinishColumn:=6, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()

    ' Act
    Dim is_succeed As Boolean
    is_succeed = rng_enum.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
End Sub

Public Sub Test_MoveNext_RowsMode_EnumeratesRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(2, 3, 4, 5, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator()

    ' Act & Assert
    Dim is_succeed As Boolean
    Assert.EqualsNumeric -1, rng_enum.Index

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 0, rng_enum.Index
    Dim row_bds As WorksheetRangeBounds
    Set row_bds = rng_enum.Current
    Assert.EqualsNumeric 2, row_bds.Row
    Assert.EqualsNumeric 2, row_bds.FinishRow
    Assert.EqualsNumeric 3, row_bds.Column
    Assert.EqualsNumeric 5, row_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 1, rng_enum.Index
    Set row_bds = rng_enum.Current
    Assert.EqualsNumeric 3, row_bds.Row
    Assert.EqualsNumeric 3, row_bds.FinishRow
    Assert.EqualsNumeric 3, row_bds.Column
    Assert.EqualsNumeric 5, row_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 2, rng_enum.Index
    Set row_bds = rng_enum.Current
    Assert.EqualsNumeric 4, row_bds.Row
    Assert.EqualsNumeric 4, row_bds.FinishRow
    Assert.EqualsNumeric 3, row_bds.Column
    Assert.EqualsNumeric 5, row_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
    Assert.EqualsNumeric 3, rng_enum.Index
End Sub

Public Sub Test_MoveNext_ColumnsMode_EnumeratesColumns(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(2, 3, 4, 5, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS)

    ' Act & Assert
    Dim is_succeed As Boolean
    Assert.EqualsNumeric -1, rng_enum.Index

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 0, rng_enum.Index
    Dim col_bds As WorksheetRangeBounds
    Set col_bds = rng_enum.Current
    Assert.EqualsNumeric 2, col_bds.Row
    Assert.EqualsNumeric 4, col_bds.FinishRow
    Assert.EqualsNumeric 3, col_bds.Column
    Assert.EqualsNumeric 3, col_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 1, rng_enum.Index
    Set col_bds = rng_enum.Current
    Assert.EqualsNumeric 2, col_bds.Row
    Assert.EqualsNumeric 4, col_bds.FinishRow
    Assert.EqualsNumeric 4, col_bds.Column
    Assert.EqualsNumeric 4, col_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 2, rng_enum.Index
    Set col_bds = rng_enum.Current
    Assert.EqualsNumeric 2, col_bds.Row
    Assert.EqualsNumeric 4, col_bds.FinishRow
    Assert.EqualsNumeric 5, col_bds.Column
    Assert.EqualsNumeric 5, col_bds.FinishColumn

    is_succeed = rng_enum.MoveNext
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
    Assert.EqualsNumeric 3, rng_enum.Index
End Sub

Public Sub Test_MoveNext_CellsMode_Horizontal(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(2, 4, 3, 6, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL)

    ' Act & Assert
    Dim idx As Long
    Dim is_succeed As Boolean
    For idx = 0 To 5
        is_succeed = rng_enum.MoveNext
        Assert.IsTrue is_succeed
        Assert.EqualsNumeric idx, rng_enum.Index

        Dim cell_bds As WorksheetRangeBounds
        Set cell_bds = rng_enum.Current

        Dim expected_row As Long
        Dim expected_col As Long
        expected_row = 2 + (idx \ 3)
        expected_col = 4 + (idx Mod 3)

        Assert.EqualsNumeric expected_row, cell_bds.Row
        Assert.EqualsNumeric expected_col, cell_bds.Column
    Next idx

    is_succeed = rng_enum.MoveNext
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
    Assert.EqualsNumeric 6, rng_enum.Index
End Sub

Public Sub Test_MoveNext_CellsMode_Vertical(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(2, 4, 3, 6, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_VERTICAL)

    ' Act & Assert
    Dim idx As Long
    Dim is_succeed As Boolean
    For idx = 0 To 5
        is_succeed = rng_enum.MoveNext
        Assert.IsTrue is_succeed
        Assert.EqualsNumeric idx, rng_enum.Index

        Dim cell_bds As WorksheetRangeBounds
        Set cell_bds = rng_enum.Current

        Dim row_offset As Long
        Dim col_offset As Long
        row_offset = idx Mod rng_bds.RowCount
        col_offset = (idx - row_offset) / rng_bds.RowCount

        Assert.EqualsNumeric rng_bds.Row + row_offset, cell_bds.Row
        Assert.EqualsNumeric rng_bds.Column + col_offset, cell_bds.Column
    Next idx

    is_succeed = rng_enum.MoveNext
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse is_succeed
    Assert.EqualsNumeric 6, rng_enum.Index
End Sub

Public Sub Test_MoveNext_SingleCellRows_EnumeratesOnceThenEndIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=7, FinishRow:=5, FinishColumn:=7, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()

    ' Act
    Dim first_moved As Boolean
    first_moved = rng_enum.MoveNext()
    Dim first_index As Long
    first_index = rng_enum.Index
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current
    Dim second_moved As Boolean
    second_moved = rng_enum.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue first_moved
    Assert.EqualsNumeric 0, first_index
    Assert.EqualsNumeric 5, current_bds.Row
    Assert.EqualsNumeric 5, current_bds.FinishRow
    Assert.EqualsNumeric 7, current_bds.Column
    Assert.EqualsNumeric 7, current_bds.FinishColumn
    Assert.IsFalse second_moved
    Assert.EqualsNumeric 1, rng_enum.Index
End Sub

Public Sub Test_MoveNext_DescendingSingleCellRows_EnumeratesOnceThenBeforeStartIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=5, Column:=7, FinishRow:=5, FinishColumn:=7, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=True)

    ' Act
    Dim first_moved As Boolean
    first_moved = rng_enum.MoveNext()
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current
    Dim second_moved As Boolean
    second_moved = rng_enum.MoveNext()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue first_moved
    Assert.EqualsNumeric 5, current_bds.Row
    Assert.EqualsNumeric 5, current_bds.FinishRow
    Assert.EqualsNumeric 7, current_bds.Column
    Assert.EqualsNumeric 7, current_bds.FinishColumn
    Assert.IsFalse second_moved
    Assert.EqualsNumeric -1, rng_enum.Index
End Sub

' --------------------------------------------------------------------------
' Reset method.
' --------------------------------------------------------------------------
Public Sub Test_Reset_AfterSomeIteration_GoesBackToIndexMinusOne(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(10, 1, 12, 1, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator()
    Dim ignored As Boolean
    ignored = rng_enum.MoveNext
    ignored = rng_enum.MoveNext

    ' Act
    Call rng_enum.Reset
    Dim moved As Boolean
    moved = rng_enum.MoveNext

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue moved
    Assert.EqualsNumeric 0, rng_enum.Index
End Sub

Public Sub Test_Reset_DescendingRows_ReturnsToAfterLastIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=True)
    Dim ignored As Boolean
    ignored = rng_enum.MoveNext()
    ignored = rng_enum.MoveNext()

    ' Act
    Call rng_enum.Reset
    Dim moved As Boolean
    moved = rng_enum.MoveNext()
    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue moved
    Assert.EqualsNumeric 2, rng_enum.Index
    Assert.EqualsNumeric 4, current_bds.Row
    Assert.EqualsNumeric 4, current_bds.FinishRow
End Sub

' --------------------------------------------------------------------------
' SkipTo method.
' --------------------------------------------------------------------------
Public Sub Test_SkipTo_WithinRange_SetsIndexCorrectly(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(1, 1, 1, 5, "TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = rng_bds.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS)
    Dim is_succeed As Boolean
    is_succeed = rng_enum.MoveNext

    ' Act
    Call rng_enum.SkipTo(3)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, rng_enum.Index

    is_succeed = rng_enum.MoveNext
    Assert.IsTrue is_succeed
    Assert.EqualsNumeric 4, rng_enum.Index

    is_succeed = rng_enum.MoveNext
    Assert.IsFalse is_succeed
End Sub

Public Sub Test_SkipTo_Length_RaisesEnumeratorBoundaryError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()

    ' Act
    Call rng_enum.SkipTo(3)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

Public Sub Test_SkipTo_DescendingRows_AllowsForwardLowerIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=3, Sheet:="TestSheet")

    Dim rng_enum As WorksheetRangeBoundsEnumerator
    Set rng_enum = New WorksheetRangeBoundsEnumerator
    Call rng_enum.Initialize(TargetCollection:=rng_bds, EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=True)

    ' Act
    Call rng_enum.SkipTo(2)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, rng_enum.Index

    Dim current_bds As WorksheetRangeBounds
    Set current_bds = rng_enum.Current
    Assert.EqualsNumeric 4, current_bds.Row
    Assert.EqualsNumeric 4, current_bds.FinishRow
End Sub

' --------------------------------------------------------------------------
' Remove method.
' --------------------------------------------------------------------------
Public Sub Test_Remove_ReadOnly_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()
    Dim ignored As Boolean
    ignored = rng_enum.MoveNext()

    ' Act
    Dim removed_value As Variant
    removed_value = rng_enum.Remove()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub

' --------------------------------------------------------------------------
' Update method.
' --------------------------------------------------------------------------
Public Sub Test_Update_ReadOnly_RaisesEnumeratorError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim rng_bds As WorksheetRangeBounds
    Set rng_bds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="TestSheet")

    Dim rng_enum As IEnumerator
    Set rng_enum = rng_bds.GetEnumerator()
    Dim ignored As Boolean
    ignored = rng_enum.MoveNext()

    ' Act
    Call rng_enum.Update(New_RangeBounds(Row:=9, Column:=9, Sheet:="TestSheet"))

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class WorksheetRangeBoundsEnumerator.")
End Sub
