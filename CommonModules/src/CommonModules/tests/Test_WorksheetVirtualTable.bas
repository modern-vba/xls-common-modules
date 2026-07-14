Attribute VB_Name = "Test_WorksheetVirtualTable"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the WorksheetVirtualTable class.
'!
' #############################################################################

Public Sub Test_NewWorksheetVirtualTable_HeaderNames_ReturnsRowDictionary(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Name"
    header_names(1) = "Status"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim first_row As ObjectDictionary
    Dim name_bounds As WorksheetRangeBounds
    Dim status_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.Item(0)
        Set name_bounds = first_row.Item("Name")
        Set status_bounds = first_row.Item("Status")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, table.Count
    Assert.EqualsNumeric 2, table.RowCount
    Assert.EqualsNumeric 2, first_row.Count
    Assert.EqualsNumeric vbBinaryCompare, first_row.CompareMode
    Assert.IsTrue name_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue status_bounds.Equals(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_NewWorksheetVirtualTableFromRangeBounds_HeaderNames_ReturnsRowDictionary(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim table_range As WorksheetRangeBounds
    Set table_range = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name)

    Dim header_names(0 To 2) As String
    header_names(0) = "Name"
    header_names(1) = "Status"
    header_names(2) = "Memo"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTableFromRangeBounds(table_range, header_names)

    Dim first_row As ObjectDictionary
    Dim name_bounds As WorksheetRangeBounds
    Dim status_bounds As WorksheetRangeBounds
    Dim memo_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.Item(0)
        Set name_bounds = first_row.Item("Name")
        Set status_bounds = first_row.Item("Status")
        Set memo_bounds = first_row.Item("Memo")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, table.Count
    Assert.EqualsNumeric 3, table.RowCount
    Assert.EqualsNumeric 3, first_row.Count
    Assert.EqualsNumeric vbBinaryCompare, first_row.CompareMode
    Assert.IsTrue name_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue status_bounds.Equals(New_RangeBounds(Row:=2, Column:=3, FinishRow:=2, FinishColumn:=3, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue memo_bounds.Equals(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_GetRow_OneBasedAndNegativeIndex_ReturnsRowDictionary(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Name"
    header_names(1) = "Status"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    ' Act
    Dim first_row As ObjectDictionary
    Dim last_row As ObjectDictionary
    Dim first_name_bounds As WorksheetRangeBounds
    Dim last_status_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.GetRow(1)
        Set last_row = table.GetRow(-1)
        Set first_name_bounds = first_row.Item("Name")
        Set last_status_bounds = last_row.Item("Status")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue first_name_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue last_status_bounds.Equals(New_RangeBounds(Row:=3, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_NewWorksheetVirtualTable_TreatFirstRowAsHeader_JoinsMultiColumnHeader(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = ThisWorkbook.Worksheets("test_input")
    Call target_sheet.Range("J40:M42").ClearContents
    target_sheet.Cells(40, 10).Value = "First"
    target_sheet.Cells(40, 11).Value = "Last"
    target_sheet.Cells(40, 13).Value = "Status"

    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=40, Column:=10, FinishRow:=42, FinishColumn:=11, Sheet:="test_input", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=40, Column:=13, FinishRow:=42, FinishColumn:=13, Sheet:="test_input", Book:=ThisWorkbook.Name))

    Dim ignored_header_names(0 To 1) As String
    ignored_header_names(0) = "IgnoredName"
    ignored_header_names(1) = "IgnoredStatus"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, ignored_header_names, TreatFirstRowAsHeader:=True, HeaderJoinDelimiter:=" ")

    Dim headers() As String
    Dim first_row As ObjectDictionary
    Dim name_bounds As WorksheetRangeBounds
    Dim status_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        headers = table.Headers
        Set first_row = table.Item(0)
        Set name_bounds = first_row.Item("First Last")
        Set status_bounds = first_row.Item("Status")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, table.RowCount
    Assert.Equals "First Last", headers(0)
    Assert.Equals "Status", headers(1)
    Assert.IsTrue name_bounds.Equals(New_RangeBounds(Row:=41, Column:=10, FinishRow:=41, FinishColumn:=11, Sheet:="test_input", Book:=ThisWorkbook.Name))
    Assert.IsTrue status_bounds.Equals(New_RangeBounds(Row:=41, Column:=13, FinishRow:=41, FinishColumn:=13, Sheet:="test_input", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_Item_ShortColumn_ReturnsEmptyRangePreservingColumnWidth(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=5, FinishRow:=2, FinishColumn:=6, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Full"
    header_names(1) = "Short"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    ' Act
    Dim second_row As ObjectDictionary
    Dim full_bounds As WorksheetRangeBounds
    Dim short_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set second_row = table.Item(1)
        Set full_bounds = second_row.Item("Full")
        Set short_bounds = second_row.Item("Short")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, table.RowCount
    Assert.IsTrue full_bounds.Equals(New_RangeBounds(Row:=3, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue short_bounds.IsEmpty
    Assert.EqualsNumeric 3, short_bounds.Row
    Assert.EqualsNumeric 5, short_bounds.Column
    Assert.EqualsNumeric 0, short_bounds.FinishRow
    Assert.EqualsNumeric 6, short_bounds.FinishColumn
    Assert.EqualsNumeric 0, short_bounds.RowCount
    Assert.EqualsNumeric 2, short_bounds.ColumnCount
End Sub

Public Sub Test_GetEnumeratorAndForEach_EnumeratesRowsInOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Name"
    header_names(1) = "Status"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    ' Act
    Dim row_enum As IEnumerator
    Dim enum_first_row As ObjectDictionary
    If Err.Number = 0 Then
        Set row_enum = table.GetEnumerator()
        If row_enum.MoveNext() Then Set enum_first_row = row_enum.Current
    End If

    Dim each_count As Long
    Dim each_last_row As ObjectDictionary
    Dim row_item As Variant
    If Err.Number = 0 Then
        For Each row_item In table
            Set each_last_row = row_item
            each_count = each_count + 1
        Next row_item
    End If

    Dim enum_name_bounds As WorksheetRangeBounds
    Dim each_status_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set enum_name_bounds = enum_first_row.Item("Name")
        Set each_status_bounds = each_last_row.Item("Status")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, row_enum.Index
    Assert.EqualsNumeric 2, each_count
    Assert.IsTrue enum_name_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue each_status_bounds.Equals(New_RangeBounds(Row:=3, Column:=4, FinishRow:=3, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_NewWorksheetVirtualTable_TextCompareDuplicateHeaders_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Name"
    header_names(1) = "name"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names, HeaderCompareMode:=vbTextCompare)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_VariantArrayHeaderNames_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names As Variant
    header_names = Array("Name", "Status")

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_EmptyTypedRangeListAndEmptyHeaders_ReturnsEmptyTable(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")

    Dim header_names() As String
    header_names = EmptyStringArray()

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim headers() As String
    If Err.Number = 0 Then headers = table.Headers

    Dim each_count As Long
    Dim row_item As Variant
    If Err.Number = 0 Then
        For Each row_item In table
            each_count = each_count + 1
        Next row_item
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, table.Count
    Assert.EqualsNumeric 0, table.RowCount
    Assert.IsTrue IsEmptyArray(headers)
    Assert.EqualsNumeric 0, each_count
End Sub

Public Sub Test_NewWorksheetVirtualTable_WithoutHeaderNames_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    Assert.ErrorRaised 0, actual_error_number, actual_error_source, actual_error_description
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_UntypedColumnRangeList_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New ObjectList

    Dim header_names() As String
    header_names = EmptyStringArray()

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_WrongColumnRangeListTypeContract_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("ObjectDictionary")

    Dim header_names() As String
    header_names = EmptyStringArray()

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_HeaderCountMismatch_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 0) As String
    header_names(0) = "Name"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_SingleEmptyHeader_ReturnsRowDictionary(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = ""
    header_names(1) = "Status"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim first_row As ObjectDictionary
    Dim blank_bounds As WorksheetRangeBounds
    Dim status_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.Item(0)
        Set blank_bounds = first_row.Item("")
        Set status_bounds = first_row.Item("Status")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, first_row.Count
    Assert.IsTrue blank_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue status_bounds.Equals(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_NewWorksheetVirtualTable_DuplicateEmptyHeaders_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = ""
    header_names(1) = ""

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_BinaryCompareCaseDifferentHeaders_ReturnsRowDictionary(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Name"
    header_names(1) = "name"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names, HeaderCompareMode:=vbBinaryCompare)

    Dim first_row As ObjectDictionary
    Dim name_bounds As WorksheetRangeBounds
    Dim lower_name_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.Item(0)
        Set name_bounds = first_row.Item("Name")
        Set lower_name_bounds = first_row.Item("name")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric vbBinaryCompare, first_row.CompareMode
    Assert.EqualsNumeric 2, first_row.Count
    Assert.IsTrue name_bounds.Equals(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Assert.IsTrue lower_name_bounds.Equals(New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
End Sub

Public Sub Test_Item_EmptyOptionalField_ReturnsEmptyRangeAtTargetRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=4, FinishRow:=1, FinishColumn:=4, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 1) As String
    header_names(0) = "Required"
    header_names(1) = "Optional"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    ' Act
    Dim first_row As ObjectDictionary
    Dim second_row As ObjectDictionary
    Dim first_optional_bounds As WorksheetRangeBounds
    Dim second_optional_bounds As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set first_row = table.GetRow(1)
        Set second_row = table.GetRow(2)
        Set first_optional_bounds = first_row.Item("Optional")
        Set second_optional_bounds = second_row.Item("Optional")
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, table.RowCount
    Assert.IsTrue first_optional_bounds.IsEmpty
    Assert.EqualsNumeric 2, first_optional_bounds.Row
    Assert.EqualsNumeric 4, first_optional_bounds.Column
    Assert.EqualsNumeric 0, first_optional_bounds.FinishRow
    Assert.EqualsNumeric 4, first_optional_bounds.FinishColumn
    Assert.EqualsNumeric 0, first_optional_bounds.RowCount
    Assert.EqualsNumeric 1, first_optional_bounds.ColumnCount
    Assert.IsTrue second_optional_bounds.IsEmpty
    Assert.EqualsNumeric 3, second_optional_bounds.Row
    Assert.EqualsNumeric 4, second_optional_bounds.Column
    Assert.EqualsNumeric 0, second_optional_bounds.FinishRow
    Assert.EqualsNumeric 4, second_optional_bounds.FinishColumn
End Sub

Public Sub Test_GetRow_ZeroIndex_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 0) As String
    header_names(0) = "Name"

    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    ' Act
    Dim actual_row As ObjectDictionary
    If Err.Number = 0 Then Set actual_row = table.GetRow(0)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub

Public Sub Test_NewWorksheetVirtualTable_MultiDimensionalStringHeaderNames_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim col_ranges As ObjectList
    Set col_ranges = New_ObjectList("WorksheetRangeBounds")
    Call col_ranges.Add(New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name))

    Dim header_names(0 To 0, 0 To 0) As String
    header_names(0, 0) = "Name"

    ' Act
    Dim table As WorksheetVirtualTable
    Set table = New_WorksheetVirtualTable(col_ranges, header_names)

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorksheetVirtualTable.")
End Sub
