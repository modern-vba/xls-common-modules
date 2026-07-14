Attribute VB_Name = "Test_UserInputSheet"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the UserInputSheet class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_Initialize_WithInputArea_ReturnsInputAreaCopy(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Row:=2, Column:=3, FinishRow:=10, FinishColumn:=5, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet

    ' Act
    Call input_sheet.Initialize(expected_area)

    Dim actual_area As WorksheetRangeBounds
    Dim second_area As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_area = input_sheet.InputArea
        Set second_area = input_sheet.InputArea
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
    Assert.NotEquals CStr(ObjPtr(expected_area)), CStr(ObjPtr(actual_area))
    Assert.NotEquals CStr(ObjPtr(actual_area)), CStr(ObjPtr(second_area))
End Sub

Public Sub Test_Initialize_InputAreaIsNothing_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet

    ' Act
    Call input_sheet.Initialize(Nothing)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class UserInputSheet.")
    Assert.Equals "InputArea is not specified.", Err.Description
End Sub

Public Sub Test_Initialize_InputAreaIsEmptyRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim empty_area As WorksheetRangeBounds
    Set empty_area = New_RangeBounds(Row:=2, Column:=3, FinishRow:=0, FinishColumn:=0, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet

    ' Act
    Call input_sheet.Initialize(empty_area)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class UserInputSheet.")
    Assert.Equals "InputArea is an empty range.", Err.Description
End Sub
Public Sub Test_GetItemRange_InputAreaOtherThanTopLeft_ReturnsValueRangeByAbsolutePosition(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim sheet_name As String
    sheet_name = "tmp_test_user_input"

    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTemporarySheet(sheet_name)

    target_sheet.Cells(2, 4).Value = "Target"
    target_sheet.Cells(6, 2).Value = "Target"
    target_sheet.Cells(11, 7).Value = "Target"
    target_sheet.Cells(6, 4).Value = "Target"
    target_sheet.Cells(6, 5).Value = "Expected1"
    target_sheet.Cells(6, 6).Value = "Expected2"

    Dim input_area As WorksheetRangeBounds
    Set input_area = New_RangeBounds(Row:=5, Column:=4, FinishRow:=10, FinishColumn:=6, Sheet:=sheet_name, Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet
    Call input_sheet.Initialize(input_area)

    Err.Clear

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_sheet.GetItemRange("Target")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Dim actual_is_nothing As Boolean
    Dim actual_row As Long
    Dim actual_column As Long
    Dim actual_finish_row As Long
    Dim actual_finish_column As Long
    actual_is_nothing = (actual_range Is Nothing)
    If Not actual_is_nothing Then
        actual_row = actual_range.Row
        actual_column = actual_range.Column
        actual_finish_row = actual_range.FinishRow
        actual_finish_column = actual_range.FinishColumn
    End If

    Call pDeleteTemporarySheet(sheet_name)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.IsFalse actual_is_nothing
    Assert.EqualsNumeric 6, actual_row
    Assert.EqualsNumeric 5, actual_column
    Assert.EqualsNumeric 6, actual_finish_row
    Assert.EqualsNumeric 6, actual_finish_column
End Sub

Public Sub Test_GetItemRange_WithSecondItem_ReturnsColumnValueRangeByAbsolutePosition(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim sheet_name As String
    sheet_name = "tmp_test_user_input_col"

    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTemporarySheet(sheet_name)

    target_sheet.Cells(3, 4).Value = "Section"
    target_sheet.Cells(6, 2).Value = "Section"
    target_sheet.Cells(11, 7).Value = "Section"
    target_sheet.Cells(6, 4).Value = "Section"
    target_sheet.Cells(6, 5).Value = "Name"
    target_sheet.Cells(6, 6).Value = "Age"
    target_sheet.Cells(7, 5).Value = "Alice"
    target_sheet.Cells(7, 6).Value = "31"
    target_sheet.Cells(8, 5).Value = "Bob"
    target_sheet.Cells(8, 6).Value = "42"
    target_sheet.Cells(9, 4).Value = "Next"
    target_sheet.Cells(9, 6).Value = "Ignored"
    target_sheet.Cells(6, 7).Value = "Age"
    target_sheet.Cells(7, 7).Value = "999"

    Dim input_area As WorksheetRangeBounds
    Set input_area = New_RangeBounds(Row:=5, Column:=4, FinishRow:=10, FinishColumn:=6, Sheet:=sheet_name, Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet
    Call input_sheet.Initialize(input_area)

    Err.Clear

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_sheet.GetItemRange("Section", "Age")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Dim actual_is_nothing As Boolean
    Dim actual_row As Long
    Dim actual_column As Long
    Dim actual_finish_row As Long
    Dim actual_finish_column As Long
    actual_is_nothing = (actual_range Is Nothing)
    If Not actual_is_nothing Then
        actual_row = actual_range.Row
        actual_column = actual_range.Column
        actual_finish_row = actual_range.FinishRow
        actual_finish_column = actual_range.FinishColumn
    End If

    Call pDeleteTemporarySheet(sheet_name)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.IsFalse actual_is_nothing
    Assert.EqualsNumeric 7, actual_row
    Assert.EqualsNumeric 6, actual_column
    Assert.EqualsNumeric 8, actual_finish_row
    Assert.EqualsNumeric 6, actual_finish_column
End Sub

Public Sub Test_GetItemRange_FirstItemOnlyOutsideArea_ReturnsNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim sheet_name As String
    sheet_name = "tmp_test_user_input_miss1"

    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTemporarySheet(sheet_name)

    target_sheet.Cells(3, 4).Value = "Target"
    target_sheet.Cells(6, 2).Value = "Target"
    target_sheet.Cells(11, 7).Value = "Target"

    Dim input_area As WorksheetRangeBounds
    Set input_area = New_RangeBounds(Row:=5, Column:=4, FinishRow:=10, FinishColumn:=6, Sheet:=sheet_name, Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet
    Call input_sheet.Initialize(input_area)

    Err.Clear

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_sheet.GetItemRange("Target")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Dim actual_is_nothing As Boolean
    actual_is_nothing = (actual_range Is Nothing)

    Call pDeleteTemporarySheet(sheet_name)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.IsTrue actual_is_nothing
End Sub

Public Sub Test_GetItemRange_SecondItemOnlyOutsideArea_ReturnsNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim sheet_name As String
    sheet_name = "tmp_test_user_input_miss2"

    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTemporarySheet(sheet_name)

    target_sheet.Cells(6, 4).Value = "Section"
    target_sheet.Cells(6, 5).Value = "Name"
    target_sheet.Cells(6, 7).Value = "Age"
    target_sheet.Cells(7, 7).Value = "999"

    Dim input_area As WorksheetRangeBounds
    Set input_area = New_RangeBounds(Row:=5, Column:=4, FinishRow:=10, FinishColumn:=6, Sheet:=sheet_name, Book:=ThisWorkbook.Name)

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet
    Call input_sheet.Initialize(input_area)

    Err.Clear

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_sheet.GetItemRange("Section", "Age")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Dim actual_is_nothing As Boolean
    actual_is_nothing = (actual_range Is Nothing)

    Call pDeleteTemporarySheet(sheet_name)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.IsTrue actual_is_nothing
End Sub
Private Function pPrepareTemporarySheet(ByVal SheetName As String) As Worksheet
    Call pDeleteTemporarySheet(SheetName)
    Set pPrepareTemporarySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    pPrepareTemporarySheet.Name = SheetName
End Function

Private Sub pDeleteTemporarySheet(ByVal SheetName As String)
    Dim display_alerts As Boolean
    display_alerts = Application.DisplayAlerts

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = display_alerts
End Sub
