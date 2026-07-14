Attribute VB_Name = "Test_Lib_InputSheet"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the Lib_InputSheet module.
'!
' #############################################################################

Public Sub Test_NewInputSheet_WithInputArea_ReturnsInitializedSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Row:=2, Column:=3, FinishRow:=10, FinishColumn:=5, Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name)

    ' Act
    Dim actual_value As UserInputSheet
    Set actual_value = New_InputSheet(expected_area)

    Dim actual_area As WorksheetRangeBounds
    If Err.Number = 0 Then Set actual_area = actual_value.InputArea

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
End Sub

Public Sub Test_NewInputSheet_WithInputArea_ReturnsInitializedInterface(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Sheet:="UNIT_TEST_SHEET", Book:=ThisWorkbook.Name)

    ' Act
    Dim actual_value As IUserInputSheet
    Set actual_value = New_InputSheet(expected_area)

    Dim actual_area As WorksheetRangeBounds
    If Err.Number = 0 Then Set actual_area = actual_value.InputArea

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
End Sub
