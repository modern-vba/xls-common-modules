Attribute VB_Name = "Test_UserInputSheetTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the UserInputSheetTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_InputArea_Unregistered_ReturnsDefaultArea(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_stub As UserInputSheetTestDouble
    Set input_stub = New UserInputSheetTestDouble

    Dim input_sheet As IUserInputSheet
    Set input_sheet = input_stub

    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Sheet:="InputSheet", Book:="ThisWorkbook.xlsm")

    ' Act
    Dim actual_area As WorksheetRangeBounds
    Set actual_area = input_sheet.InputArea

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
End Sub

Public Sub Test_InputArea_WithRegisteredReturn_ReturnsRegisteredArea(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_stub As UserInputSheetTestDouble
    Set input_stub = New UserInputSheetTestDouble

    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Row:=2, Column:=3, FinishRow:=9, FinishColumn:=5, Sheet:="InputSheet", Book:="Book.xlsm")
    Call input_stub.Store.SetReturn("InputArea", expected_area)

    Dim input_sheet As IUserInputSheet
    Set input_sheet = input_stub

    ' Act
    Dim actual_area As WorksheetRangeBounds
    Set actual_area = input_sheet.InputArea

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
End Sub

Public Sub Test_GetItemRange_WithRegisteredReturn_ReturnsRangeAndRecordsCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_stub As UserInputSheetTestDouble
    Set input_stub = New UserInputSheetTestDouble

    Dim expected_range As WorksheetRangeBounds
    Set expected_range = New_RangeBounds(Row:=2, Column:=3, Sheet:="InputSheet", Book:="Book.xlsm")
    Call input_stub.Store.SetReturn("GetItemRange", expected_range, "Condition", "Address")

    Dim input_sheet As IUserInputSheet
    Set input_sheet = input_stub

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_sheet.GetItemRange("Condition", "Address")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_range, actual_range
    Assert.EqualsNumeric 1, input_stub.Store.GetCallCount("GetItemRange", "Condition", "Address")
End Sub

Public Sub Test_GetItemRange_Unregistered_RaisesStoreErrorAndNoCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_stub As UserInputSheetTestDouble
    Set input_stub = New UserInputSheetTestDouble

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_stub.GetItemRange("Condition", "Address")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, input_stub.Store.GetCallCount("GetItemRange", "Condition", "Address")
End Sub

Public Sub Test_GetItemRange_WithSetError_RaisesErrorAndNoCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim input_stub As UserInputSheetTestDouble
    Set input_stub = New UserInputSheetTestDouble
    Call input_stub.Store.SetError("GetItemRange", vbObjectError + 2049, "Class Injected", "Injected error.", "Condition", "Service")

    ' Act
    Dim actual_range As WorksheetRangeBounds
    Set actual_range = input_stub.GetItemRange("Condition", "Service")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 2049, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class Injected", Err.Source
    Assert.Equals "Injected error.", Err.Description
    Assert.EqualsNumeric 0, input_stub.Store.GetCallCount("GetItemRange", "Condition", "Service")
End Sub
