Attribute VB_Name = "Test_WorkbookServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the WorkbookServiceTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

' ----------------------------------------------------------------------------
' Tests for GetThisWorkbookName.
' ----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookName_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    Call wb_srv_stub.Store.SetReturn("GetThisWorkbookName", "Tool.xlsm")

    Dim wb_srv As IWorkbookService
    Set wb_srv = wb_srv_stub

    ' Act
    Dim actual_name As String
    actual_name = wb_srv.GetThisWorkbookName()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Tool.xlsm", actual_name
End Sub
' ----------------------------------------------------------------------------
' Tests for GetThisWorkbookDirectoryPath.
' ----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookDirectoryPath_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    Call wb_srv_stub.Store.SetReturn("GetThisWorkbookDirectoryPath", "C:\Temp")

    Dim wb_srv As IWorkbookService
    Set wb_srv = wb_srv_stub

    ' Act
    Dim actual_path As String
    actual_path = wb_srv.GetThisWorkbookDirectoryPath()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Temp", actual_path
End Sub

' ----------------------------------------------------------------------------
' Tests for WorkbookExists.
' ----------------------------------------------------------------------------

Public Sub Test_WorkbookExists_SetValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "TestBook.xlsm"

    Call wb_srv_stub.Store.SetReturn("WorkbookExists", True, wb_name)

    ' Act
    Dim actual_result As Boolean
    actual_result = wb_srv_stub.WorkbookExists(wb_name)

    ' Assert
    Assert.Equals True, actual_result
End Sub

Public Sub Test_WorkbookExists_SetValueButNotMatch_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble

    Call wb_srv_stub.Store.SetReturn("WorkbookExists", True, "TestBook.xlsm")

    ' Act
    Dim actual_result As Boolean
    actual_result = wb_srv_stub.WorkbookExists("other_name")

    ' Assert
    Assert.Equals False, actual_result
End Sub

Public Sub Test_WorkbookExists_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    ' Key not registered.

    ' Act
    Dim actual_result As Boolean
    actual_result = wb_double.WorkbookExists("UnknownBook.xlsx")

    ' Assert - False when no registration exists.
    Assert.IsFalse actual_result
End Sub

' ----------------------------------------------------------------------------
' Tests for GetAllWorkbooks.
' ----------------------------------------------------------------------------
Public Sub Test_GetAllWorkbooks_ReturnsSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim expected_arr(0 To 2) As String
    expected_arr(0) = "BookA"
    expected_arr(1) = "BookB"
    expected_arr(2) = "BookC"
    Call wb_double.Store.SetReturn("GetAllWorkbooks", expected_arr)

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetAllWorkbooks

    ' Assert
    Assert.EqualsArray expected_arr, actual_arr
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("GetAllWorkbooks")
End Sub

Public Sub Test_GetAllWorkbooks_Unregistered_ReturnsThisWorkbookNameArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim expected_arr(0 To 0) As String
    expected_arr(0) = "ThisWorkbook.xlsm"

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetAllWorkbooks

    ' Assert
    Assert.EqualsArray expected_arr, actual_arr
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("GetAllWorkbooks")
End Sub

Public Sub Test_GetAllWorkbooks_DefaultUsesConfiguredThisWorkbookNameWithoutRecordingInternalCall(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    Call wb_double.Store.SetReturn("GetThisWorkbookName", "Tool.xlsm")

    Dim expected_arr(0 To 0) As String
    expected_arr(0) = "Tool.xlsm"

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetAllWorkbooks

    ' Assert
    Assert.EqualsArray expected_arr, actual_arr
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("GetAllWorkbooks")
    Assert.EqualsNumeric 0, wb_double.Store.GetCallCount("GetThisWorkbookName")
End Sub

Public Sub Test_GetAllWorkbooks_DefaultPropagatesThisWorkbookNameErrorWithoutRecordingCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    Call wb_double.Store.SetError("GetThisWorkbookName", vbObjectError + 222, "Test Source", "name failed")

    Dim actual_arr() As String
    actual_arr = wb_double.GetAllWorkbooks

    If Not Assert.ErrorRaised(vbObjectError + 222, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Test Source", Err.Source
    Assert.EqualsNumeric 0, wb_double.Store.GetCallCount("GetAllWorkbooks")
    Assert.EqualsNumeric 0, wb_double.Store.GetCallCount("GetThisWorkbookName")
End Sub

Public Sub Test_GetOtherWorkbooks_ReturnsSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim expected_arr(0 To 1) As String
    expected_arr(0) = "BookA"
    expected_arr(1) = "BookB"
    Call wb_double.Store.SetReturn("GetOtherWorkbooks", expected_arr)

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetOtherWorkbooks

    ' Assert
    Assert.EqualsArray expected_arr, actual_arr
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("GetOtherWorkbooks")
End Sub

Public Sub Test_GetOtherWorkbooks_Unregistered_ReturnsEmptyArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetOtherWorkbooks

    ' Assert
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("GetOtherWorkbooks")
End Sub

' ----------------------------------------------------------------------------
' Tests for OpenWorkbook.
' ----------------------------------------------------------------------------
Public Sub Test_OpenWorkbook_RegisteredValue_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim file_path As String
    file_path = "C:\Data\MyWorkbook.xlsx"
    Call wb_double.Store.SetReturn("OpenWorkbook", "Book1.xlsx", file_path)

    ' Act
    Dim actual_name As String
    actual_name = wb_double.OpenWorkbook(file_path)

    ' Assert
    Assert.Equals "Book1.xlsx", actual_name

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_double.Store.GetLatestCall("OpenWorkbook", file_path)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_OpenWorkbook_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim actual_name As String
    actual_name = wb_double.OpenWorkbook("C:\Data\MyWorkbook.xlsx")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, wb_double.Store.GetCallCount("OpenWorkbook", "C:\Data\MyWorkbook.xlsx")
End Sub

' ----------------------------------------------------------------------------
' Tests for SaveWorkbook.
' ----------------------------------------------------------------------------
Public Sub Test_SaveWorkbook_RegisteredValue_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "Book1"
    Dim file_path As String
    file_path = "C:\Temp\Sub\Output.xlsx"
    Call wb_double.Store.SetReturn("SaveWorkbook", "Output.xlsx", wb_name, file_path, True)

    ' Act
    Dim actual_leaf As String
    actual_leaf = wb_double.SaveWorkbook(wb_name, file_path, True)

    ' Assert
    Assert.Equals "Output.xlsx", actual_leaf

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_double.Store.GetLatestCall("SaveWorkbook", wb_name, file_path, True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' Tests for IsSaved.
' ----------------------------------------------------------------------------
Public Sub Test_IsSaved_RegisteredValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "BookX"

    ' Register.
    Call wb_double.Store.SetReturn("IsSaved", True, wb_name)

    ' Act
    Dim actual_saved As Boolean
    actual_saved = wb_double.IsSaved(wb_name)

    ' Assert
    Assert.IsTrue actual_saved
End Sub

Public Sub Test_IsSaved_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    ' Act
    Dim actual_saved As Boolean
    actual_saved = wb_double.IsSaved("NoSuchBook")

    ' Assert
    Assert.IsFalse actual_saved
End Sub

' ----------------------------------------------------------------------------
' Tests for HasPath.
' ----------------------------------------------------------------------------
Public Sub Test_HasPath_RegisteredValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "BookX"

    Call wb_double.Store.SetReturn("HasPath", True, wb_name)

    ' Act
    Dim actual_has_path As Boolean
    actual_has_path = wb_double.HasPath(wb_name)

    ' Assert
    Assert.IsTrue actual_has_path
End Sub

Public Sub Test_HasPath_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    ' Act
    Dim actual_has_path As Boolean
    actual_has_path = wb_double.HasPath("NoSuchBook")

    ' Assert
    Assert.IsFalse actual_has_path
End Sub

' ----------------------------------------------------------------------------
' Tests for CloseWorkbook.
' ----------------------------------------------------------------------------
Public Sub Test_CloseWorkbook_WithFalse_RecordsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "TempBook"

    ' Act
    wb_double.CloseWorkbook wb_name, False

    ' Assert
    ' Here, arguments (wb_name, False) are the key and the value is True.
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_double.Store.GetLatestCall("CloseWorkbook", wb_name, False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_CloseWorkbook_NoForceArg_RecordsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    ' Act (Force omitted -> default True)
    wb_double.CloseWorkbook "NoArgBook"

    ' Assert
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_double.Store.GetLatestCall("CloseWorkbook", "NoArgBook", True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' Tests for WorksheetExists.
' ----------------------------------------------------------------------------
Public Sub Test_WorksheetExists_WithSetValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim ws_name As String
    ws_name = "SheetA"
    Dim wb_name As String
    wb_name = "MyBook.xlsm"

    ' Register (Sheet, Book) as the key with True as the value.
    Call wb_stub.Store.SetReturn("WorksheetExists", True, ws_name, wb_name)

    ' Act
    Dim actual_result As Boolean
    actual_result = wb_stub.WorksheetExists(ws_name, wb_name)

    ' Assert
    Assert.Equals True, actual_result
End Sub

Public Sub Test_WorksheetExists_NotRegistered_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    ' Act
    Dim actual_result As Boolean
    actual_result = wb_stub.WorksheetExists("NoSuchSheet", "NoSuchBook")

    ' Assert
    Assert.IsFalse actual_result
End Sub

Public Sub Test_WorksheetExists_PartialMatch_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Mismatch cases, such as the same sheet name with a different workbook or the same workbook name with a different sheet name.
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    ' Register.
    Call wb_stub.Store.SetReturn("WorksheetExists", True, "SheetX", "Book1.xlsm")

    ' Act
    Dim result_sameSheet_otherBook As Boolean
    result_sameSheet_otherBook = wb_stub.WorksheetExists("SheetX", "Book2.xlsm")

    Dim result_otherSheet_sameBook As Boolean
    result_otherSheet_sameBook = wb_stub.WorksheetExists("SheetY", "Book1.xlsm")

    ' Assert
    Assert.IsFalse result_sameSheet_otherBook
    Assert.IsFalse result_otherSheet_sameBook
End Sub

' ----------------------------------------------------------------------------
' Tests for GetAllWorksheets.
' ----------------------------------------------------------------------------
Public Sub Test_GetAllWorksheets_WithSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "SomeBook.xlsm"

    Dim expected_ws(0 To 2) As String
    expected_ws(0) = "SheetA"
    expected_ws(1) = "SheetB"
    expected_ws(2) = "SheetC"

    ' Register (Book) -> array.
    Call wb_stub.Store.SetReturn("GetAllWorksheets", expected_ws, wb_name)

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_stub.GetAllWorksheets(wb_name)

    ' Assert
    Assert.Equals "SheetA", actual_arr(0)
    Assert.Equals "SheetB", actual_arr(1)
    Assert.Equals "SheetC", actual_arr(2)
End Sub

Public Sub Test_GetOtherWorksheets_WithSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "SomeBook.xlsm"

    Dim expected_ws(0 To 1) As String
    expected_ws(0) = "SheetB"
    expected_ws(1) = "SheetC"

    Call wb_stub.Store.SetReturn("GetOtherWorksheets", expected_ws, wb_name)

    ' Act
    Dim actual_arr() As String
    actual_arr = wb_stub.GetOtherWorksheets(wb_name)

    ' Assert
    Assert.EqualsArray expected_ws, actual_arr
    Assert.EqualsNumeric 1, wb_stub.Store.GetCallCount("GetOtherWorksheets", wb_name)
End Sub

' ----------------------------------------------------------------------------
' Tests for AddWorksheet.
' ----------------------------------------------------------------------------
Public Sub Test_AddWorksheet_RegisteredGeneratedName_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim wb_name As String
    wb_name = "DataBook.xlsm"
    Call wb_stub.Store.SetReturn("AddWorksheet", "Sheet1", "", wb_name, 0&, False)

    ' Act
    Dim actual_sheet As String
    actual_sheet = wb_stub.AddWorksheet(, wb_name)

    ' Assert
    Assert.Equals "Sheet1", actual_sheet

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_stub.Store.GetLatestCall("AddWorksheet", "", wb_name, 0&, False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_AddWorksheet_RegisteredCustomName_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("AddWorksheet", "CustomSheet", "CustomSheet", "MyBook.xlsm", 2&, True)

    ' Act
    Dim actual_sheet As String
    actual_sheet = wb_stub.AddWorksheet("CustomSheet", "MyBook.xlsm", 2, True)

    ' Assert
    Assert.Equals "CustomSheet", actual_sheet

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_stub.Store.GetLatestCall("AddWorksheet", "CustomSheet", "MyBook.xlsm", 2&, True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_AddWorksheet_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim actual_sheet As String
    actual_sheet = wb_stub.AddWorksheet("NewSheet", "DataBook.xlsm")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, wb_stub.Store.GetCallCount("AddWorksheet", "NewSheet", "DataBook.xlsm", 0&, False)
End Sub

' ----------------------------------------------------------------------------
' Tests for RemoveWorksheet.
' ----------------------------------------------------------------------------
Public Sub Test_RemoveWorksheet_RecordsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    ' Act
    wb_stub.RemoveWorksheet "SheetToDelete", "Book99.xlsx"

    ' Assert
    ' Use ("SheetToDelete", "Book99.xlsx") as the dictionary key, with True as the value.
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_stub.Store.GetLatestCall("RemoveWorksheet", "SheetToDelete", "Book99.xlsx")
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' Tests for CopyWorksheet.
' ----------------------------------------------------------------------------
Public Sub Test_CopyWorksheet_RegisteredGeneratedName_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("CopyWorksheet", "Sheet1", "SrcSheet", "SrcBook.xlsm", "", "", 0&, False)

    ' Act
    Dim new_sheet As String
    new_sheet = wb_stub.CopyWorksheet("SrcSheet", "SrcBook.xlsm")

    ' Assert
    Assert.Equals "Sheet1", new_sheet

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_stub.Store.GetLatestCall("CopyWorksheet", "SrcSheet", "SrcBook.xlsm", "", "", 0&, False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_CopyWorksheet_RegisteredDestName_ReturnsAndRecordsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("CopyWorksheet", "NewSheetName", "OriginalSheet", "SourceBook.xlsm", "NewSheetName", "DestBook.xlsm", 5&, True)

    ' Act
    Dim copied_name As String
    copied_name = wb_stub.CopyWorksheet( _
                        "OriginalSheet", _
                        "SourceBook.xlsm", _
                        "NewSheetName", _
                        "DestBook.xlsm", _
                        5, _
                        True)

    ' Assert
    Assert.Equals "NewSheetName", copied_name

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = wb_stub.Store.GetLatestCall("CopyWorksheet", "OriginalSheet", "SourceBook.xlsm", "NewSheetName", "DestBook.xlsm", 5&, True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub
' ----------------------------------------------------------------------------
' Tests for RemoveVBComponents.
' ----------------------------------------------------------------------------

Public Sub Test_RemoveVBComponents_ArrayNames_DistinguishesCommaContainingNames(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    Dim first_names As Variant
    first_names = Array("A,B", "C")

    Dim second_names As Variant
    second_names = Array("A", "B,C")

    ' Act
    Call wb_stub.RemoveVBComponents(first_names, 1&, "Book.xlsm")
    Call wb_stub.RemoveVBComponents(second_names, 1&, "Book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, wb_stub.Store.GetLatestCall("RemoveVBComponents", first_names, 1&, "Book.xlsm").ArgumentCount
    Assert.EqualsNumeric 3, wb_stub.Store.GetLatestCall("RemoveVBComponents", second_names, 1&, "Book.xlsm").ArgumentCount
    Assert.EqualsNumeric 2, wb_stub.Store.Count
    Assert.EqualsNumeric 1, wb_stub.Store.GetCallCount("RemoveVBComponents", first_names, 1&, "Book.xlsm")
    Assert.EqualsNumeric 1, wb_stub.Store.GetCallCount("RemoveVBComponents", second_names, 1&, "Book.xlsm")
End Sub

Public Sub Test_CloseWorkbook_SameArgsTwice_RecordsBothCalls(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble

    ' Act
    Call wb_stub.CloseWorkbook("Book.xlsm", True)
    Call wb_stub.CloseWorkbook("Book.xlsm", True)

    ' Assert
    Assert.EqualsNumeric 2, wb_stub.Store.Count
    Assert.EqualsNumeric 2, wb_stub.Store.GetCallCount("CloseWorkbook", "Book.xlsm", True)
End Sub

Public Sub Test_WorkbookExists_SetError_RaisesWithoutRecordingCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    Call wb_double.Store.SetError("WorkbookExists", vbObjectError + 333, "Test Source", "exists failed", "Book.xlsx")

    Dim actual_result As Boolean
    actual_result = wb_double.WorkbookExists("Book.xlsx")

    If Not Assert.ErrorRaised(vbObjectError + 333, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, wb_double.Store.GetCallCount("WorkbookExists", "Book.xlsx")
End Sub

Public Sub Test_Find_RegisteredEmptyObjectList_ReturnsCountZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble

    Dim expected_list As ObjectList
    Set expected_list = New ObjectList
    Call wb_double.Store.SetReturn("Find", expected_list, "missing", "Book1.xlsm", True, True, True, True)

    Dim actual_list As ObjectList
    Set actual_list = wb_double.Find("missing", "Book1.xlsm")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_list.Count
    Assert.EqualsNumeric 1, wb_double.Store.GetCallCount("Find", "missing", "Book1.xlsm", True, True, True, True)
End Sub
