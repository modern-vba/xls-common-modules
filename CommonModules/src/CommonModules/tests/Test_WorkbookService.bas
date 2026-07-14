Attribute VB_Name = "Test_WorkbookService"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the WorkbookService class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Private Const C_VBEXT_CT_STDMODULE As Long = 1
Private Const C_VBEXT_CT_CLASSMODULE As Long = 2
Private Const C_TEMP_VB_COMPONENT_PREFIX As String = "Tmp_WbSvcRemove"

' -----------------------------------------------------------------------------
' GetThisWorkbookDirectoryPath
' -----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookDirectoryPath_SavedThisWorkbook_ReturnsWorkbookFolder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim expected_path As String
    expected_path = GetParentPath(ThisWorkbook.FullName)

    ' Act
    Dim actual_path As String
    actual_path = book_srv.GetThisWorkbookDirectoryPath()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

Private Function pPrepareTestSheet(ByRef SheetName As String) As Worksheet
    Dim target_sheet As Worksheet

    On Error Resume Next
        Set target_sheet = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If Not target_sheet Is Nothing Then
        Call target_sheet.Cells.Clear
    Else
        Set target_sheet = ThisWorkbook.Worksheets.Add
        target_sheet.Name = SheetName
    End If

    Set pPrepareTestSheet = target_sheet
End Function

Private Function pAddTempStandardModule(ByVal ComponentName As String) As Object
    Call pRemoveTempVBComponent(ComponentName)

    Dim vb_comp As Object
    Set vb_comp = ThisWorkbook.VBProject.VBComponents.Add(C_VBEXT_CT_STDMODULE)
    vb_comp.Name = ComponentName
    vb_comp.CodeModule.AddFromString "Option Explicit" & vbCrLf

    Set pAddTempStandardModule = vb_comp
End Function

Private Function pHasVBComponent(ByVal ComponentName As String) As Boolean
    On Error Resume Next
    Dim vb_comp As Object
    Set vb_comp = ThisWorkbook.VBProject.VBComponents.Item(ComponentName)
    pHasVBComponent = (Err.Number = 0 And Not vb_comp Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub pRemoveTempVBComponent(ByVal ComponentName As String)
    If Left$(ComponentName, Len(C_TEMP_VB_COMPONENT_PREFIX)) <> C_TEMP_VB_COMPONENT_PREFIX Then Err.Raise vbObjectError + 1, "Test_WorkbookService", "The name is not a temporary VBComponent name."

    On Error Resume Next
    Dim vb_comp As Object
    Set vb_comp = ThisWorkbook.VBProject.VBComponents.Item(ComponentName)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Call ThisWorkbook.VBProject.VBComponents.Remove(vb_comp)
    Err.Clear
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' WorkbookExists
' -----------------------------------------------------------------------------

Public Sub Test_WorkbookExists_NotWorkbookExists_ReturnsFalseAndClearsErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_value As Boolean
    actual_value = book_srv.WorkbookExists("NotWorkbookExists")

    Dim actual_error_number As Long
    actual_error_number = Err.Number
    On Error GoTo 0

    ' Assert
    Assert.Equals False, actual_value
    Assert.EqualsNumeric 0, actual_error_number
End Sub

Public Sub Test_WorkbookExists_WorkbookExists_ReturnsTrueAndKeepsErrClear(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_value As Boolean
    actual_value = book_srv.WorkbookExists(ThisWorkbook.Name)

    Dim actual_error_number As Long
    actual_error_number = Err.Number
    On Error GoTo 0

    ' Assert
    Assert.Equals True, actual_value
    Assert.EqualsNumeric 0, actual_error_number
End Sub

' -----------------------------------------------------------------------------
' GetAllWorkbooks
' -----------------------------------------------------------------------------

Public Sub Test_GetAllWorkbooks_IncludesThisWorkbook(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetAllWorkbooks

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue ContainsInArray(ThisWorkbook.Name, actual_value)
End Sub

Public Sub Test_GetOtherWorkbooks_ExcludesThisWorkbookAndIncludesAddedWorkbook(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetOtherWorkbooks

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue ContainsInArray(target_book_name, actual_value)
    Assert.IsFalse ContainsInArray(ThisWorkbook.Name, actual_value)
End Sub
' -----------------------------------------------------------------------------
' WorksheetExists
' -----------------------------------------------------------------------------

Public Sub Test_WorksheetExists_NotWorkbookExists_ReturnsFalseAndClearsErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_value As Boolean
    actual_value = book_srv.WorksheetExists("Sheet1", Book:="NotWorkbookExists")

    Dim actual_error_number As Long
    actual_error_number = Err.Number
    On Error GoTo 0

    ' Assert
    Assert.Equals False, actual_value
    Assert.EqualsNumeric 0, actual_error_number
End Sub

Public Sub Test_WorksheetExists_NotWorksheetExists_ReturnsFalseAndClearsErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_value As Boolean
    actual_value = book_srv.WorksheetExists("NotWorksheetExists", Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    actual_error_number = Err.Number
    On Error GoTo 0

    ' Assert
    Assert.Equals False, actual_value
    Assert.EqualsNumeric 0, actual_error_number
End Sub

Public Sub Test_WorksheetExists_WorksheetExists_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Dim actual_value As Boolean
    actual_value = book_srv.WorksheetExists("test_output")

    ' Assert
    Assert.Equals True, actual_value
End Sub

' -----------------------------------------------------------------------------
' GetAllWorksheets
' -----------------------------------------------------------------------------

Public Sub Test_GetAllWorksheets_Call_ReturnsAllWorksheets(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Dim actual_value() As String
    actual_value = book_srv.GetAllWorksheets

    ' Assert
    Assert.EqualsNumeric ThisWorkbook.Worksheets.Count, UBound(actual_value) - LBound(actual_value) + 1
End Sub

Public Sub Test_GetAllWorksheets_IncludesTargetActiveWorksheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "active_sheet"
    target_book.Worksheets(2).Name = "other_sheet"
    Call target_book.Worksheets("active_sheet").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim expected_count As Long
    expected_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetAllWorksheets(Book:=target_book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric expected_count, UBound(actual_value) - LBound(actual_value) + 1
    Assert.IsTrue ContainsInArray("active_sheet", actual_value)
    Assert.IsTrue ContainsInArray("other_sheet", actual_value)
End Sub

Public Sub Test_GetOtherWorksheets_ExcludesTargetActiveWorksheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "active_sheet"
    target_book.Worksheets(2).Name = "other_sheet"
    Call target_book.Worksheets("active_sheet").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim expected_count As Long
    expected_count = target_book.Worksheets.Count - 1

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetOtherWorksheets(Book:=target_book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric expected_count, UBound(actual_value) - LBound(actual_value) + 1
    Assert.IsFalse ContainsInArray("active_sheet", actual_value)
    Assert.IsTrue ContainsInArray("other_sheet", actual_value)
End Sub

Public Sub Test_GetOtherWorksheets_ActiveChartSheet_ReturnsAllWorksheets(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "first_sheet"
    target_book.Worksheets(2).Name = "second_sheet"

    Dim chart_sheet As Chart
    Set chart_sheet = target_book.Charts.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    chart_sheet.Name = "chart_sheet"
    Call chart_sheet.Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim expected_count As Long
    expected_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetOtherWorksheets(Book:=target_book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric expected_count, UBound(actual_value) - LBound(actual_value) + 1
    Assert.IsTrue ContainsInArray("first_sheet", actual_value)
    Assert.IsTrue ContainsInArray("second_sheet", actual_value)
    Assert.IsFalse ContainsInArray("chart_sheet", actual_value)
End Sub

Public Sub Test_GetOtherWorksheets_ActiveChartSheetWithSingleWorksheet_ReturnsSingleWorksheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim display_alerts As Boolean
    display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Do While target_book.Worksheets.Count > 1
        Call target_book.Worksheets(target_book.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = display_alerts

    target_book.Worksheets(1).Name = "only_sheet"

    Dim chart_sheet As Chart
    Set chart_sheet = target_book.Charts.Add(After:=target_book.Worksheets(1))
    chart_sheet.Name = "chart_sheet"
    Call chart_sheet.Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_value() As String
    actual_value = book_srv.GetOtherWorksheets(Book:=target_book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Application.DisplayAlerts = display_alerts
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric 1, UBound(actual_value) - LBound(actual_value) + 1
    Assert.IsTrue ContainsInArray("only_sheet", actual_value)
    Assert.IsFalse ContainsInArray("chart_sheet", actual_value)
End Sub

' -----------------------------------------------------------------------------
' AddWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_AddWorksheet_NegativeSheetIndexMinusOne_AddsAfterLastSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 3
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop

    Dim book_name As String
    book_name = target_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Dim actual_sheet_name As String
    Err.Clear
    actual_sheet_name = book_srv.AddWorksheet(Sheet:="test_negative_index", Book:=book_name, SheetIndex:=-1, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_index As Long
    If err_num = 0 Then
        actual_index = target_book.Worksheets(actual_sheet_name).Index
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.EqualsNumeric initial_sheet_count + 1, actual_index
End Sub

Public Sub Test_AddWorksheet_ActiveChartSheetIndexZero_AddsAfterChartSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "before_chart"
    target_book.Worksheets(2).Name = "after_chart"

    Dim chart_sheet As Chart
    Set chart_sheet = target_book.Charts.Add(After:=target_book.Worksheets("before_chart"))
    chart_sheet.Name = "chart_anchor"
    Call chart_sheet.Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_sheet_name As String
    actual_sheet_name = book_srv.AddWorksheet(Sheet:="added_after_chart", Book:=target_book_name, SheetIndex:=0, Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_added_index As Long
    Dim actual_chart_index As Long
    Dim actual_active_sheet_name As String
    If actual_error_number = 0 Then
        actual_added_index = target_book.Sheets("added_after_chart").Index
        actual_chart_index = target_book.Sheets("chart_anchor").Index
        actual_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "added_after_chart", actual_sheet_name
    Assert.EqualsNumeric actual_chart_index + 1, actual_added_index
    Assert.Equals "chart_anchor", actual_active_sheet_name
End Sub

Public Sub Test_AddWorksheet_NonActiveWorkbook_RestoresActiveWorkbookAndSheets(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_active"
    target_book.Worksheets(2).Name = "target_anchor"
    Call target_book.Worksheets("target_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Dim actual_sheet_name As String
    Err.Clear
    actual_sheet_name = book_srv.AddWorksheet(Sheet:="added_sheet", Book:=target_book_name, SheetIndex:=2, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "target_active", actual_target_active_sheet_name
End Sub

Public Sub Test_AddWorksheet_NameFailure_RestoresActiveWorkbookAndSheets(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_active"
    target_book.Worksheets(2).Name = "target_anchor"
    Call target_book.Worksheets("target_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Call book_srv.AddWorksheet(Sheet:="invalid:name", Book:=target_book_name, SheetIndex:=2, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    Dim actual_sheet_count As Long
    actual_active_book_name = ActiveWorkbook.Name
    actual_active_sheet_name = ActiveSheet.Name
    actual_target_active_sheet_name = target_book.ActiveSheet.Name
    actual_sheet_count = target_book.Worksheets.Count

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "target_active", actual_target_active_sheet_name
    Assert.EqualsNumeric initial_sheet_count, actual_sheet_count
End Sub

Public Sub Test_AddWorksheet_RenameFailure_DoesNotLeaveAddedSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_active"
    target_book.Worksheets(2).Name = "target_anchor"

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Call book_srv.AddWorksheet(Sheet:="'rename_fail", Book:=target_book_name, SheetIndex:=2, Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_sheet_count As Long
    actual_sheet_count = target_book.Worksheets.Count

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric initial_sheet_count, actual_sheet_count
End Sub
' -----------------------------------------------------------------------------
' RemoveWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_RemoveWorksheet_LastSheet_RestoresActiveWorkbook(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim old_display_alerts As Boolean
    old_display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Do While target_book.Worksheets.Count > 1
        Call target_book.Worksheets(target_book.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = old_display_alerts

    target_book.Worksheets(1).Name = "remove_me"

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Err.Clear
    Call book_srv.RemoveWorksheet(Sheet:="remove_me", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_sheet_count As Long
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_sheet_count = target_book.Worksheets.Count
    End If

    ' Cleanup
    Application.DisplayAlerts = old_display_alerts
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.EqualsNumeric 1, actual_target_sheet_count
End Sub

Public Sub Test_RemoveWorksheet_NonActiveWorkbook_RestoresTargetActiveSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "keep_active"
    target_book.Worksheets(2).Name = "delete_me"
    Call target_book.Worksheets("keep_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Err.Clear
    Call book_srv.RemoveWorksheet(Sheet:="delete_me", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "keep_active", actual_target_active_sheet_name
End Sub

' -----------------------------------------------------------------------------
' CopyWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_CopyWorksheet_DestinationNameUnused_ClearsInternalSearchErr(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "source_sheet"
    target_book.Worksheets(2).Name = "anchor_sheet"

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_sheet_name As String
    actual_sheet_name = book_srv.CopyWorksheet( _
            SourceWorksheetName:="source_sheet", _
            SourceWorkbookName:=target_book_name, _
            DestinationWorksheetName:="copied_sheet", _
            DestinationWorkbookName:=target_book_name, _
            SheetIndex:=2, _
            Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    If actual_error_number = 0 Then
        Dim copied_sheet As Worksheet
        Set copied_sheet = target_book.Worksheets("copied_sheet")
        actual_exists = Not copied_sheet Is Nothing
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "copied_sheet", actual_sheet_name
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_CopyWorksheet_ActiveChartSheetIndexZero_CopiesAfterChartSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "source_sheet"
    target_book.Worksheets(2).Name = "after_chart"

    Dim chart_sheet As Chart
    Set chart_sheet = target_book.Charts.Add(After:=target_book.Worksheets("source_sheet"))
    chart_sheet.Name = "chart_anchor"
    Call chart_sheet.Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_sheet_name As String
    actual_sheet_name = book_srv.CopyWorksheet( _
            SourceWorksheetName:="source_sheet", _
            SourceWorkbookName:=target_book_name, _
            DestinationWorksheetName:="copied_after_chart", _
            DestinationWorkbookName:=target_book_name, _
            SheetIndex:=0, _
            Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_copied_index As Long
    Dim actual_chart_index As Long
    Dim actual_active_sheet_name As String
    If actual_error_number = 0 Then
        actual_copied_index = target_book.Sheets("copied_after_chart").Index
        actual_chart_index = target_book.Sheets("chart_anchor").Index
        actual_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "copied_after_chart", actual_sheet_name
    Assert.EqualsNumeric actual_chart_index + 1, actual_copied_index
    Assert.Equals "chart_anchor", actual_active_sheet_name
End Sub

Public Sub Test_CopyWorksheet_TooLongDestinationName_DoesNotCopySheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "source_sheet"
    target_book.Worksheets(2).Name = "anchor_sheet"

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Call book_srv.CopyWorksheet( _
            SourceWorksheetName:="source_sheet", _
            SourceWorkbookName:=target_book_name, _
            DestinationWorksheetName:=String$(32, "a"), _
            DestinationWorkbookName:=target_book_name, _
            SheetIndex:=2, _
            Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_sheet_count As Long
    actual_sheet_count = target_book.Worksheets.Count

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
    Assert.EqualsNumeric initial_sheet_count, actual_sheet_count
End Sub

Public Sub Test_CopyWorksheet_RenameFailure_DoesNotLeaveCopiedSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "source_sheet"
    target_book.Worksheets(2).Name = "anchor_sheet"

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Err.Clear
    Call book_srv.CopyWorksheet( _
            SourceWorksheetName:="source_sheet", _
            SourceWorkbookName:=target_book_name, _
            DestinationWorksheetName:="'rename_fail", _
            DestinationWorkbookName:=target_book_name, _
            SheetIndex:=2, _
            Before:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_sheet_count As Long
    actual_sheet_count = target_book.Worksheets.Count

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric initial_sheet_count, actual_sheet_count
End Sub
' -----------------------------------------------------------------------------
' ActivateWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_ActivateWorksheet_NonActiveWorkbook_ActivatesTargetBookAndSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_start"
    target_book.Worksheets(2).Name = "target_activate"
    Call target_book.Worksheets("target_start").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    ' Act
    Err.Clear
    Call book_srv.ActivateWorksheet("target_activate", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals target_book_name, actual_active_book_name
    Assert.Equals "target_activate", actual_active_sheet_name
End Sub

' -----------------------------------------------------------------------------
' SaveWorkbook
' -----------------------------------------------------------------------------

Public Sub Test_SaveWorkbook_VisibleWorkbook_KeepsWindowVisible(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call InitializeCommonService
    Call InitializeFileSystemService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_workbook_service_visible.xlsx")
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Act
    book_name = book_srv.SaveWorkbook(book_name, file_path, Force:=True)

    Dim actual_visible As Boolean
    actual_visible = Workbooks(book_name).Windows(1).Visible

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name)
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_visible
End Sub

Public Sub Test_SaveWorkbook_DriveRelativePath_RaisesPathError(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String

    ' Act
    On Error GoTo ON_SAVE_ERROR
        Call book_srv.SaveWorkbook(book_name, "C:tmp_workbook_service_drive_relative.xlsx", Force:=True)
    On Error GoTo 0
    GoTo AFTER_SAVE

ON_SAVE_ERROR:
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    Resume AFTER_SAVE

AFTER_SAVE:
    On Error Resume Next
    Call book_srv.CloseWorkbook(book_name, Force:=True)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals "Function GetAbsolutePathFromParent", err_source
End Sub

Public Sub Test_SaveWorkbook_SaveAsFailure_RaisesError(ByVal Assert As UnitTestAssert)
    ' Arrange
    Call InitializeCommonService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_invalid:name.xlsx")

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String

    ' Act
    On Error GoTo ON_SAVE_ERROR
        Call book_srv.SaveWorkbook(book_name, file_path, Force:=True)
    On Error GoTo 0
    GoTo AFTER_SAVE

ON_SAVE_ERROR:
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    Resume AFTER_SAVE

AFTER_SAVE:
    On Error Resume Next
    Call book_srv.CloseWorkbook(book_name, Force:=True)
    On Error GoTo 0

    ' Assert
    Assert.ErrorRaised 0, err_num, err_source, err_desc
End Sub

' -----------------------------------------------------------------------------
' IsSaved
' -----------------------------------------------------------------------------

Public Sub Test_IsSaved_SavedWorkbookWithUnsavedChanges_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call InitializeCommonService
    Call InitializeFileSystemService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_workbook_service_is_saved.xlsx")
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    book_name = book_srv.SaveWorkbook(book_name, file_path, Force:=True)
    Workbooks(book_name).Worksheets(1).Range("A1").Value = "changed"

    ' Act
    Err.Clear
    Dim actual_saved As Boolean
    actual_saved = book_srv.IsSaved(book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name, Force:=True)
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_saved
End Sub

Public Sub Test_IsSaved_NewWorkbookWithoutChanges_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_saved As Boolean
    actual_saved = book_srv.IsSaved(book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_saved
End Sub

' -----------------------------------------------------------------------------
' HasPath
' -----------------------------------------------------------------------------

Public Sub Test_HasPath_NewWorkbook_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    ' Act
    Err.Clear
    Dim actual_has_path As Boolean
    actual_has_path = book_srv.HasPath(book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_has_path
End Sub

Public Sub Test_HasPath_SavedWorkbookWithUnsavedChanges_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call InitializeCommonService
    Call InitializeFileSystemService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_workbook_service_has_path.xlsx")
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    book_name = book_srv.SaveWorkbook(book_name, file_path, Force:=True)
    Workbooks(book_name).Worksheets(1).Range("A1").Value = "changed"

    ' Act
    Err.Clear
    Dim actual_has_path As Boolean
    actual_has_path = book_srv.HasPath(book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name, Force:=True)
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_has_path
End Sub


' -----------------------------------------------------------------------------
' RemoveVBComponents
' -----------------------------------------------------------------------------

Public Sub Test_RemoveVBComponents_StandardModule_RemovesModule(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "Ok"
    Call pAddTempStandardModule(component_name)

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(component_name, ComponentType:=C_VBEXT_CT_STDMODULE, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_RemoveVBComponents_ArrayNamesWithStaleErr_RemovesModule(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "ArrErr"
    Call pAddTempStandardModule(component_name)

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Err.Raise 5, "Test_WorkbookService", "stale error"
    Call book_srv.RemoveVBComponents(Array(component_name), ComponentType:=C_VBEXT_CT_STDMODULE, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_RemoveVBComponents_ScalarNameWithStaleErr_RemovesModule(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "FindErr"
    Call pAddTempStandardModule(component_name)

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Err.Raise 5, "Test_WorkbookService", "stale error"
    Call book_srv.RemoveVBComponents(component_name, ComponentType:=C_VBEXT_CT_STDMODULE, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_RemoveVBComponents_CaseVariantNames_DefaultRemovesOnce(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "Case"
    Call pAddTempStandardModule(component_name)

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(Array(LCase$(component_name), UCase$(component_name)), ComponentType:=C_VBEXT_CT_STDMODULE, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_RemoveVBComponents_DocumentMixed_PreservesStandardModule(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "Doc"
    Call pAddTempStandardModule(component_name)

    Dim document_component_name As String
    document_component_name = ThisWorkbook.Worksheets(1).CodeName

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(Array(component_name, document_component_name), Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_RemoveVBComponents_ComponentTypeMismatch_PreservesStandardModule(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_name As String
    component_name = C_TEMP_VB_COMPONENT_PREFIX & "Type"
    Call pAddTempStandardModule(component_name)

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(component_name, ComponentType:=C_VBEXT_CT_CLASSMODULE, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_exists As Boolean
    actual_exists = pHasVBComponent(component_name)
    Call pRemoveTempVBComponent(component_name)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_RemoveVBComponents_NullComponentName_RaisesWorkbookServiceError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(Null, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
End Sub

Public Sub Test_RemoveVBComponents_ErrorValueComponentName_RaisesWorkbookServiceError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(CVErr(xlErrNA), Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
End Sub

Public Sub Test_RemoveVBComponents_UninitializedArray_RaisesWorkbookServiceError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim component_names() As String

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Call book_srv.RemoveVBComponents(component_names, Book:=ThisWorkbook.Name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue StartsWith(actual_error_source, "Class WorkbookService.")
End Sub

Public Sub Test_Find_UninitializedWsSrv_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WsSrv = Nothing

    Dim book_srv As WorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = book_srv.Find("not-found", Book:=ThisWorkbook.Name)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class WorkbookService.Find", actual_err_source
    Assert.Equals "InitializeCommonService has not been called.", actual_err_desc
End Sub
