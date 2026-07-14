Attribute VB_Name = "Test_CommonServiceBoundary"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for shared service resolution and test replacement boundaries.
'!
' #############################################################################

Public Sub Test_InitializeCommonService_NotInitialized_InitializesAllPublicServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    ' Act
    Err.Clear
    Call InitializeCommonService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "FileSystemService", actual_fs_type
    Assert.Equals "TextFileService", actual_tf_type
End Sub

Public Sub Test_InitializeCommonService_InjectedServices_PreservesAllPublicServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set WbSrv = New WorkbookServiceTestDouble
    Set WsSrv = New WorksheetServiceTestDouble
    Set FsSrv = New FileSystemServiceTestDouble
    Set TfSrv = New TextFileServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeCommonService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookServiceTestDouble", actual_wb_type
    Assert.Equals "WorksheetServiceTestDouble", actual_ws_type
    Assert.Equals "FileSystemServiceTestDouble", actual_fs_type
    Assert.Equals "TextFileServiceTestDouble", actual_tf_type
End Sub

Public Sub Test_InitializeCommonService_ForceTrue_RecreatesAllPublicServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set WbSrv = New WorkbookServiceTestDouble
    Set WsSrv = New WorksheetServiceTestDouble
    Set FsSrv = New FileSystemServiceTestDouble
    Set TfSrv = New TextFileServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeCommonService(Force:=True)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "FileSystemService", actual_fs_type
    Assert.Equals "TextFileService", actual_tf_type
End Sub

Public Sub Test_InitializeWorkbookService_NotInitialized_InitializesWorkbookService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    ' Act
    Err.Clear
    Call InitializeWorkbookService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "Nothing", actual_ws_type
    Assert.Equals "Nothing", actual_fs_type
    Assert.Equals "Nothing", actual_tf_type
End Sub

Public Sub Test_InitializeWorksheetService_NotInitialized_InitializesWorksheetService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    ' Act
    Err.Clear
    Call InitializeWorksheetService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "Nothing", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "Nothing", actual_fs_type
    Assert.Equals "Nothing", actual_tf_type
End Sub

Public Sub Test_InitializeUdfCommonService_NotInitialized_InitializesWorkbookAndWorksheetServicesOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    ' Act
    Err.Clear
    Call InitializeUdfCommonService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "Nothing", actual_fs_type
    Assert.Equals "Nothing", actual_tf_type
End Sub

Public Sub Test_InitializeUdfCommonService_ForceTrue_RecreatesWorkbookAndWorksheetServicesOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set WbSrv = New WorkbookServiceTestDouble
    Set WsSrv = New WorksheetServiceTestDouble
    Set FsSrv = New FileSystemServiceTestDouble
    Set TfSrv = New TextFileServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeUdfCommonService(Force:=True)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "FileSystemServiceTestDouble", actual_fs_type
    Assert.Equals "TextFileServiceTestDouble", actual_tf_type
End Sub

Public Sub Test_InitializeUdfCommonService_InjectedServices_PreservesAllServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set WbSrv = New WorkbookServiceTestDouble
    Set WsSrv = New WorksheetServiceTestDouble
    Set FsSrv = New FileSystemServiceTestDouble
    Set TfSrv = New TextFileServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeUdfCommonService

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    actual_wb_type = pGetTypeName(WbSrv)
    actual_ws_type = pGetTypeName(WsSrv)
    actual_fs_type = pGetTypeName(FsSrv)
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "WorkbookServiceTestDouble", actual_wb_type
    Assert.Equals "WorksheetServiceTestDouble", actual_ws_type
    Assert.Equals "FileSystemServiceTestDouble", actual_fs_type
    Assert.Equals "TextFileServiceTestDouble", actual_tf_type
End Sub

Public Sub Test_InitializeFileSystemService_ForceTrue_RecreatesFileSystemService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set FsSrv = New FileSystemServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeFileSystemService(Force:=True)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_fs_type As String
    actual_fs_type = pGetTypeName(FsSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "FileSystemService", actual_fs_type
End Sub

Public Sub Test_InitializeTextFileService_ForceTrue_RecreatesTextFileService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Set TfSrv = New TextFileServiceTestDouble

    ' Act
    Err.Clear
    Call InitializeTextFileService(Force:=True)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_tf_type As String
    actual_tf_type = pGetTypeName(TfSrv)

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "TextFileService", actual_tf_type
End Sub

Public Sub Test_InitializeCommonService_ApostropheWorkbookName_InitializesOptionalServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim old_display_alerts As Boolean
    old_display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim temp_book_path As String
    temp_book_path = Environ$("TEMP") & "\CommonModules_O'Brien_optional_service_probe.xlsm"
    If Dir$(temp_book_path) <> "" Then Kill temp_book_path

    Call ThisWorkbook.SaveCopyAs(temp_book_path)

    Dim probe_book As Workbook
    Set probe_book = Workbooks.Open(temp_book_path)

    ' Act
    Err.Clear
    Dim actual_probe_result As Variant
    actual_probe_result = Application.Run(pBuildWorkbookMacroName( _
            probe_book.Name, _
            "Test_CommonServiceBoundary.Probe_InitializeCommonServiceOptionalServiceTypes"))

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_has_probe_error As Boolean
    Dim actual_probe_error_number As Long
    Dim actual_probe_error_source As String
    Dim actual_probe_error_description As String
    Dim actual_wb_type As String
    Dim actual_ws_type As String
    Dim actual_fs_type As String
    Dim actual_tf_type As String
    If actual_error_number = 0 Then
        actual_has_probe_error = CBool(actual_probe_result(0))
        If actual_has_probe_error Then
            actual_probe_error_number = CLng(actual_probe_result(1))
            actual_probe_error_source = CStr(actual_probe_result(2))
            actual_probe_error_description = CStr(actual_probe_result(3))
        Else
            actual_wb_type = CStr(actual_probe_result(1))
            actual_ws_type = CStr(actual_probe_result(2))
            actual_fs_type = CStr(actual_probe_result(3))
            actual_tf_type = CStr(actual_probe_result(4))
        End If
    End If

    ' Cleanup
    If Not probe_book Is Nothing Then Call probe_book.Close(SaveChanges:=False)
    If Dir$(temp_book_path) <> "" Then Kill temp_book_path
    Application.DisplayAlerts = old_display_alerts
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsFalse actual_has_probe_error
    If actual_has_probe_error Then
        If Not Assert.ErrorNotRaised(0, actual_probe_error_number, actual_probe_error_source, actual_probe_error_description) Then Exit Sub
    End If
    Assert.Equals "WorkbookService", actual_wb_type
    Assert.Equals "WorksheetService", actual_ws_type
    Assert.Equals "FileSystemService", actual_fs_type
    Assert.Equals "TextFileService", actual_tf_type
End Sub

Public Sub Test_FileSystemService_GetAbsolutePath_UsesInjectedWorkbookService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("GetThisWorkbookDirectoryPath", "C:\InjectedBase")
    Set WbSrv = wb_stub

    Dim fs_srv As FileSystemService
    Set fs_srv = New FileSystemService

    ' Act
    Err.Clear
    Dim actual_path As String
    actual_path = fs_srv.GetAbsolutePath("Child.txt")

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "C:\InjectedBase\Child.txt", actual_path
End Sub

Public Sub Test_TextFileEntity_Initialize_UsesInjectedFileSystemService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    Call fs_stub.Store.SetReturn("GetAbsolutePath", "C:\Resolved\input.txt", "relative.txt")
    Set FsSrv = fs_stub

    Dim txt_entity As TextFileEntity
    Set txt_entity = New TextFileEntity

    ' Act
    Err.Clear
    Call txt_entity.Initialize("relative.txt")

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_path As String
    If actual_error_number = 0 Then actual_path = txt_entity.FilePath

    ' Cleanup
    Set txt_entity = Nothing
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "C:\Resolved\input.txt", actual_path
End Sub

Public Sub Test_WorksheetRangeBounds_DefaultBook_UsesInjectedWorkbookService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices

    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("GetThisWorkbookName", "InjectedBook.xlsm")
    Set WbSrv = wb_stub

    ' Act
    Err.Clear
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="Sheet1")

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_book_name As String
    If actual_error_number = 0 Then actual_book_name = range_bounds.WorkbookName

    ' Cleanup
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "InjectedBook.xlsm", actual_book_name
End Sub

Public Sub Test_UserInputSheet_Initialize_UsesInjectedWorkbookService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("GetThisWorkbookName", "InjectedBook.xlsm")
    Call wb_stub.Store.SetReturn("WorksheetExists", True, "InputSheet", "InjectedBook.xlsm")
    Set WbSrv = wb_stub

    Dim expected_area As WorksheetRangeBounds
    Set expected_area = New_RangeBounds(Sheet:="InputSheet")

    Dim input_sheet As UserInputSheet
    Set input_sheet = New UserInputSheet

    ' Act
    Err.Clear
    Call input_sheet.Initialize(expected_area)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_area As WorksheetRangeBounds
    Dim actual_exists_call_count As Long
    If actual_error_number = 0 Then
        Set actual_area = input_sheet.InputArea
        actual_exists_call_count = wb_stub.Store.GetCallCount("WorksheetExists", "InputSheet", "InjectedBook.xlsm")
    End If

    ' Cleanup
    Set input_sheet = Nothing
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue expected_area.Equals(actual_area)
    Assert.EqualsNumeric 1, actual_exists_call_count
End Sub

Public Sub Test_WorkbookService_Find_UsesInjectedWorksheetService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call pClearCommonServices
    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim old_display_alerts As Boolean
    old_display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Do While target_book.Worksheets.Count > 1
        Call target_book.Worksheets(target_book.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = old_display_alerts

    target_book.Worksheets(1).Name = "find_sheet"

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim search_range As WorksheetRangeBounds
    Set search_range = New_RangeBounds(Sheet:="find_sheet", Book:=target_book_name)

    Dim found_cell As WorksheetRangeBounds
    Set found_cell = New_RangeBounds(Row:=2, Column:=3, Sheet:="find_sheet", Book:=target_book_name)

    Dim found_list As ObjectList
    Set found_list = New ObjectList
    Call found_list.Add(found_cell)

    Call ws_stub.Store.SetReturn("Find", found_list, "needle", search_range, True, True, True, True)

    Dim book_srv As WorkbookService
    Set book_srv = New WorkbookService

    ' Act
    Err.Clear
    Dim actual_list As ObjectList
    Set actual_list = book_srv.Find("needle", Book:=target_book_name)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_count As Long
    If actual_error_number = 0 Then actual_count = actual_list.Count

    ' Cleanup
    Application.DisplayAlerts = old_display_alerts
    Call target_book.Close(SaveChanges:=False)
    Call pClearCommonServices
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_count
End Sub

Public Function Probe_InitializeCommonServiceOptionalServiceTypes() As Variant
    On Error GoTo ON_ERROR

    Call pClearCommonServices
    Call InitializeCommonService

    Dim success_result(0 To 4) As Variant
    success_result(0) = False
    success_result(1) = pGetTypeName(WbSrv)
    success_result(2) = pGetTypeName(WsSrv)
    success_result(3) = pGetTypeName(FsSrv)
    success_result(4) = pGetTypeName(TfSrv)

    Probe_InitializeCommonServiceOptionalServiceTypes = success_result

    Call pClearCommonServices
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Err.Clear

    On Error Resume Next
    Call pClearCommonServices
    On Error GoTo 0

    Dim error_result(0 To 4) As Variant
    error_result(0) = True
    error_result(1) = err_num
    error_result(2) = err_source
    error_result(3) = err_desc
    error_result(4) = ""

    Probe_InitializeCommonServiceOptionalServiceTypes = error_result
End Function

Private Function pBuildWorkbookMacroName(ByVal WorkbookName As String, ByVal MacroName As String) As String
    pBuildWorkbookMacroName = "'" & Replace(WorkbookName, "'", "''") & "'!" & MacroName
End Function

Private Function pGetTypeName(ByVal TargetObject As Object) As String
    If TargetObject Is Nothing Then
        pGetTypeName = "Nothing"
    Else
        pGetTypeName = TypeName(TargetObject)
    End If
End Function

Private Sub pClearCommonServices()
    Set WbSrv = Nothing
    Set WsSrv = Nothing
    Set FsSrv = Nothing
    Set TfSrv = Nothing
End Sub
