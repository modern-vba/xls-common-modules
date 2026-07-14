VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorkbookService

' #############################################################################
'!
'! @brief
'! Implementation of the service that manages opening, closing, and writing Excel files.
'!
' #############################################################################

Private Const C_VBEXT_CT_DOCUMENT As Long = 100

'* Gets the workbook name of ThisWorkbook.
'*
'* @return Workbook name of ThisWorkbook.
'*
'* @details
'* Used by callers to get their own workbook name without directly referencing ThisWorkbook.
Public Function GetThisWorkbookName() As String
    GetThisWorkbookName = ThisWorkbook.Name
End Function
Private Function IWorkbookService_GetThisWorkbookName() As String
    IWorkbookService_GetThisWorkbookName = GetThisWorkbookName()
End Function

'* Gets the directory path based on ThisWorkbook.
'*
'* @return Absolute directory path.
'*
'* @details
'* If ThisWorkbook has been saved, returns its saved destination directory.
'* If ThisWorkbook has not been saved, returns the result of CurDir converted to an absolute path.
Public Function GetThisWorkbookDirectoryPath() As String
    Dim dir_path As String

    dir_path = ThisWorkbook.Path
    If dir_path = "" Then
        dir_path = CurDir()
    End If

    GetThisWorkbookDirectoryPath = GetAbsolutePathFromParent(dir_path, ".")
End Function
Private Function IWorkbookService_GetThisWorkbookDirectoryPath() As String
    IWorkbookService_GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath()
End Function

' Checks whether a workbook exists.
'
' @param Book Workbook name.
' @return Whether the workbook exists.
'
' @details
' Checks whether the workbook name exists.
Public Function WorkbookExists(ByVal Book As String) As Boolean
    Dim target_book As Workbook
    On Error Resume Next
        Set target_book = Workbooks(Book)
    Err.Clear
    On Error GoTo 0

    If target_book Is Nothing Then
        WorkbookExists = False
    Else
        WorkbookExists = True
    End If
End Function
Private Function IWorkbookService_WorkbookExists(ByVal Book As String) As Boolean
    IWorkbookService_WorkbookExists = WorkbookExists(Book)
End Function

' Gets a list of all open workbooks.
'
' @return List of workbooks.
'
' @details
' Gets a list of all open workbook names, including ThisWorkbook, as a string array.
Public Function GetAllWorkbooks() As String()
    Dim result_arr() As String
    result_arr = EmptyStringArray()

    Dim book_item As Workbook
    For Each book_item In Workbooks
        Call pAddItemToStringArray(result_arr, book_item.Name)
    Next book_item
    GetAllWorkbooks = result_arr
End Function
Private Function IWorkbookService_GetAllWorkbooks() As String()
    IWorkbookService_GetAllWorkbooks = GetAllWorkbooks()
End Function

' Gets a list of workbooks other than ThisWorkbook.
'
' @return List of workbooks.
'
' @details
' Gets a list of workbook names other than ThisWorkbook as a string array.
Public Function GetOtherWorkbooks() As String()
    Dim result_arr() As String
    result_arr = EmptyStringArray()

    Dim book_item As Workbook
    For Each book_item In Workbooks
        If Not book_item Is ThisWorkbook Then Call pAddItemToStringArray(result_arr, book_item.Name)
    Next book_item
    GetOtherWorkbooks = result_arr
End Function
Private Function IWorkbookService_GetOtherWorkbooks() As String()
    IWorkbookService_GetOtherWorkbooks = GetOtherWorkbooks()
End Function

' Checks whether a worksheet exists.
'
' @param Sheet Worksheet name.
' @param Book Workbook name. If omitted, checks inside ThisWorkbook.
' @return Whether the worksheet exists.
'
' @details
' Checks whether the worksheet exists in the workbook.
' Returns False instead of raising an error even when no workbook with the specified name exists.
Public Function WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        On Error Resume Next
            Set target_book = Workbooks(Book)
        Err.Clear
        On Error GoTo 0

        If target_book Is Nothing Then
            WorksheetExists = False
            Exit Function
        End If
    End If

    Dim target_sheet As Worksheet
    On Error Resume Next
        Set target_sheet = target_book.Worksheets(Sheet)
    Err.Clear
    On Error GoTo 0

    If target_sheet Is Nothing Then
        WorksheetExists = False
    Else
        WorksheetExists = True
    End If
End Function
Private Function IWorkbookService_WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    IWorkbookService_WorksheetExists = WorksheetExists(Sheet, Book:=Book)
End Function

' Gets a list of worksheets.
'
' @param Book Workbook name. If omitted, checks inside ThisWorkbook.
' @return List of worksheets.
'
' @details
' Gets a list of all worksheet names in the workbook, including the target book's ActiveSheet, as a string array.
Public Function GetAllWorksheets(Optional ByVal Book As String = "") As String()
    GetAllWorksheets = pGetWorksheetNames(Book, IncludeActiveSheet:=True)
End Function
Private Function IWorkbookService_GetAllWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetAllWorksheets = GetAllWorksheets(Book:=Book)
End Function

' Gets a list of worksheets other than the target book's ActiveSheet.
'
' @param Book Workbook name. If omitted, checks inside ThisWorkbook.
' @return List of worksheets.
'
' @details
' Gets a list of worksheet names excluding the target book's ActiveSheet as a string array.
Public Function GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    GetOtherWorksheets = pGetWorksheetNames(Book, IncludeActiveSheet:=False)
End Function
Private Function IWorkbookService_GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetOtherWorksheets = GetOtherWorksheets(Book:=Book)
End Function

' Opens a workbook.
'
' @param FilePath [Optional] File path.
' @return Book name.
'
' @details
' Opens a workbook. If FilePath is omitted, creates a new workbook.
' Does not lock the original file.
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    Dim current_active_workbook As Workbook
    On Error Resume Next
    Set current_active_workbook = ActiveWorkbook
    Err.Clear
    On Error GoTo ON_ERROR

    ' Open the workbook.
    Dim workbook_obj As Workbook
    If FilePath = "" Then
        Set workbook_obj = Workbooks.Add()
    Else
        Dim file_path As String
        file_path = pGetAbsolutePath(FilePath)
        Set workbook_obj = Workbooks.Add(Template:=file_path)
    End If
    ActiveWindow.Visible = False

    ' Restore application settings changed by opening the workbook.
    If Not current_active_workbook Is Nothing Then Call current_active_workbook.Activate
    Call app_state.Restore

    OpenWorkbook = workbook_obj.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If Not current_active_workbook Is Nothing Then Call current_active_workbook.Activate
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_OpenWorkbook(Optional ByVal FilePath As String = "") As String
    IWorkbookService_OpenWorkbook = OpenWorkbook(FilePath:=FilePath)
End Function

' Saves a workbook.
'
' @param Book Target book name.
' @param FilePath Destination file path.
' @param Force [Optional] Whether to force saving.
' @return Book name after saving.
'
' @details
' Saves a workbook. When Force is False and the file already exists, a confirmation message is displayed.
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    Dim target_book As Workbook
    Set target_book = Workbooks(Book)

    Dim file_path As String
    file_path = pGetAbsolutePath(FilePath)

    ' Save the state of the active book and the target window.
    Dim current_active_workbook As Workbook
    On Error Resume Next
    Set current_active_workbook = ActiveWorkbook
    Err.Clear
    On Error GoTo ON_ERROR

    Dim target_window_visible As Boolean
    Dim has_target_window_visible As Boolean
    target_window_visible = target_book.Windows(1).Visible
    has_target_window_visible = True

    Application.Calculation = xlCalculationAutomatic
    If Force Then Application.DisplayAlerts = False
    If Not target_window_visible Then target_book.Windows(1).Visible = True
    Call target_book.Activate
    Call target_book.SaveAs(file_path, FileFormat:=GetExcelFileFormat(file_path))
    target_book.Windows(1).Visible = target_window_visible

    ' Restore the active book and application settings changed by saving.
    If Not current_active_workbook Is Nothing Then
        If current_active_workbook.Name <> target_book.Name Or target_window_visible Then
            Call current_active_workbook.Activate
        End If
    End If
    Call app_state.Restore

    ' Return the book name after saving.
    SaveWorkbook = target_book.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If has_target_window_visible Then target_book.Windows(1).Visible = target_window_visible
    If Not current_active_workbook Is Nothing Then
        If target_book Is Nothing Then
            Call current_active_workbook.Activate
        ElseIf current_active_workbook.Name <> target_book.Name Or target_window_visible Then
            Call current_active_workbook.Activate
        End If
    End If
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    IWorkbookService_SaveWorkbook = SaveWorkbook(Book, FilePath, Force:=Force)
End Function

' Checks whether the workbook has not changed since it was last saved.
'
' @param Book Target book name.
' @return Whether the workbook has not changed since it was last saved.
'
' @details
' Has the same meaning as Excel's Workbook.Saved and does not check whether a destination path exists.
Public Function IsSaved(ByVal Book As String) As Boolean
    IsSaved = Workbooks(Book).Saved
End Function
Private Function IWorkbookService_IsSaved(ByVal Book As String) As Boolean
    IWorkbookService_IsSaved = IsSaved(Book)
End Function

' Checks whether the workbook has a destination path.
'
' @param Book Target book name.
' @return Whether the workbook has a destination path.
'
' @details
' Returns False when the workbook has no destination path, such as a new unsaved workbook. Does not check whether it has changed.
Public Function HasPath(ByVal Book As String) As Boolean
    HasPath = (Workbooks(Book).Path <> "")
End Function
Private Function IWorkbookService_HasPath(ByVal Book As String) As Boolean
    IWorkbookService_HasPath = HasPath(Book)
End Function

' Closes a workbook.
'
' @param Book Target book name.
' @param Force [Optional] Whether to force closing. When True, closes without confirmation and discards changes.
'
' @details
' Closes a workbook. When Force is False, a confirmation message is displayed.
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    If Force Then
        Call Workbooks(Book).Close(SaveChanges:=False)
    Else
        Call Workbooks(Book).Close
    End If
End Sub
Private Sub IWorkbookService_CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call CloseWorkbook(Book, Force:=Force)
End Sub

' Searches the entire workbook.
'
' @param What String to search for.
' @param LookInValues [Optional] Whether to search values. The default is True. When False, searches formulas.
' @param LookAtWhole [Optional] Whether to search for cells whose contents are exactly the same. The default is True. True searches for exact matches. False searches for partial matches.
' @param MatchCase [Optional] Whether to distinguish uppercase and lowercase. The default is True. True searches case-sensitively. False searches case-insensitively.
' @param MatchByte [Optional] Whether to distinguish full-width and half-width characters. The default is True. True distinguishes them. False does not distinguish them.
' @return ObjectList containing WorksheetRangeBounds for found cells. Count = 0 when there are no search results.
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Call pRequireWorksheetService("Find")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = WsSrv

    Dim book_name As String
    If Book = "" Then
        book_name = ThisWorkbook.Name
    Else
        book_name = Book
    End If

    Dim sheet_arr() As String
    sheet_arr = GetAllWorksheets(book_name)

    If Not IsEmptyArray(sheet_arr) Then
        Dim sheet_item As Variant 'String
        For Each sheet_item In sheet_arr
            Dim range_item As WorksheetRangeBounds
            Set range_item = New_RangeBounds(Sheet:=sheet_item, Book:=book_name)
            Dim tmp_list As ObjectList
            Set tmp_list = sheet_srv.Find(What, range_item, LookInValues:=LookInValues, LookAtWhole:=LookAtWhole, MatchCase:=MatchCase, MatchByte:=MatchByte)
            Call result_value.AddOther(tmp_list)
        Next sheet_item
    End If

    Set Find = result_value
End Function
Private Function IWorkbookService_Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Set IWorkbookService_Find = Find( _
            What, Book:=Book, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

Private Sub pRequireWorksheetService(ByVal MethodName As String)
    If WsSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WorkbookService" & "." & MethodName, Description:="InitializeCommonService has not been called."
    End If
End Sub

' Adds a worksheet.
'
' @param Sheet [Optional] Worksheet name to add. If omitted, it is named automatically.
' @param Book [Optional] Workbook name. If omitted, ThisWorkbook's workbook name is used.
' @param SheetIndex [Optional] Worksheet index. If omitted, the active sheet is treated as specified.
' @param Before [Optional] Whether to add before. If omitted, adds after the specified worksheet index.
' @return Name of the added worksheet.
'
' @details
' Adds a worksheet.
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    On Error GoTo ON_ERROR

    ' Get the workbook to process.
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' Save the active state.
    Dim current_active_book As Workbook
    Dim current_target_active_sheet As Object
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_target_active_sheet = target_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' Check the worksheet name.
    Call pRaiseIfInvalidWorksheetName(Sheet, "Sheet")

    Dim existing_sheet As Worksheet
    On Error Resume Next
    Set existing_sheet = target_book.Worksheets(Sheet)
    Err.Clear
    On Error GoTo ON_ERROR
    If Not existing_sheet Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService.AddWorksheet", "The sheet name to add already exists. (" & Sheet & "@" & Book & ")"

    ' Get the target sheet for insertion.
    Dim target_sheet As Object
    Set target_sheet = pGetTargetSheet(target_book, SheetIndex)

    Dim added_sheet As Worksheet

    ' Add the worksheet.
    If Before Then
        Set added_sheet = target_book.Worksheets.Add(Before:=target_sheet)
    Else
        Set added_sheet = target_book.Worksheets.Add(After:=target_sheet)
    End If

    ' Rename the worksheet if a worksheet name was specified.
    If Sheet <> "" Then added_sheet.Name = Sheet

    Call pRestoreActiveState(current_active_book, current_target_active_sheet)

    AddWorksheet = added_sheet.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    Call pDeleteWorksheetIfCreated(added_sheet)
    Call pRestoreActiveState(current_active_book, current_target_active_sheet)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    IWorkbookService_AddWorksheet = AddWorksheet( _
            Sheet:=Sheet, Book:=Book, _
            SheetIndex:=SheetIndex, Before:=Before)
End Function

' Deletes a worksheet.
'
' @param Sheet [Optional] Worksheet name to delete. If omitted, the active sheet is treated as specified.
' @param Book [Optional] Workbook name. If omitted, ThisWorkbook's workbook name is used.
'
' @details
' Deletes a worksheet.
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    ' Get the workbook that contains the worksheet to delete.
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' Save the active state.
    Dim current_active_book As Workbook
    Dim current_target_active_sheet As Object
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_target_active_sheet = target_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' Get the worksheet to delete.
    Dim target_sheet As Worksheet
    If Sheet = "" Then
        Set target_sheet = target_book.ActiveSheet
    Else
        Set target_sheet = target_book.Worksheets(Sheet)
    End If

    ' If the workbook containing the target worksheet has no worksheets other than the target, add an empty worksheet.
    If target_book.Worksheets.Count <= 1 Then target_book.Worksheets.Add

    ' Delete the target worksheet.
    Application.DisplayAlerts = False
    Call target_sheet.Delete
    Call pRestoreActiveState(current_active_book, current_target_active_sheet)
    Call app_state.Restore
    On Error GoTo 0
    Exit Sub

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    Call pRestoreActiveState(current_active_book, current_target_active_sheet)
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub
Private Sub IWorkbookService_RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call RemoveWorksheet(Sheet:=Sheet, Book:=Book)
End Sub

' Copies a worksheet.
'
' @param SourceWorksheetName [Optional] Worksheet name to copy. If omitted, the active sheet is treated as specified.
' @param SourceWorkbookName [Optional] Workbook name containing the worksheet to copy. If omitted, ThisWorkbook's workbook name is used.
' @param DestinationWorksheetName [Optional] Name to give to the copied worksheet. If omitted, it is named automatically.
' @param DestinationWorkbookName [Optional] Destination workbook name. If omitted, ThisWorkbook's workbook name is used.
' @param SheetIndex [Optional] Worksheet index. If omitted, the active sheet in the destination workbook is treated as specified.
' @param Before [Optional] Whether to add the copied sheet before the target. If omitted, adds after the specified worksheet index.
' @return Name of the added worksheet.
'
' @details
' Copies a worksheet.
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    On Error GoTo ON_ERROR

    ' Get the workbook containing the source worksheet.
    Dim src_book As Workbook
    If SourceWorkbookName = "" Then
        Set src_book = ThisWorkbook
    Else
        Set src_book = Application.Workbooks(SourceWorkbookName)
    End If

    ' Get the source worksheet.
    Dim src_sheet As Worksheet
    If SourceWorksheetName = "" Then
        Set src_sheet = src_book.ActiveSheet
    Else
        Set src_sheet = src_book.Worksheets(SourceWorksheetName)
    End If

    ' Get the destination workbook.
    Dim dst_book As Workbook
    If DestinationWorkbookName = "" Then
        Set dst_book = ThisWorkbook
    Else
        Set dst_book = Application.Workbooks(DestinationWorkbookName)
    End If

    ' Check the worksheet name.
    Call pRaiseIfInvalidWorksheetName(DestinationWorksheetName, "DestinationWorksheetName")

    Dim existing_sheet As Worksheet
    On Error Resume Next
    Set existing_sheet = dst_book.Worksheets(DestinationWorksheetName)
    Err.Clear
    On Error GoTo ON_ERROR
    If Not existing_sheet Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService.CopyWorksheet", "The destination sheet name already exists. (" & DestinationWorksheetName & "@" & DestinationWorkbookName & ")"

    ' Save the active state.
    Dim current_active_book As Workbook
    Dim current_destination_active_sheet As Object
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_destination_active_sheet = dst_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' Get the target sheet for insertion.
    Dim target_sheet As Object
    Set target_sheet = pGetTargetSheet(dst_book, SheetIndex)

    Dim added_sheet As Worksheet

    ' Add the worksheet.
    If Before Then
        Call src_sheet.Copy(Before:=target_sheet)
    Else
        Call src_sheet.Copy(After:=target_sheet)
    End If

    ' Get the copied sheet.
    Set added_sheet = dst_book.ActiveSheet

    ' Rename the worksheet if a worksheet name was specified.
    If DestinationWorksheetName <> "" Then added_sheet.Name = DestinationWorksheetName

    ' Restore the active state.
    Call pRestoreActiveState(current_active_book, current_destination_active_sheet)

    CopyWorksheet = added_sheet.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    Call pDeleteWorksheetIfCreated(added_sheet)

    Call pRestoreActiveState(current_active_book, current_destination_active_sheet)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    IWorkbookService_CopyWorksheet = CopyWorksheet( _
            SourceWorksheetName:=SourceWorksheetName, _
            SourceWorkbookName:=SourceWorkbookName, _
            DestinationWorksheetName:=DestinationWorksheetName, _
            DestinationWorkbookName:=DestinationWorkbookName, _
            SheetIndex:=SheetIndex, _
            Before:=Before)
End Function

'* Removes VBA components.
'*
'* @param ComponentNames Component name or array of component names to remove.
'* @param ComponentType [Optional] VBIDE component type to remove. When 0, the type is not limited.
'* @param Book [Optional] Workbook name. When omitted, the workbook name of ThisWorkbook is used.
'*
'* @details
'* Removes VBA components with the specified names from the target workbook VBProject.
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    Dim vb_comps As Object
    Set vb_comps = target_book.VBProject.VBComponents

    Dim component_name_set As Object
    Set component_name_set = pBuildComponentNameSet(ComponentNames)

    Call pValidateVBComponentsForRemove(vb_comps, component_name_set, ComponentType)

    Dim component_name As Variant
    For Each component_name In component_name_set.Keys
        Call vb_comps.Remove(vb_comps.Item(CStr(component_name)))
    Next component_name
End Sub
Private Sub IWorkbookService_RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call RemoveVBComponents(ComponentNames, ComponentType:=ComponentType, Book:=Book)
End Sub

Private Function pBuildComponentNameSet(ByVal ComponentNames As Variant) As Object
    Dim component_name_set As Object
    Set component_name_set = CreateObject("Scripting.Dictionary")
    component_name_set.CompareMode = vbTextCompare

    If IsArray(ComponentNames) Then
        If Not pIsArrayInitialized(ComponentNames) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pBuildComponentNameSet", "Specify one or more component names in ComponentNames."

        Dim name_item As Variant
        For Each name_item In ComponentNames
            Call pAddComponentName(component_name_set, name_item)
        Next name_item
    Else
        Call pAddComponentName(component_name_set, ComponentNames)
    End If

    If component_name_set.Count = 0 Then Err.Raise vbObjectError + 1, "Class WorkbookService.pBuildComponentNameSet", "Specify one or more component names in ComponentNames."

    Set pBuildComponentNameSet = component_name_set
End Function

Private Sub pValidateVBComponentsForRemove(ByVal VBComponents As Object, ByVal ComponentNameSet As Object, ByVal ComponentType As Long)
    Dim component_name As Variant
    For Each component_name In ComponentNameSet.Keys
        Dim vb_comp As Object
        Set vb_comp = pFindVBComponent(VBComponents, CStr(component_name))

        If vb_comp Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService.pValidateVBComponentsForRemove", "The VBA component to remove does not exist. (" & CStr(component_name) & ")"
        If ComponentType <> 0 And vb_comp.Type <> ComponentType Then Err.Raise vbObjectError + 1, "Class WorkbookService.pValidateVBComponentsForRemove", "The VBA component type to remove does not match. (" & CStr(component_name) & ")"
        If vb_comp.Type = C_VBEXT_CT_DOCUMENT Then Err.Raise vbObjectError + 1, "Class WorkbookService.pValidateVBComponentsForRemove", "Document modules cannot be removed. (" & CStr(component_name) & ")"
    Next component_name
End Sub

'
' Activates a worksheet.
'
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    ' Get the workbook containing the target worksheet.
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' Get the target worksheet.
    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(Sheet)

    Call target_book.Activate
    Call target_sheet.Activate
End Sub
Private Sub IWorkbookService_ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call ActivateWorksheet(Sheet, Book:=Book)
End Sub

Private Sub pAddItemToStringArray(ByRef TargetArray() As String, ByVal ItemToBeAdded As String)
    Dim new_ubound As Long

    If IsEmptyArray(TargetArray) Then
        new_ubound = 0
        ReDim TargetArray(0 To 0) As String
    Else
        new_ubound = UBound(TargetArray) + 1
        ReDim Preserve TargetArray(LBound(TargetArray) To new_ubound) As String
    End If
    TargetArray(new_ubound) = ItemToBeAdded
End Sub

Private Function pGetWorksheetNames(ByVal Book As String, ByVal IncludeActiveSheet As Boolean) As String()
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Workbooks(Book)
    End If

    Dim active_sheet As Object
    If Not IncludeActiveSheet Then Set active_sheet = target_book.ActiveSheet

    Dim result_arr() As String
    result_arr = EmptyStringArray()

    Dim sheet_item As Worksheet
    For Each sheet_item In target_book.Worksheets
        If IncludeActiveSheet Then
            Call pAddItemToStringArray(result_arr, sheet_item.Name)
        ElseIf Not TypeOf active_sheet Is Worksheet Then
            Call pAddItemToStringArray(result_arr, sheet_item.Name)
        ElseIf Not sheet_item Is active_sheet Then
            Call pAddItemToStringArray(result_arr, sheet_item.Name)
        End If
    Next sheet_item

    pGetWorksheetNames = result_arr
End Function

Private Function pGetAbsolutePath(ByVal PathLikeString As String) As String
    pGetAbsolutePath = GetAbsolutePathFromParent(GetThisWorkbookDirectoryPath(), PathLikeString)
End Function

Private Function pGetTargetSheet(ByVal TargetBook As Workbook, ByVal SheetIndex As Long) As Object
    If SheetIndex = 0 Then
        ' If the sheet index is 0, use the active sheet across the entire Sheets collection as the reference.
        Set pGetTargetSheet = TargetBook.ActiveSheet
    ElseIf SheetIndex < 0 Then
        ' If the sheet index is less than 0, count from the end of the worksheets. -1 is the last sheet, and -2 is the second from last.
        Set pGetTargetSheet = TargetBook.Worksheets(MaxLng(TargetBook.Worksheets.Count + 1 + SheetIndex, 1))
    Else
        ' If the sheet index exceeds the worksheet count, clamp it to the end.
        Set pGetTargetSheet = TargetBook.Worksheets(MinLng(SheetIndex, TargetBook.Worksheets.Count))
    End If
End Function

Private Sub pRaiseIfInvalidWorksheetName(ByVal WorksheetName As String, ByVal ArgumentName As String)
    If WorksheetName = "" Then Exit Sub

    If Len(WorksheetName) > 31 Then
        Err.Raise vbObjectError + 1, "Class WorkbookService.pRaiseIfInvalidWorksheetName", ArgumentName & " must be a sheet name of 31 characters or fewer. (" & WorksheetName & ")"
    End If
    If Left$(WorksheetName, 1) = "'" Or Right$(WorksheetName, 1) = "'" Then
        Err.Raise vbObjectError + 1, "Class WorkbookService.pRaiseIfInvalidWorksheetName", ArgumentName & " cannot start or end with an apostrophe. (" & WorksheetName & ")"
    End If

    Dim invalid_chars As Variant
    invalid_chars = Array(":", "\", "/", "?", "*", "[", "]")

    Dim invalid_char As Variant
    For Each invalid_char In invalid_chars
        If InStr(1, WorksheetName, CStr(invalid_char), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 1, "Class WorkbookService.pRaiseIfInvalidWorksheetName", ArgumentName & " contains a character that cannot be used in a sheet name. (" & CStr(invalid_char) & ")"
        End If
    Next invalid_char
End Sub

Private Sub pRestoreActiveState(ByVal ActiveBook As Workbook, ByVal TargetBookActiveSheet As Object)
    On Error Resume Next
    If Not TargetBookActiveSheet Is Nothing Then Call TargetBookActiveSheet.Activate
    If Not ActiveBook Is Nothing Then Call ActiveBook.Activate
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub pDeleteWorksheetIfCreated(ByVal CreatedSheet As Worksheet)
    If CreatedSheet Is Nothing Then Exit Sub

    On Error Resume Next
    Dim display_alerts As Boolean
    display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Call CreatedSheet.Delete
    Application.DisplayAlerts = display_alerts
    Err.Clear
    On Error GoTo 0
End Sub

Private Function pIsArrayInitialized(ByVal TargetArray As Variant) As Boolean
    Err.Clear
    On Error Resume Next
    Dim lower_bound As Long
    lower_bound = LBound(TargetArray)
    Dim upper_bound As Long
    upper_bound = UBound(TargetArray)
    pIsArrayInitialized = (Err.Number = 0 And lower_bound <= upper_bound)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub pAddComponentName(ByVal ComponentNameSet As Object, ByVal ComponentNameValue As Variant)
    If IsArray(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames must contain string component names."
    If IsObject(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames must contain string component names."
    If IsError(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames cannot contain error values."
    If IsNull(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames cannot contain Null."
    If IsEmpty(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames cannot contain Empty."
    If VarType(ComponentNameValue) <> vbString Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames must contain string component names."

    Dim component_name As String
    component_name = CStr(ComponentNameValue)
    If component_name = "" Then Err.Raise vbObjectError + 1, "Class WorkbookService.pAddComponentName", "ComponentNames cannot contain empty strings."

    If Not ComponentNameSet.Exists(component_name) Then Call ComponentNameSet.Add(component_name, component_name)
End Sub

Private Function pFindVBComponent(ByVal VBComponents As Object, ByVal ComponentName As String) As Object
    Err.Clear
    On Error Resume Next
    Set pFindVBComponent = VBComponents.Item(ComponentName)
    If Err.Number <> 0 Then
        Set pFindVBComponent = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

