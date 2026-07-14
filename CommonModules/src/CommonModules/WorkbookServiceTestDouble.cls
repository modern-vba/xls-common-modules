VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorkbookService

' #############################################################################
'!
'! @brief
'! Test double for the service that manages opening, closing, and writing Excel files.
'!
' #############################################################################

'* Stub settings and history of successfully completed calls.
'* Members that need return values and depend on external state or parent resources are strict stubs; members where a conservative default is natural, such as existence checks, use safe defaults.
'* Operations without return values are spy only, and unregistered strict stub errors and explicit errors from SetError are not recorded in call history.
'* spy only records call arguments in the order of the original method definition.
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* Safe default for GetThisWorkbookName. Returns "ThisWorkbook.xlsm" when unregistered.
Public Function GetThisWorkbookName() As String
    Call Store.RaiseIfError("GetThisWorkbookName")

    Dim result_value As String
    result_value = Store.GetReturnOrDefault("GetThisWorkbookName", "ThisWorkbook.xlsm")
    Call Store.RecordCall("GetThisWorkbookName")

    GetThisWorkbookName = result_value
End Function
Private Function IWorkbookService_GetThisWorkbookName() As String
    IWorkbookService_GetThisWorkbookName = GetThisWorkbookName()
End Function

'* Safe default for GetThisWorkbookDirectoryPath. Returns an empty string as unsaved ThisWorkbook equivalent when unregistered.
Public Function GetThisWorkbookDirectoryPath() As String
    Call Store.RaiseIfError("GetThisWorkbookDirectoryPath")

    Dim result_value As String
    result_value = Store.GetReturnOrDefault("GetThisWorkbookDirectoryPath", "")
    Call Store.RecordCall("GetThisWorkbookDirectoryPath")

    GetThisWorkbookDirectoryPath = result_value
End Function
Private Function IWorkbookService_GetThisWorkbookDirectoryPath() As String
    IWorkbookService_GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath()
End Function

'* Safe default for WorkbookExists. Returns False when unregistered.
Public Function WorkbookExists(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("WorkbookExists", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("WorkbookExists", False, Book)
    Call Store.RecordCall("WorkbookExists", Book)

    WorkbookExists = result_value
End Function
Private Function IWorkbookService_WorkbookExists(ByVal Book As String) As Boolean
    IWorkbookService_WorkbookExists = WorkbookExists(Book)
End Function

'* Safe default for GetAllWorkbooks. Returns a one-item array linked to GetThisWorkbookName when unregistered.
Public Function GetAllWorkbooks() As String()
    Call Store.RaiseIfError("GetAllWorkbooks")

    Dim result_value() As String
    If Store.HasReturn("GetAllWorkbooks") Then
        result_value = Store.GetReturn("GetAllWorkbooks")
    Else
        ReDim result_value(0 To 0) As String
        result_value(0) = pGetThisWorkbookNameForDefault()
    End If

    Call Store.RecordCall("GetAllWorkbooks")
    GetAllWorkbooks = result_value
End Function
Private Function IWorkbookService_GetAllWorkbooks() As String()
    IWorkbookService_GetAllWorkbooks = GetAllWorkbooks()
End Function

Private Function pGetThisWorkbookNameForDefault() As String
    Call Store.RaiseIfError("GetThisWorkbookName")
    pGetThisWorkbookNameForDefault = Store.GetReturnOrDefault("GetThisWorkbookName", "ThisWorkbook.xlsm")
End Function

'* Safe default for GetOtherWorkbooks. Returns an empty array representing no other workbooks when unregistered.
Public Function GetOtherWorkbooks() As String()
    Call Store.RaiseIfError("GetOtherWorkbooks")

    Dim result_value() As String
    If Store.HasReturn("GetOtherWorkbooks") Then
        result_value = Store.GetReturn("GetOtherWorkbooks")
    Else
        result_value = EmptyStringArray()
    End If

    Call Store.RecordCall("GetOtherWorkbooks")
    GetOtherWorkbooks = result_value
End Function
Private Function IWorkbookService_GetOtherWorkbooks() As String()
    IWorkbookService_GetOtherWorkbooks = GetOtherWorkbooks()
End Function

'* Safe default for WorksheetExists. Returns False when unregistered.
Public Function WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    Call Store.RaiseIfError("WorksheetExists", Sheet, Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("WorksheetExists", False, Sheet, Book)
    Call Store.RecordCall("WorksheetExists", Sheet, Book)

    WorksheetExists = result_value
End Function
Private Function IWorkbookService_WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    IWorkbookService_WorksheetExists = WorksheetExists(Sheet, Book:=Book)
End Function

'* Strict stub for GetAllWorksheets. Register the return value explicitly in Store.
Public Function GetAllWorksheets(Optional ByVal Book As String = "") As String()
    Call Store.RaiseIfError("GetAllWorksheets", Book)

    Dim result_value() As String
    result_value = Store.GetReturn("GetAllWorksheets", Book)
    Call Store.RecordCall("GetAllWorksheets", Book)

    GetAllWorksheets = result_value
End Function
Private Function IWorkbookService_GetAllWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetAllWorksheets = GetAllWorksheets(Book:=Book)
End Function

'* Strict stub for GetOtherWorksheets. Register the return value explicitly in Store.
Public Function GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    Call Store.RaiseIfError("GetOtherWorksheets", Book)

    Dim result_value() As String
    result_value = Store.GetReturn("GetOtherWorksheets", Book)
    Call Store.RecordCall("GetOtherWorksheets", Book)

    GetOtherWorksheets = result_value
End Function
Private Function IWorkbookService_GetOtherWorksheets(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetOtherWorksheets = GetOtherWorksheets(Book:=Book)
End Function

'* Strict stub for OpenWorkbook. Register the return value explicitly in Store.
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
    Call Store.RaiseIfError("OpenWorkbook", FilePath)

    Dim result_value As String
    result_value = Store.GetReturn("OpenWorkbook", FilePath)
    Call Store.RecordCall("OpenWorkbook", FilePath)

    OpenWorkbook = result_value
End Function
Private Function IWorkbookService_OpenWorkbook(Optional ByVal FilePath As String = "") As String
    IWorkbookService_OpenWorkbook = OpenWorkbook(FilePath:=FilePath)
End Function

'* Strict stub for SaveWorkbook. Register the return value explicitly in Store.
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    Call Store.RaiseIfError("SaveWorkbook", Book, FilePath, Force)

    Dim result_value As String
    result_value = Store.GetReturn("SaveWorkbook", Book, FilePath, Force)
    Call Store.RecordCall("SaveWorkbook", Book, FilePath, Force)

    SaveWorkbook = result_value
End Function
Private Function IWorkbookService_SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    IWorkbookService_SaveWorkbook = SaveWorkbook(Book, FilePath, Force:=Force)
End Function

'* Safe default for IsSaved. Returns False when unregistered.
Public Function IsSaved(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("IsSaved", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("IsSaved", False, Book)
    Call Store.RecordCall("IsSaved", Book)

    IsSaved = result_value
End Function
Private Function IWorkbookService_IsSaved(ByVal Book As String) As Boolean
    IWorkbookService_IsSaved = IsSaved(Book)
End Function

'* Safe default for HasPath. Returns False when unregistered.
Public Function HasPath(ByVal Book As String) As Boolean
    Call Store.RaiseIfError("HasPath", Book)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("HasPath", False, Book)
    Call Store.RecordCall("HasPath", Book)

    HasPath = result_value
End Function
Private Function IWorkbookService_HasPath(ByVal Book As String) As Boolean
    IWorkbookService_HasPath = HasPath(Book)
End Function

'* spy only for CloseWorkbook. Records the call; inject failures with Store.SetError.
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call Store.RaiseIfError("CloseWorkbook", Book, Force)
    Call Store.RecordCall("CloseWorkbook", Book, Force)
End Sub
Private Sub IWorkbookService_CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call CloseWorkbook(Book, Force:=Force)
End Sub

'* Strict stub for Find. Explicitly register an empty ObjectList as well to represent no search results.
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Call Store.RaiseIfError("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)

    Dim result_value As ObjectList
    Set result_value = Store.GetReturn("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)
    Call Store.RecordCall("Find", What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)

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

'* Strict stub for AddWorksheet. Register the return value explicitly in Store.
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    Call Store.RaiseIfError("AddWorksheet", Sheet, Book, SheetIndex, Before)

    Dim result_value As String
    result_value = Store.GetReturn("AddWorksheet", Sheet, Book, SheetIndex, Before)
    Call Store.RecordCall("AddWorksheet", Sheet, Book, SheetIndex, Before)

    AddWorksheet = result_value
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

'* spy only for RemoveWorksheet. Records the call; inject failures with Store.SetError.
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call Store.RaiseIfError("RemoveWorksheet", Sheet, Book)
    Call Store.RecordCall("RemoveWorksheet", Sheet, Book)
End Sub
Private Sub IWorkbookService_RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call RemoveWorksheet(Sheet:=Sheet, Book:=Book)
End Sub

'* Strict stub for CopyWorksheet. Register the return value explicitly in Store.
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    Call Store.RaiseIfError("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)

    Dim result_value As String
    result_value = Store.GetReturn("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)
    Call Store.RecordCall("CopyWorksheet", SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)

    CopyWorksheet = result_value
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

'* spy only for ActivateWorksheet. Records the call; inject failures with Store.SetError.
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call Store.RaiseIfError("ActivateWorksheet", Sheet, Book)
    Call Store.RecordCall("ActivateWorksheet", Sheet, Book)
End Sub
Private Sub IWorkbookService_ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call ActivateWorksheet(Sheet, Book:=Book)
End Sub

'* spy only for RemoveVBComponents. Records the call; inject failures with Store.SetError.
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call Store.RaiseIfError("RemoveVBComponents", ComponentNames, ComponentType, Book)
    Call Store.RecordCall("RemoveVBComponents", ComponentNames, ComponentType, Book)
End Sub
Private Sub IWorkbookService_RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call RemoveVBComponents(ComponentNames, ComponentType:=ComponentType, Book:=Book)
End Sub
