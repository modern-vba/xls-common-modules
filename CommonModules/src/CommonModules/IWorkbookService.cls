VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkbookService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for a service that manages opening, closing, and writing Excel files.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWorkbookService", Description:="This class is intended for interfaces."
End Sub

'* Gets the workbook name of ThisWorkbook.
'*
'* @return Workbook name of ThisWorkbook.
'*
'* @details
'* Used by callers to get their own workbook name without directly referencing ThisWorkbook.
Public Function GetThisWorkbookName() As String
End Function

'* Gets the directory path based on ThisWorkbook.
'*
'* @return Absolute directory path.
'*
'* @details
'* If ThisWorkbook has been saved, returns its saved destination directory.
'* If ThisWorkbook has not been saved, returns the result of CurDir converted to an absolute path.
Public Function GetThisWorkbookDirectoryPath() As String
End Function

'* Checks whether a workbook exists.
'*
'* @param Book Workbook name.
'* @return Whether the workbook exists.
'*
'* @details
'* Checks whether the workbook name exists.
Public Function WorkbookExists(ByVal Book As String) As Boolean
End Function

'* Gets a list of all open workbooks.
'*
'* @return Workbook list.
'*
'* @details
'* Gets an array of names for all open workbooks, including ThisWorkbook.
Public Function GetAllWorkbooks() As String()
End Function

'* Gets a list of workbooks other than ThisWorkbook.
'*
'* @return Workbook list.
'*
'* @details
'* Gets an array of workbook names excluding ThisWorkbook.
Public Function GetOtherWorkbooks() As String()
End Function

'* Checks whether a worksheet exists.
'*
'* @param Sheet Worksheet name.
'* @param Book Workbook name. If omitted, checks inside ThisWorkbook.
'* @return Whether the worksheet exists.
'*
'* @details
'* Checks whether a worksheet exists in a workbook.
'* If no workbook with the specified name exists, no error is raised and False is returned.
Public Function WorksheetExists(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
End Function

'* Gets a list of worksheets.
'*
'* @param Book Workbook name. If omitted, checks inside ThisWorkbook.
'* @return Worksheet list.
'*
'* @details
'* Gets an array of all worksheet names in the target workbook, including the target workbook ActiveSheet.
Public Function GetAllWorksheets(Optional ByVal Book As String = "") As String()
End Function

'* Gets a list of worksheets other than the target workbook ActiveSheet.
'*
'* @param Book Workbook name. If omitted, checks inside ThisWorkbook.
'* @return Worksheet list.
'*
'* @details
'* Gets an array of worksheet names excluding the target workbook ActiveSheet.
Public Function GetOtherWorksheets(Optional ByVal Book As String = "") As String()
End Function

'* Opens a workbook.
'*
'* @param FilePath [Optional] File path.
'* @return Workbook name.
'*
'* @details
'* Opens a workbook. If the file path is omitted, creates a new workbook.
'* The source file is not locked.
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
End Function

'* Saves a workbook.
'*
'* @param Book Target workbook name.
'* @param FilePath Destination file path.
'* @param Force [Optional] Whether to save forcibly.
'* @return Workbook name after saving.
'*
'* @details
'* Saves a workbook. When Force is False and the file already exists, a confirmation message is displayed.
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
End Function

'* Checks whether the workbook has not changed since its last save.
'*
'* @param Book Target workbook name.
'* @return Whether the workbook has not changed since its last save.
'*
'* @details
'* Has the same meaning as Excel Workbook.Saved and does not check whether a save destination path exists.
Public Function IsSaved(ByVal Book As String) As Boolean
End Function

'* Checks whether the workbook has a save destination path.
'*
'* @param Book Target workbook name.
'* @return Whether the workbook has a save destination path.
'*
'* @details
'* Returns False when the workbook has no save destination path, such as a new unsaved workbook. Does not check whether it has changes.
Public Function HasPath(ByVal Book As String) As Boolean
End Function

'* Closes a workbook.
'*
'* @param Book Target workbook name.
'* @param Force [Optional] Whether to close forcibly. When True, discards changes and closes without confirmation.
'*
'* @details
'* Closes a workbook. When Force is False, a confirmation message is displayed.
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
End Sub

'* Searches the entire workbook.
'*
'* @param What String to search for.
'* @param Book [Optional] Workbook to search. If omitted, ThisWorkbook.Name is used.
'* @param LookInValues [Optional] Whether to search values. The default is True, which searches values. When False, searches formulas.
'* @param LookAtWhole [Optional] Whether to search for cells whose contents match exactly. The default is True, which searches exact matches. When False, searches partial matches.
'* @param MatchCase [Optional] Whether to distinguish uppercase and lowercase. The default is True. When True, searches case-sensitively. When False, searches case-insensitively.
'* @param MatchByte [Optional] Whether to distinguish full-width and half-width characters. The default is True. When True, distinguishes them. When False, does not distinguish them.
'* @return ObjectList containing WorksheetRangeBounds for found cells. Count = 0 when there are no search results.
'*
'* @details
'* Searches the entire workbook.
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList
End Function

'* Adds a worksheet.
'*
'* @param Sheet [Optional] Worksheet name to add. When omitted, it is named automatically.
'* @param Book [Optional] Workbook name. When omitted, the workbook name of ThisWorkbook is used.
'* @param SheetIndex [Optional] Worksheet index. When omitted, the active sheet is used.
'* @param Before [Optional] Whether to add before the target. When omitted, adds after the specified worksheet index.
'* @return Name of the added worksheet.
'*
'* @details
'* Adds a worksheet. After processing, restores the active workbook before the call and the active sheet of the target workbook.
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
End Function

'* Deletes a worksheet.
'*
'* @param Sheet [Optional] Worksheet name to delete. When omitted, the active sheet is used.
'* @param Book [Optional] Workbook name. When omitted, the workbook name of ThisWorkbook is used.
'*
'* @details
'* Deletes a worksheet. After processing, restores the active workbook before the call and the active sheet of the target workbook as far as possible.
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
End Sub

'* Duplicates a worksheet.
'*
'* @param SourceWorksheetName [Optional] Worksheet name to duplicate. When omitted, the active sheet is used.
'* @param SourceWorkbookName [Optional] Name of the workbook containing the worksheet to duplicate. When omitted, the workbook name of ThisWorkbook is used.
'* @param DestinationWorksheetName [Optional] Name to give to the duplicated worksheet. When omitted, it is named automatically.
'* @param DestinationWorkbookName [Optional] Destination workbook name. When omitted, the workbook name of ThisWorkbook is used.
'* @param SheetIndex [Optional] Worksheet index. When omitted, the active sheet in the destination workbook is used.
'* @param Before [Optional] Whether to add the duplicated sheet before the target. When omitted, adds after the specified worksheet index.
'* @return Name of the added worksheet.
'*
'* @details
'* Duplicates a worksheet.
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
End Function

'* Activates a worksheet.
'*
'* @param Sheet [Optional] Worksheet name to activate.
'* @param Book [Optional] Workbook name. When omitted, the workbook name of ThisWorkbook is used.
'*
'* @details
'* Activates the specified workbook and worksheet.
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
End Sub

'* Removes VBA components.
'*
'* @param ComponentNames Component name or array of component names to remove.
'* @param ComponentType [Optional] VBIDE component type to remove. When 0, the type is not limited.
'* @param Book [Optional] Workbook name. When omitted, the workbook name of ThisWorkbook is used.
'*
'* @details
'* Removes VBA components with the specified names from the target workbook VBProject.
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
End Sub
