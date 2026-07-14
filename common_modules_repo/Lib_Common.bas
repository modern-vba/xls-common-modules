Attribute VB_Name = "Lib_Common"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Standard module that groups basic functions and related helpers.
'! It also contains members not used by this tool because they are shared with other tools.
'!
' #############################################################################

'* WorkbookService. Replace with a test double during unit tests.
Public WbSrv As IWorkbookService

'* WorksheetService. Replace with a test double during unit tests.
Public WsSrv As IWorksheetService

'* DebugInformation. Shared run state that holds debug information during GUI execution.
Public DbgInfo As DebugInformation

'* ProgressStatus. Shared run state that holds progress state during GUI execution.
Public ProgStat As ProgressStatus


'* Represents an unset date.
Public Const G_DATE_NULL As Date = #12/31/1899#

'* Minimum valid date.
Public Const G_DATE_MINIMUM As Date = #3/1/1900#

'* Maximum row number.
Public Const G_ROW_MAX As Long = 1048576

'* Maximum column number.
Public Const G_COL_MAX As Long = 16384

'* Omitted value for row and column numbers.
Public Const G_OMIT_CELL_INDEX As Long = -2147483648#

'* Value indicating that WorksheetService.SetAlignment should not change the setting.
Public Const G_ALIGNMENT_NO_CHANGE As Long = -2147483648#

'* File system path separator.
Public Const G_FS_PATH_SEP As String = "\"

'* Value indicating that GetTypedValueKey / GetValueKey should key object references.
Public Const G_OBJECT_KEY_MODE_REFERENCE As Long = 0

'* Value indicating that GetTypedValueKey / GetValueKey should key IEquatable.GetIdentityString.
Public Const G_OBJECT_KEY_MODE_I_EQUATABLE As Long = 1

'* Value indicating that GetTypedValueKey / GetValueKey should key IDuplicateCheckable.GetKey.
Public Const G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE As Long = 2

'* Value indicating that WorksheetRangeBounds should be enumerated by row.
Public Const G_RANGE_ENUM_MODE_ROWS As Long = 0

'* Value indicating that WorksheetRangeBounds should be enumerated by column.
Public Const G_RANGE_ENUM_MODE_COLUMNS As Long = 1

'* Value indicating that WorksheetRangeBounds should be enumerated by cells in row-major order.
Public Const G_RANGE_ENUM_MODE_CELLS_HORIZONTAL As Long = 2

'* Value indicating that WorksheetRangeBounds should be enumerated by cells in column-major order.
Public Const G_RANGE_ENUM_MODE_CELLS_VERTICAL As Long = 3

Private Const C_LONG_MAX As Long = 2147483647#
Private Const C_LONG_MIN As Long = -2147483648#
Private Const C_INTEGER_MAX As Long = 32767
Private Const C_INTEGER_MIN As Long = -32768
Private Const C_A1_TOKEN_CELL As Long = 1
Private Const C_A1_TOKEN_ROW As Long = 2
Private Const C_A1_TOKEN_COLUMN As Long = 3

' #############################################################################
'
' Common service initialization
'
' #############################################################################

'* Initializes WorkbookService.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Creates WorkbookService when WbSrv is unset or Force is True. If it has been replaced in tests, keeps it only when Force is False.
Public Sub InitializeWorkbookService(Optional ByVal Force As Boolean = False)
    If Force Or (WbSrv Is Nothing) Then Set WbSrv = New WorkbookService
End Sub

'* Initializes WorksheetService.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Creates WorksheetService when WsSrv is unset or Force is True. If it has been replaced in tests, keeps it only when Force is False.
Public Sub InitializeWorksheetService(Optional ByVal Force As Boolean = False)
    If Force Or (WsSrv Is Nothing) Then Set WsSrv = New WorksheetService
End Sub

'* Initializes shared services that can be used safely from UDFs.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Initializes only WorkbookService and WorksheetService so Application.Run is not reached during recalculation of Excel worksheet functions. FileSystemService and TextFileService are not initialized.
Public Sub InitializeUdfCommonService(Optional ByVal Force As Boolean = False)
    Call InitializeWorkbookService(Force)
    Call InitializeWorksheetService(Force)
End Sub

'* Initializes shared services.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Initializes shared services. The initialization targets are as follows.
'*
'* * WorkbookService
'* * WorksheetService
'* * FileSystemService (if imported).
'* * TextFileService (if imported).
Public Sub InitializeCommonService(Optional ByVal Force As Boolean = False)
    Call InitializeWorkbookService(Force)
    Call InitializeWorksheetService(Force)
    Call pTryInitializeOptionalCommonService("InitializeFileSystemService", Force)
    Call pTryInitializeOptionalCommonService("InitializeTextFileService", Force)
End Sub

Private Sub pTryInitializeOptionalCommonService(ByVal InitializerName As String, ByVal Force As Boolean)
    On Error Resume Next
    Err.Clear
    Call Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & InitializerName, Force)

    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num = 0 Then Exit Sub
    If err_num = 1004 And 0 < InStr(1, err_desc, InitializerName, vbTextCompare) Then Exit Sub

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

'* Expands multiple WorksheetRangeBounds objects to the largest range shape.
'*
'* @param ExpandRows Whether to expand row counts to the maximum row count.
'* @param ExpandColumns Whether to expand column counts to the maximum column count.
'* @param RangeBoundsList List of WorksheetRangeBounds objects to expand.
'* @return ObjectList holding WorksheetRangeBounds objects corresponding to the input order.
'*
'* @details
'* Calculates the maximum RowCount / ColumnCount across all input ranges and expands each range to the maximum dimensions only in the specified directions.
'* workbook / worksheet are not grouped; each returned range preserves the source workbook / worksheet from its input.
Public Function ExpandRangeBoundsToMax( _
        ByVal ExpandRows As Boolean, _
        ByVal ExpandColumns As Boolean, _
        ParamArray RangeBoundsList() As Variant) As ObjectList

    Dim result_value As ObjectList
    Set result_value = New_ObjectList("WorksheetRangeBounds")

    Dim max_row_count As Long
    Dim max_col_count As Long
    max_row_count = 0
    max_col_count = 0

    Dim item_idx As Long
    If UBound(RangeBoundsList) <> -1 Then
        For item_idx = LBound(RangeBoundsList) To UBound(RangeBoundsList)
            Dim target_bounds As WorksheetRangeBounds
            Set target_bounds = pGetRangeBoundsArgument(RangeBoundsList(item_idx), item_idx)

            If max_row_count < target_bounds.RowCount Then max_row_count = target_bounds.RowCount
            If max_col_count < target_bounds.ColumnCount Then max_col_count = target_bounds.ColumnCount
        Next item_idx

        For item_idx = LBound(RangeBoundsList) To UBound(RangeBoundsList)
            Set target_bounds = pGetRangeBoundsArgument(RangeBoundsList(item_idx), item_idx)

            Dim add_row As Long
            Dim add_col As Long
            add_row = 0
            add_col = 0
            If ExpandRows Then add_row = max_row_count - target_bounds.RowCount
            If ExpandColumns Then add_col = max_col_count - target_bounds.ColumnCount

            Call result_value.Add(target_bounds.Transform(AddRow:=add_row, AddColumn:=add_col))
        Next item_idx
    End If

    Set ExpandRangeBoundsToMax = result_value
End Function

Private Function pGetRangeBoundsArgument(ByVal RangeBoundsValue As Variant, ByVal ArgumentIndex As Long) As WorksheetRangeBounds
    If Not IsObject(RangeBoundsValue) Then
        Err.Raise vbObjectError + 1, "Function ExpandRangeBoundsToMax", "RangeBoundsList must contain WorksheetRangeBounds values. (" & CStr(ArgumentIndex) & ": " & TypeName(RangeBoundsValue) & ")"
    End If

    If RangeBoundsValue Is Nothing Then
        Err.Raise vbObjectError + 1, "Function ExpandRangeBoundsToMax", "RangeBoundsList cannot contain Nothing. (" & CStr(ArgumentIndex) & ")"
    End If

    If TypeName(RangeBoundsValue) <> "WorksheetRangeBounds" Then
        Err.Raise vbObjectError + 1, "Function ExpandRangeBoundsToMax", "RangeBoundsList must contain WorksheetRangeBounds values. (" & CStr(ArgumentIndex) & ": " & TypeName(RangeBoundsValue) & ")"
    End If

    Set pGetRangeBoundsArgument = RangeBoundsValue
End Function

' #############################################################################
'
' Error handling
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Error information
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Subroutine for processing error information and supplementing or accumulating information as needed.
'*
'* @param ErrNumber [Output] Variable that stores the error number. Default is 0. ErrNumber is overwritten even when Append is True. (Optional.)
'* @param ErrSource [Output] String variable that stores the error source. (Optional.)
'* @param ErrDescription [Output] String variable that stores the error description. (Optional.)
'* @param Supplementation Supplemental information to append to the error description. Default is an empty string.
'* @param ErrClear Whether to clear the error. When True, error information is cleared. Default is True.
'* @param Append Whether to join multiple pieces of error information. When True, information is appended to the existing information. Default is False.
'*
'* @details
'* This subroutine processes standard VBA error information in a customized form.
'* In particular, it can get the error number, source, and description, and append supplemental information.
'* When multiple errors occur, the Append parameter can be used to concatenate information.
'*
'* Usage example:
'* @code
'* Dim ErrNum As Long
'* Dim ErrSrc As String
'* Dim ErrDesc As String
'*
'* On Error Resume Next
'*
'* '// Example that intentionally raises an error (division by zero).
'* Debug.Print 1 / 0
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "Supplemental information 1", Append:=True)
'*
'* '// Example that intentionally raises an error (undefined sub procedure call).
'* Call NonExistentProcedure
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "Supplemental information 2", Append:=True)
'*
'* On Error GoTo 0
'*
'* Err.Raise ErrNum, ErrSrc, ErrDesc
'* @endcode
Public Sub HandleError( _
        Optional ByRef ErrNumber As Long = 0, _
        Optional ByRef ErrSource As String = "", _
        Optional ByRef ErrDescription As String = "", _
        Optional ByVal Supplementation As String = "", _
        Optional ByVal ErrClear As Boolean = True, _
        Optional ByVal Append As Boolean = False)

    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        If Append Then
            ErrSource = ErrSource & Err.Source & vbCrLf
            If Supplementation <> "" Then
                ErrDescription = ErrDescription & Err.Description & " [ " & Supplementation & " ]" & vbCrLf
            Else
                ErrDescription = ErrDescription & Err.Description & vbCrLf
            End If
        Else
            ErrSource = Err.Source
            If Supplementation <> "" Then
                ErrDescription = Err.Description & " [ " & Supplementation & " ]"
            Else
                ErrDescription = Err.Description
            End If
        End If

        If ErrClear Then
            Err.Clear
        End If
    End If
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel error values
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts an Excel error value to an Excel display string.
'*
'* @param ErrorValue Excel error value.
'* @return Excel error display string.
'*
'* @details
'* Converts an Excel error value represented by `CVErr(...)` to a string usable in cell display or formula literals.
'* Unknown error values are returned as `#ERRNO_xxx!` including the error number.
Public Function ExcelErrorToString(ByVal ErrorValue As Variant) As String
    If Not IsError(ErrorValue) Then
        Err.Raise vbObjectError + 1, "Function ExcelErrorToString", "The value is not an Excel error value. (" & TypeName(ErrorValue) & ")"
    End If

    Dim result_value As String
    Select Case ErrorValue
     Case CVErr(xlErrDiv0)
        result_value = "#DIV/0!"
     Case CVErr(xlErrNA)
        result_value = "#N/A"
     Case CVErr(xlErrName)
        result_value = "#NAME?"
     Case CVErr(xlErrNull)
        result_value = "#NULL!"
     Case CVErr(xlErrNum)
        result_value = "#NUM!"
     Case CVErr(xlErrRef)
        result_value = "#REF!"
     Case CVErr(xlErrValue)
        result_value = "#VALUE!"
     Case CVErr(xlErrGettingData)
        result_value = "#GETTING_DATA"
     Case CVErr(xlErrSpill)
        result_value = "#SPILL!"
     Case CVErr(xlErrConnect)
        result_value = "#CONNECT!"
     Case CVErr(xlErrBlocked)
        result_value = "#BLOCKED!"
     Case CVErr(xlErrUnknown)
        result_value = "#UNKNOWN!"
     Case CVErr(xlErrField)
        result_value = "#FIELD!"
     Case CVErr(xlErrCalc)
        result_value = "#CALC!"
     Case Else
        result_value = "#ERRNO_" & CStr(CLng(ErrorValue)) & "!"
    End Select

    ExcelErrorToString = result_value
End Function

'* Converts an Excel error display string to an Excel error value.
'*
'* @param Expression Excel error display string.
'* @param ErrorValue [Output] Excel error value when conversion succeeds.
'* @return True when conversion succeeds; otherwise, False.
'*
'* @details
'* Converts only known Excel error display strings to `CVErr(...)`. Case-insensitive.
'* For unknown strings, explicitly assigns `Empty` to `ErrorValue` and returns False.
Public Function TryConvertExcelErrorStringToCVErr( _
        ByVal Expression As String, _
        ByRef ErrorValue As Variant) As Boolean

    Dim result_value As Boolean
    result_value = True

    Select Case UCase$(Expression)
     Case "#DIV/0!"
        ErrorValue = CVErr(xlErrDiv0)
     Case "#N/A"
        ErrorValue = CVErr(xlErrNA)
     Case "#NAME?"
        ErrorValue = CVErr(xlErrName)
     Case "#NULL!"
        ErrorValue = CVErr(xlErrNull)
     Case "#NUM!"
        ErrorValue = CVErr(xlErrNum)
     Case "#REF!"
        ErrorValue = CVErr(xlErrRef)
     Case "#VALUE!"
        ErrorValue = CVErr(xlErrValue)
     Case "#GETTING_DATA"
        ErrorValue = CVErr(xlErrGettingData)
     Case "#SPILL!"
        ErrorValue = CVErr(xlErrSpill)
     Case "#CONNECT!"
        ErrorValue = CVErr(xlErrConnect)
     Case "#BLOCKED!"
        ErrorValue = CVErr(xlErrBlocked)
     Case "#UNKNOWN!"
        ErrorValue = CVErr(xlErrUnknown)
     Case "#FIELD!"
        ErrorValue = CVErr(xlErrField)
     Case "#CALC!"
        ErrorValue = CVErr(xlErrCalc)
     Case Else
        ErrorValue = Empty
        result_value = False
    End Select

    TryConvertExcelErrorStringToCVErr = result_value
End Function


' #############################################################################
'
' GUI
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Buttons
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Creates a button on a worksheet.
'*
'* @param TargetSheet Target worksheet.
'* @param RowIndex Row number of the cell where the button is added.
'* @param ColumnIndex Column number of the cell where the button is added.
'* @param Caption Display text of the button.
'* @param OnAction Sub procedure to run when clicked.
'* @param OnAction [Optional] Button component name (Name property).
'* @return Added button object.
'*
'* @details
'* Creates a button on a worksheet.
Public Function AddButton( _
    ByVal TargetSheet As Worksheet, _
    ByVal RowIndex As Long, _
    ByVal ColumnIndex As Long, _
    ByVal Caption As String, _
    ByVal OnAction As String, _
    Optional ByVal Name As String = "") As Shape

    Dim tgt_rng As Range
    Set tgt_rng = TargetSheet.Cells(RowIndex, ColumnIndex)

    Dim shp_btn As Shape
    Set shp_btn = TargetSheet.Shapes.AddShape( _
                 Type:=msoShapeRectangle, _
                 Left:=tgt_rng.Left, _
                 Top:=tgt_rng.Top, _
                 Width:=tgt_rng.Width, _
                 Height:=tgt_rng.Height)

    If Name <> "" Then
        shp_btn.Name = Name
    End If

    shp_btn.OnAction = OnAction

    With shp_btn.TextFrame2
        .TextRange.Text = Caption

        With .TextRange.Font
            .NameComplexScript = "Yu Gothic UI Semibold"
            .NameFarEast = "Yu Gothic UI Semibold"
            .Name = "Yu Gothic UI Semibold"
            .Size = 9
            .Fill.ForeColor.RGB = RGB(64, 64, 64)
        End With

        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With

    shp_btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp_btn.TextFrame.VerticalAlignment = xlVAlignCenter

    shp_btn.Fill.ForeColor.RGB = RGB(192, 192, 192)
    shp_btn.Fill.Solid
    shp_btn.Line.ForeColor.RGB = RGB(64, 64, 64)
    shp_btn.Line.Weight = 1

    shp_btn.Shadow.Visible = msoFalse
    shp_btn.ThreeD.Visible = False

    Set AddButton = shp_btn
End Function

'* Deletes a worksheet button.
'*
'* @param TargetSheet Target worksheet.
'* @param Name Button component name (Name property).
'*
'* @details
'* Deletes a button on a worksheet by specifying its component name.
Public Sub DeleteButton(ByVal TargetSheet As Worksheet, ByVal Name As String)

    Dim shape_obj As Shape
    Dim is_deleted As Boolean: is_deleted = False

    For Each shape_obj In TargetSheet.Shapes
        If shape_obj.Name = Name Then
            shape_obj.Delete
            is_deleted = True
            Exit For
        End If
    Next shape_obj

    If Not is_deleted Then
        Err.Raise vbObjectError + 1, "Sub DeleteButton", "The button was not found. (" & Name & ")"
    End If

End Sub

'* Deletes all buttons on a worksheet.
'*
'* @param TargetSheet Target worksheet.
'*
'* @details
'* Deletes all buttons on a worksheet.
Public Sub ClearButton(ByVal TargetSheet As Worksheet)
    Dim item_idx As Long
    ' Deleting from the end is safer because issues can occur if the collection changes during the loop.
    ' The For i = Shapes.Count To 1 Step -1 form is safe.
    For item_idx = TargetSheet.Shapes.Count To 1 Step -1

        If TargetSheet.Shapes(item_idx).OnAction <> "" Then
            ' If OnAction is set, treat it as a shape button and delete it.
            TargetSheet.Shapes(item_idx).Delete
        End If

    Next item_idx

End Sub


' #############################################################################
'
' Clipboard
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Clipboard operations
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Sends a string to the clipboard.
'*
'* @param SourceText String to copy to the clipboard.
'*
'* @details
'* Copies the specified string to the clipboard using the VBA `Forms.TextBox.1` object.
Public Sub SetClipboard(ByVal SourceText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = SourceText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

'* Gets a string from the clipboard.
'*
'* @return String stored in the clipboard. Returns an empty string if the clipboard does not contain a string.
'*
'* @details
'* Gets the clipboard contents using the VBA `Forms.TextBox.1` object.
Public Function GetClipboard() As String
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        If .CanPaste = True Then .Paste
        GetClipboard = .Text
    End With
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Paste formulas
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Performs paste while removing formatting as much as possible.
'*
'* @details
'* Attempts to paste clipboard contents in the following order:
'* 1. Paste as formulas (`Paste:=xlPasteFormulas`).
'* 2. Paste text only (`NoHTMLFormatting:=True`).
'* 3. If both fail, perform a standard paste.
Public Sub PasteFormulas()
    On Error Resume Next

    Call Selection.PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    Call ActiveSheet.PasteSpecial(Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True)
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    Call ActiveSheet.Paste
End Sub


' #############################################################################
'
' File operations
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' File-system path strings
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Replaces special characters in a path with full-width characters.
'*
'* @param Path Path string to replace.
'* @return Path string with special characters replaced by full-width characters.
'*
'* @details
'* Replaces special characters that cannot be used in the file system with full-width characters.
'* The following characters are replaced: `\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`.
Public Function ReplaceSpecialCharacterOnFileSystemPath(ByVal Path As String) As String
    Dim result_value As String

    result_value = Path

    result_value = Replace(result_value, "\", "Å_")
    result_value = Replace(result_value, "/", "Å^")
    result_value = Replace(result_value, ":", "ÅF")
    result_value = Replace(result_value, "*", "Åñ")
    result_value = Replace(result_value, "?", "ÅH")
    result_value = Replace(result_value, """", "Åç")
    result_value = Replace(result_value, "<", "ÅÉ")
    result_value = Replace(result_value, ">", "ÅÑ")
    result_value = Replace(result_value, "|", "Åb")

    ReplaceSpecialCharacterOnFileSystemPath = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Path combining
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Joins paths.
'*
'* @param Path1 First path string.
'* @param Path2 Second path string to join.
'* @param Paths Additional variable-length path strings.
'* @return Joined path string.
'*
'* @details
'* Joins multiple specified strings with the path separator.
'* Additional path strings can be passed as variable-length arguments.
'*
'* @note
'* JoinPath is a string-joining API. Even when an absolute path or URL is passed to Path2 or later, it does not replace the preceding path and is joined as a normal path element.
'* When Path1 is a URL, Path2 and later are normalized to URL separators, but the query / fragment on the Path1 side is not moved.
'* Therefore, if Path1 contains ? or #, subsequent paths are joined after the URL suffix.
Public Function JoinPath(ByVal Path1 As String, ByVal Path2 As String, ParamArray Paths() As Variant) As String
    Dim result_value As String

    result_value = pJoinPathCore(Path1, Path2)

    If UBound(Paths) = -1 Then
        JoinPath = result_value
        Exit Function
    End If

    Dim item_idx As Long
    For item_idx = LBound(Paths) To UBound(Paths)
        result_value = pJoinPathCore(result_value, Paths(item_idx))
    Next item_idx

    JoinPath = result_value
End Function

Private Function pJoinPathCore(ByVal Path1 As String, ByVal Path2 As String) As String
    If IsUrlPath(Path1) Then
        pJoinPathCore = pJoinPathCoreWithSep(Path1, pNormalizeUrlRelativePathForJoin(Path2), "/")
    Else
        pJoinPathCore = pJoinPathCoreWithSep(pNormalizeWindowsPathSeparators(Path1), pNormalizeWindowsPathSeparators(Path2), G_FS_PATH_SEP)
    End If
End Function

Private Function pJoinPathCoreWithSep(ByVal Path1 As String, ByVal Path2 As String, ByVal PathSep As String) As String
    If EndsWith(Path1, PathSep) Then
        If StartsWith(Path2, PathSep) Then
            pJoinPathCoreWithSep = Left$(Path1, Len(Path1) - 1) & Path2
        Else
            pJoinPathCoreWithSep = Path1 & Path2
        End If
    Else
        If StartsWith(Path2, PathSep) Then
            pJoinPathCoreWithSep = Path1 & Path2
        Else
            pJoinPathCoreWithSep = Path1 & PathSep & Path2
        End If
    End If
End Function

Private Function pNormalizeUrlRelativePathForJoin(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pNormalizeUrlRelativePathForJoin = pNormalizePathSeparators(path_part) & path_suffix
End Function

Private Sub pSplitPathSuffix(ByRef PathPart As String, ByRef PathSuffix As String, ByVal PathWithSuffix As String)
    Dim query_pos As Long
    query_pos = InStr(1, PathWithSuffix, "?", vbBinaryCompare)

    Dim fragment_pos As Long
    fragment_pos = InStr(1, PathWithSuffix, "#", vbBinaryCompare)

    Dim suffix_pos As Long
    If query_pos > 0 And fragment_pos > 0 Then
        suffix_pos = IIf(query_pos < fragment_pos, query_pos, fragment_pos)
    ElseIf query_pos > 0 Then
        suffix_pos = query_pos
    Else
        suffix_pos = fragment_pos
    End If

    If suffix_pos = 0 Then
        PathPart = PathWithSuffix
        PathSuffix = ""
    Else
        PathPart = Left$(PathWithSuffix, suffix_pos - 1)
        PathSuffix = Mid$(PathWithSuffix, suffix_pos)
    End If
End Sub

Private Function pNormalizePathSeparators(ByVal PathLikeString As String) As String
    pNormalizePathSeparators = Replace(PathLikeString, G_FS_PATH_SEP, "/")
End Function

Private Function pNormalizeWindowsPathSeparators(ByVal PathLikeString As String) As String
    pNormalizeWindowsPathSeparators = Replace(PathLikeString, "/", G_FS_PATH_SEP)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Convert to absolute paths
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a path to an absolute path using a base directory.
'*
'* @param ParentAbsolutePath Absolute directory path to use as the base.
'* @param PathLikeString Path string to convert to an absolute path.
'* @return Absolute path.
'*
'* @details
'* When PathLikeString is a relative path, converts it to an absolute path based on ParentAbsolutePath.
'* When PathLikeString is an absolute path, does not use ParentAbsolutePath and returns a normalized absolute path.
'* Root-relative paths such as `\foo` are based on the root of ParentAbsolutePath.
'* Drive-relative paths such as `C:foo` are current-directory-dependent, so they are errors.
'* URLs are not passed to FSO; only the path part is normalized with `/` separators.
'*
'* @note
'* When PathLikeString is an absolute path or URL, does not use ParentAbsolutePath and returns the normalized path.
'* When PathLikeString is a relative path, follows the same joining rules as JoinPath.
'* Therefore, if ParentAbsolutePath contains a URL suffix, subsequent relative paths are joined after the suffix.
Public Function GetAbsolutePathFromParent(ByVal ParentAbsolutePath As String, ByVal PathLikeString As String) As String
    If IsDriveRelativePath(ParentAbsolutePath) Or Not IsAbsolutePath(ParentAbsolutePath) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "ParentAbsolutePath must be an absolute path. (" & ParentAbsolutePath & ")"
    If IsDriveRelativePath(PathLikeString) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "Drive-relative paths are not allowed. (" & PathLikeString & ")"

    Dim parent_path As String
    parent_path = pNormalizeAbsolutePath(ParentAbsolutePath)

    Dim target_path As String
    If PathLikeString = "" Or PathLikeString = "." Or PathLikeString = ".\" Or PathLikeString = "./" Then
        target_path = parent_path
    ElseIf IsAbsolutePath(PathLikeString) Then
        target_path = PathLikeString
    ElseIf pIsRootRelativePath(PathLikeString) Then
        target_path = JoinPath(GetPathRoot(parent_path), Mid$(pNormalizePathPartSeparators(PathLikeString), 2) & pGetPathSuffix(PathLikeString))
    Else
        target_path = JoinPath(parent_path, PathLikeString)
    End If

    GetAbsolutePathFromParent = pNormalizeAbsolutePath(target_path)
End Function

Private Function pNormalizeAbsolutePath(ByVal PathLikeString As String) As String
    If IsUrlPath(PathLikeString) Then
        pNormalizeAbsolutePath = pNormalizeUrlPath(PathLikeString)
    Else
        pNormalizeAbsolutePath = pNormalizeLocalAbsolutePath(PathLikeString)
    End If
End Function

Private Function pNormalizeUrlPath(ByVal UrlPath As String) As String
    Dim url_prefix As String
    Dim url_path As String
    Dim url_suffix As String
    Call pSplitUrl(url_prefix, url_path, url_suffix, UrlPath)
    pNormalizeUrlPath = url_prefix & pNormalizeSlashPath(url_path, True) & url_suffix
End Function

Private Sub pSplitUrl(ByRef UrlPrefix As String, ByRef UrlPathPart As String, ByRef UrlSuffix As String, ByVal UrlPath As String)
    Dim scheme_sep_pos As Long
    scheme_sep_pos = InStr(1, UrlPath, "://", vbBinaryCompare)
    If scheme_sep_pos = 0 Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "The value is not in URL format. (" & UrlPath & ")"

    Dim scan_start As Long
    scan_start = scheme_sep_pos + Len("://")

    Dim path_start As Long
    Dim suffix_start As Long
    Dim char_idx As Long
    For char_idx = scan_start To Len(UrlPath)
        Dim current_char As String
        current_char = Mid$(UrlPath, char_idx, 1)
        If current_char = "/" Or current_char = G_FS_PATH_SEP Then
            path_start = char_idx
            Exit For
        ElseIf current_char = "?" Or current_char = "#" Then
            suffix_start = char_idx
            Exit For
        End If
    Next char_idx

    If path_start = 0 Then
        If suffix_start = 0 Then
            UrlPrefix = UrlPath
            UrlSuffix = ""
        Else
            UrlPrefix = Left$(UrlPath, suffix_start - 1)
            UrlSuffix = Mid$(UrlPath, suffix_start)
        End If
        UrlPathPart = "/"
        Exit Sub
    End If

    UrlPrefix = Left$(UrlPath, path_start - 1)
    Call pSplitPathSuffix(UrlPathPart, UrlSuffix, Mid$(UrlPath, path_start))
    UrlPathPart = pNormalizePathSeparators(UrlPathPart)
End Sub

Private Function pNormalizeLocalAbsolutePath(ByVal PathLikeString As String) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(PathLikeString)

    If IsDriveRelativePath(slash_path) Then Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "Drive-relative paths are not allowed. (" & PathLikeString & ")"

    If IsDriveAbsolutePath(slash_path) Then
        Dim drive_root As String
        drive_root = UCase$(Left$(slash_path, 1)) & ":/"

        Dim drive_rest As String
        drive_rest = Mid$(slash_path, Len(drive_root) + 1)

        Dim drive_normalized_path As String
        drive_normalized_path = pNormalizeSlashPath("/" & drive_rest, True)
        pNormalizeLocalAbsolutePath = pNormalizeWindowsPathSeparators(drive_root & Mid$(drive_normalized_path, 2))
    ElseIf IsUncPath(slash_path) Then
        Dim unc_root As String
        unc_root = pGetUncPathRootSlash(slash_path)

        Dim unc_rest As String
        unc_rest = Mid$(slash_path, Len(unc_root) + 1)

        Dim unc_normalized_path As String
        unc_normalized_path = pNormalizeSlashPath("/" & unc_rest, True)
        pNormalizeLocalAbsolutePath = pNormalizeWindowsPathSeparators(unc_root & Mid$(unc_normalized_path, 2))
    Else
        Err.Raise vbObjectError + 1, "Function GetAbsolutePathFromParent", "Specify an absolute path. (" & PathLikeString & ")"
    End If
End Function

Private Function pNormalizeSlashPath(ByVal PathPart As String, ByVal IsAbsolute As Boolean) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(PathPart)

    Dim path_parts As Variant
    path_parts = Split(slash_path, "/")

    Dim stack_items() As String
    Dim stack_count As Long
    stack_count = 0

    Dim path_part As Variant
    For Each path_part In path_parts
        Dim part_text As String
        part_text = CStr(path_part)

        If part_text = "" Or part_text = "." Then
            ' no operation
        ElseIf part_text = ".." Then
            If stack_count > 0 Then
                stack_count = stack_count - 1
            ElseIf Not IsAbsolute Then
                Call pPushPathSegment(stack_items, stack_count, part_text)
            End If
        Else
            Call pPushPathSegment(stack_items, stack_count, part_text)
        End If
    Next path_part

    Dim result_path As String
    If IsAbsolute Then result_path = "/"

    Dim item_idx As Long
    For item_idx = 0 To stack_count - 1
        If result_path <> "" And Not EndsWith(result_path, "/") Then result_path = result_path & "/"
        result_path = result_path & stack_items(item_idx)
    Next item_idx

    If result_path = "" Then
        If IsAbsolute Then
            result_path = "/"
        Else
            result_path = "."
        End If
    End If

    pNormalizeSlashPath = result_path
End Function

Private Sub pPushPathSegment(ByRef StackItems() As String, ByRef StackCount As Long, ByVal PathSegment As String)
    If StackCount = 0 Then
        ReDim StackItems(0 To 0)
    Else
        ReDim Preserve StackItems(0 To StackCount)
    End If

    StackItems(StackCount) = PathSegment
    StackCount = StackCount + 1
End Sub

Private Function pIsRootRelativePath(ByVal PathWithSuffix As String) As Boolean
    pIsRootRelativePath = StartsWith(pNormalizePathPartSeparators(PathWithSuffix), "/")
End Function

Private Function pNormalizePathPartSeparators(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pNormalizePathPartSeparators = pNormalizePathSeparators(path_part)
End Function

Private Function pGetPathSuffix(ByVal PathWithSuffix As String) As String
    Dim path_part As String
    Dim path_suffix As String
    Call pSplitPathSuffix(path_part, path_suffix, PathWithSuffix)
    pGetPathSuffix = path_suffix
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Get path roots
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Gets the root part of a path.
'*
'* @param PathLikeString Path string whose root is retrieved.
'* @return Root in the form `C:\` for drive paths, `\\server\share\` for UNC paths, or `scheme://authority/` for URLs.
'*
'* @details
'* If the first character is one alphabetic character and the second character is `:`, the path is treated as a drive path.
'* If the first two characters are `\` or `//`, the path is treated as a UNC path.
'* URLs in `scheme://` format are also treated as absolute paths.
'* Other path strings raise an error.
Public Function GetPathRoot(ByVal PathLikeString As String) As String
    If IsUrlPath(PathLikeString) Then
        Dim url_prefix As String
        Dim url_path As String
        Dim url_suffix As String
        Call pSplitUrl(url_prefix, url_path, url_suffix, PathLikeString)
        GetPathRoot = url_prefix & "/"
    ElseIf IsDriveRelativePath(PathLikeString) Then
        Err.Raise vbObjectError + 1, "Function GetPathRoot", "Cannot resolve the root of a drive-relative path. (" & PathLikeString & ")"
    ElseIf IsDriveAbsolutePath(PathLikeString) Then
        GetPathRoot = UCase$(Left$(PathLikeString, 1)) & ":" & G_FS_PATH_SEP
    ElseIf IsUncPath(PathLikeString) Then
        GetPathRoot = pNormalizeWindowsPathSeparators(pGetUncPathRootSlash(PathLikeString))
    Else
        Err.Raise vbObjectError + 1, "Function GetPathRoot", "Specify a drive path, UNC path, or URL. (" & PathLikeString & ")"
    End If
End Function

Private Function pGetUncPathRootSlash(ByVal UncPath As String) As String
    Dim slash_path As String
    slash_path = pNormalizePathSeparators(UncPath)

    Dim server_sep As Long
    server_sep = InStr(3, slash_path, "/")
    If server_sep <= 3 Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "Cannot resolve the UNC path root. (" & UncPath & ")"

    Dim share_sep As Long
    share_sep = InStr(server_sep + 1, slash_path, "/")

    Dim share_name As String
    If share_sep = 0 Then
        share_name = Mid$(slash_path, server_sep + 1)
        If share_name = "" Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "Cannot resolve the UNC path root. (" & UncPath & ")"

        pGetUncPathRootSlash = slash_path
        If Not EndsWith(pGetUncPathRootSlash, "/") Then pGetUncPathRootSlash = pGetUncPathRootSlash & "/"
    Else
        share_name = Mid$(slash_path, server_sep + 1, share_sep - server_sep - 1)
        If share_name = "" Then Err.Raise vbObjectError + 1, "Function GetPathRoot", "Cannot resolve the UNC path root. (" & UncPath & ")"

        pGetUncPathRootSlash = Left$(slash_path, share_sep)
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Path type checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Determines whether a path has a drive specifier.
'*
'* @param TestPath Path string to test.
'* @return True if the path has a drive specifier; otherwise, False.
'*
'* @details
'* Returns True when the path starts with a drive specifier, such as `C:\foo`, `C:/foo`, `C:foo`, or `C:`.
Public Function IsDrivePath(ByVal TestPath As String) As Boolean
    If Len(TestPath) < 2 Then Exit Function
    If Mid$(TestPath, 2, 1) <> ":" Then Exit Function
    IsDrivePath = pIsAsciiAlphabet(Left$(TestPath, 1))
End Function

Private Function pIsAsciiAlphabet(ByVal TestChar As String) As Boolean
    If Len(TestChar) <> 1 Then Exit Function
    pIsAsciiAlphabet = ("A" <= UCase$(TestChar) And UCase$(TestChar) <= "Z")
End Function

'* Determines whether a path is a drive-absolute path.
'*
'* @param TestPath Path string to test.
'* @return True if the path is a drive-absolute path; otherwise, False.
'*
'* @details
'* Returns True for drive-absolute paths such as `C:\foo` or `C:/foo`. Drive-relative paths such as `C:foo` are False.
Public Function IsDriveAbsolutePath(ByVal TestPath As String) As Boolean
    If Not IsDrivePath(TestPath) Then Exit Function
    If Len(TestPath) < 3 Then Exit Function
    IsDriveAbsolutePath = (Mid$(TestPath, 3, 1) = G_FS_PATH_SEP Or Mid$(TestPath, 3, 1) = "/")
End Function

'* Determines whether a path is a drive-relative path.
'*
'* @param TestPath Path string to test.
'* @return True if the path is a drive-relative path; otherwise, False.
'*
'* @details
'* Returns True for paths that have a drive specifier and do not start with a separator, such as `C:foo` or `C:`.
Public Function IsDriveRelativePath(ByVal TestPath As String) As Boolean
    If Not IsDrivePath(TestPath) Then Exit Function
    IsDriveRelativePath = Not IsDriveAbsolutePath(TestPath)
End Function

'* Determines whether a path is a UNC path.
'*
'* @param TestPath Path string to test.
'* @return True if the path is a UNC path; otherwise, False.
'*
'* @details
'* Returns True for UNC paths such as `\\server\share` or `//server/share`.
Public Function IsUncPath(ByVal TestPath As String) As Boolean
    If Not (Left$(TestPath, 2) = "\\" Or Left$(TestPath, 2) = "//") Then Exit Function

    On Error Resume Next
    Err.Clear
    Dim root_path As String
    root_path = pGetUncPathRootSlash(TestPath)
    IsUncPath = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'* Determines whether a path is in URL format.
'*
'* @param TestPath Path string to test.
'* @return True if the path is in URL format; otherwise, False.
'*
'* @details
'* Treats `scheme://` format as a URL. Scheme names start with a letter and can contain letters, digits, `+`, `.`, and `-`.
Public Function IsUrlPath(ByVal TestPath As String) As Boolean
    Dim scheme_sep_pos As Long
    scheme_sep_pos = InStr(1, TestPath, "://", vbBinaryCompare)
    If scheme_sep_pos <= 1 Then Exit Function

    Dim scheme_name As String
    scheme_name = Left$(TestPath, scheme_sep_pos - 1)
    If Not pIsAsciiAlphabet(Left$(scheme_name, 1)) Then Exit Function

    Dim char_idx As Long
    For char_idx = 2 To Len(scheme_name)
        Dim scheme_char As String
        scheme_char = Mid$(scheme_name, char_idx, 1)
        If Not (pIsAsciiAlphabet(scheme_char) Or pIsAsciiDigit(scheme_char) Or scheme_char = "+" Or scheme_char = "." Or scheme_char = "-") Then Exit Function
    Next char_idx

    IsUrlPath = True
End Function

Private Function pIsAsciiDigit(ByVal TestChar As String) As Boolean
    If Len(TestChar) <> 1 Then Exit Function
    pIsAsciiDigit = ("0" <= TestChar And TestChar <= "9")
End Function

'* Determines whether a path string is an absolute path.
'*
'* @param TestPath Path string to test.
'* @return True if the path is an absolute path; otherwise, False.
'*
'* @details
'* Determines whether the input path string is an absolute path.
'* URL-format strings are also treated as absolute paths.
Public Function IsAbsolutePath(ByVal TestPath As String) As Boolean
    If IsDriveRelativePath(TestPath) Then Exit Function

    Err.Clear
    On Error Resume Next
    Call GetPathRoot(TestPath)
    IsAbsolutePath = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Path splitting
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Splits a path into parent path and leaf path.
'*
'* @param ParentPath [Output] Parent path.
'* @param LeafPath [Output] Leaf path.
'* @param Path Input path string.
'* @param IgnoreEndSep [Optional] Whether to ignore the trailing separator in the input path string.
'*
'* @details
'* Splits Path into parent path and leaf path at the last path separator.
'* For URLs, excludes separators in query / fragment from split targets and returns PathSuffix as part of LeafPath.
Public Sub SplitPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal Path As String, _
        Optional ByVal IgnoreEndSep As Boolean = False)

    If IsUrlPath(Path) Then
        Call pSplitUrlPath(ParentPath, LeafPath, Path, IgnoreEndSep)
    Else
        Call pSplitWindowsPath(ParentPath, LeafPath, Path, IgnoreEndSep)
    End If
End Sub

Private Sub pSplitUrlPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal UrlPath As String, _
        ByVal IgnoreEndSep As Boolean)

    Dim url_prefix As String
    Dim url_path As String
    Dim url_suffix As String
    Call pSplitUrl(url_prefix, url_path, url_suffix, UrlPath)

    Dim normalized_path As String
    normalized_path = pNormalizeSlashPath(url_path, True)
    If EndsWith(url_path, "/") And normalized_path <> "/" Then normalized_path = normalized_path & "/"
    url_path = normalized_path

    If EndsWith(url_path, "/") And url_path <> "/" Then
        url_path = Left$(url_path, Len(url_path) - 1)
        If Not IgnoreEndSep Then
            ParentPath = url_prefix & url_path
            LeafPath = url_suffix
            Exit Sub
        End If
    End If

    Dim last_sep As Long
    last_sep = InStrRev(url_path, "/")
    If last_sep <= 1 Then
        ParentPath = url_prefix & "/"
        LeafPath = Mid$(url_path, 2) & url_suffix
    Else
        ParentPath = url_prefix & Left$(url_path, last_sep - 1)
        LeafPath = Mid$(url_path, last_sep + 1) & url_suffix
    End If
End Sub

Private Sub pSplitWindowsPath( _
        ByRef ParentPath As String, _
        ByRef LeafPath As String, _
        ByVal WindowsPath As String, _
        ByVal IgnoreEndSep As Boolean)

    WindowsPath = pNormalizeWindowsPathSeparators(WindowsPath)

    Dim root_path As String
    root_path = pGetWindowsPathRootForSplit(WindowsPath)

    If root_path <> "" And StrComp(WindowsPath, root_path, vbTextCompare) = 0 Then
        If IgnoreEndSep Then
            ParentPath = ""
            LeafPath = WindowsPath
        Else
            ParentPath = WindowsPath
            LeafPath = ""
        End If
        Exit Sub
    End If

    If EndsWith(WindowsPath, G_FS_PATH_SEP) Then
        WindowsPath = Left$(WindowsPath, Len(WindowsPath) - 1)
        If Not IgnoreEndSep Then
            ParentPath = WindowsPath
            LeafPath = ""
            Exit Sub
        End If
    End If

    Dim last_sep As Long
    last_sep = InStrRev(WindowsPath, G_FS_PATH_SEP)
    If last_sep = 0 Then
        ParentPath = ""
        LeafPath = WindowsPath
    ElseIf root_path <> "" And last_sep = Len(root_path) Then
        ParentPath = root_path
        LeafPath = Mid$(WindowsPath, last_sep + 1)
    Else
        ParentPath = Left$(WindowsPath, last_sep - 1)
        LeafPath = Mid$(WindowsPath, last_sep + 1)
    End If
End Sub

Private Function pGetWindowsPathRootForSplit(ByVal WindowsPath As String) As String
    If IsDriveAbsolutePath(WindowsPath) Then
        pGetWindowsPathRootForSplit = Left$(WindowsPath, 3)
    ElseIf IsUncPath(WindowsPath) Then
        pGetWindowsPathRootForSplit = pNormalizeWindowsPathSeparators(pGetUncPathRootSlash(WindowsPath))
    End If
End Function

'* Parses a leaf path into base name, extension, and path suffix.
'*
'* @param BaseName [Output] Leaf name excluding extension and path suffix.
'* @param Extension [Output] Extension including `.`. Empty string when there is no extension.
'* @param PathSuffix [Output] Path suffix including query / fragment.
'* @param LeafPath Input leaf path string.
'* @param AsUrl [Optional] Whether to treat LeafPath as a URL leaf path and separate query / fragment. Default is False.
'*
'* @details
'* By default, treats `?` and `#` as part of the leaf name and splits at the last `.`.
'* Only when AsUrl is True, treats `?` or `#` and later as PathSuffix and splits the preceding string at the last `.`.
Public Sub ParseLeafPath( _
        ByRef BaseName As String, _
        ByRef Extension As String, _
        ByRef PathSuffix As String, _
        ByVal LeafPath As String, _
        Optional ByVal AsUrl As Boolean = False)

    Dim leaf_body As String
    If AsUrl Then
        Call pSplitPathSuffix(leaf_body, PathSuffix, LeafPath)
    Else
        leaf_body = LeafPath
        PathSuffix = ""
    End If

    Dim last_period As Long
    last_period = InStrRev(leaf_body, ".")

    If 1 < last_period Then
        BaseName = Left$(leaf_body, last_period - 1)
        Extension = Mid$(leaf_body, last_period)
    Else
        BaseName = leaf_body
        Extension = ""
    End If
End Sub

'* Gets the path excluding the last part of the path.
'*
'* @param Path Input path string.
'* @param IgnoreEndSep [Optional] Whether to ignore the trailing separator in the input path string.
'* @return Path string excluding the last part.
'*
'* @details
'* Returns the portion before the last path separator in the input path string.
'* For example, for `Path\to\File`, returns `Path\to`.
Public Function GetParentPath(ByVal Path As String, Optional ByVal IgnoreEndSep As Boolean = False) As String
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, Path, IgnoreEndSep:=IgnoreEndSep)

    GetParentPath = parent_path
End Function

'* Gets the last part of a path.
'*
'* @param Path Input path string.
'* @param BaseName [Optional] Whether to include the base name. Default is True.
'* @param Extension [Optional] Whether to include the extension. Default is True.
'* @param IgnoreEndSep [Optional] Whether to ignore the trailing separator in the input path string.
'* @return Last part (file name or folder name).
'*
'* @details
'* Returns the portion after the last path separator in the input path string.
'* When the optional argument `Extension` is False, returns the result with the extension removed.
'* URL PathSuffix is not included in the return value. If PathSuffix is needed, use SplitPath and ParseLeafPath.
'* For example, for `Path\to\File.txt`, returns `File`.
Public Function GetLeafFromPath(ByVal Path As String, Optional ByVal BaseName As Boolean = True, Optional ByVal Extension As Boolean = True, Optional ByVal IgnoreEndSep As Boolean = False) As String
    If Not BaseName And Not Extension Then Exit Function

    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, Path, IgnoreEndSep:=IgnoreEndSep)

    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, leaf_path, AsUrl:=IsUrlPath(Path))

    If BaseName Then GetLeafFromPath = GetLeafFromPath & base_name
    If Extension Then GetLeafFromPath = GetLeafFromPath & file_ext
End Function


' #############################################################################
'
' Date operations
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Date checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Determines whether a date is unset.
'*
'* @param TestDate Date to test.
'* @return True if the date is unset; otherwise, False.
'*
'* @details
'* Returns True when the specified date is less than `G_DATE_MINIMUM` or equal to `G_DATE_NULL`.
'* `G_DATE_NULL` represents an unset date, and `G_DATE_MINIMUM` represents the minimum valid date.
Public Function IsNullDate(ByVal TestDate As Date) As Boolean
    If TestDate < G_DATE_MINIMUM Then
        IsNullDate = True
    Else
        IsNullDate = False
    End If
End Function


' #############################################################################
'
' Conversions
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' String/array conversions
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a string to an array of characters.
'*
'* @param Expression String.
'* @return Array of characters.
'*
'* @details
'* Converts a string to an array of characters.
Public Function ConvertStringToCharArray(ByVal Expression As String) As String()
    Dim result_ubound As Long
    result_ubound = Len(Expression) - 1

    If result_ubound < 0 Then
        ConvertStringToCharArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To result_ubound)

    Dim char_idx As Long
    For char_idx = 0 To result_ubound
        result_value(char_idx) = Mid(Expression, char_idx + 1, 1)
    Next char_idx

    ConvertStringToCharArray = result_value
End Function

'* Converts a String() array to a Variant() array.
'*
'* @param StringArray String array to convert.
'* @return Variant array.
'*
'* @details
'* Converts a String array to Variant. The contents of each element are preserved.
'* A 0-element array is returned as a 0-element Variant array.
Public Function ConvertArrayStringToVariant(ByRef StringArray() As String) As Variant()
    If IsEmptyArray(StringArray) Then
        ConvertArrayStringToVariant = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    Dim item_idx As Long

    ReDim result_value(LBound(StringArray) To UBound(StringArray))

    For item_idx = LBound(StringArray) To UBound(StringArray)
        result_value(item_idx) = StringArray(item_idx)
    Next item_idx

    ConvertArrayStringToVariant = result_value
End Function

'* Converts a Variant() array to a String() array.
'*
'* @param VariantArray Variant array to convert.
'* @return String array.
'*
'* @details
'* Converts a Variant array to String. The contents of each element are preserved.
'* A 0-element array is returned as a 0-element String array.
Public Function ConvertArrayVariantToString(ByRef VariantArray() As Variant) As String()
    If IsEmptyArray(VariantArray) Then
        ConvertArrayVariantToString = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    Dim item_idx As Long

    ReDim result_value(LBound(VariantArray) To UBound(VariantArray))

    For item_idx = LBound(VariantArray) To UBound(VariantArray)
        result_value(item_idx) = VariantArray(item_idx)
    Next item_idx

    ConvertArrayVariantToString = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Boolean conversions
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a Boolean value to a string.
'*
'* @param BooleanValue Boolean value to convert.
'* @param FlagOnString String returned for True (default is "Å°").
'* @param FlagOffString String returned for False (default is an empty string).
'* @return Converted string.
'*
'* @details
'* Converts a Boolean value to the specified string.
'* Returns `FlagOnString` for True and `FlagOffString` for False.
Public Function ConvertBooleanToString( _
        ByVal BooleanValue As Boolean, _
        Optional ByVal FlagOnString As String = "Å°", _
        Optional ByVal FlagOffString As String = "") As String

    If BooleanValue Then
        ConvertBooleanToString = FlagOnString
    Else
        ConvertBooleanToString = FlagOffString
    End If
End Function

'* Converts a string to a Boolean value.
'*
'* @param FlagValue String to convert.
'* @param FlagOnString String treated as True (default is "Å°").
'* @param FlagOffString String treated as False (default is an empty string).
'* @return Converted Boolean value.
'*
'* @details
'* Determines a Boolean value based on the specified string.
'* Returns True for `FlagOnString` and False for `FlagOffString`.
'* Raises an error for any other value.
Public Function ConvertStringToBoolean( _
        ByVal FlagValue As String, _
        Optional ByVal FlagOnString As String = "Å°", _
        Optional ByVal FlagOffString As String = "") As Boolean

    If FlagValue = FlagOnString Then
        ConvertStringToBoolean = True
    ElseIf FlagValue = FlagOffString Then
        ConvertStringToBoolean = False
    Else
        Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="Allowed FlagValue values are """ & FlagOnString & """ or """ & FlagOffString & """. (" & FlagValue & ")"
    End If
End Function


' #############################################################################
'
' Arrays
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Array dimension conversion
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a two-dimensional array to a one-dimensional array.
'*
'* @param OriginalArray Two-dimensional array.
'* @param ColumnDirection [Optional] Whether to read in the column direction (vertical direction, first dimension). Default is False, which reads in the row direction (horizontal direction, second dimension).
'* @return One-dimensional array.
'*
'* @details
'* Converts a two-dimensional array to a 0-based one-dimensional array.
Public Function ConvertArray2dTo1d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()

    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray2dTo1d", Description:="The argument is not an array. (" & TypeName(OriginalArray) & ")"
    End If

    Dim row_low_bnd As Long
    Dim row_up_bnd As Long
    Dim col_low_bnd As Long
    Dim col_up_bnd As Long
    Dim total_cnt As Long
    Dim res_arr() As Variant
    Dim arr_idx As Long
    Dim row_index As Long
    Dim col_index As Long

    row_low_bnd = LBound(OriginalArray, 1)
    row_up_bnd = UBound(OriginalArray, 1)
    col_low_bnd = LBound(OriginalArray, 2)
    col_up_bnd = UBound(OriginalArray, 2)

    total_cnt = (row_up_bnd - row_low_bnd + 1) * (col_up_bnd - col_low_bnd + 1)
    ReDim res_arr(0 To total_cnt - 1)

    arr_idx = 0
    If ColumnDirection Then
        For col_index = col_low_bnd To col_up_bnd
            For row_index = row_low_bnd To row_up_bnd
                Call pAssignVariant(res_arr(arr_idx), OriginalArray(row_index, col_index))
                arr_idx = arr_idx + 1
            Next row_index
        Next col_index
    Else
        For row_index = row_low_bnd To row_up_bnd
            For col_index = col_low_bnd To col_up_bnd
                Call pAssignVariant(res_arr(arr_idx), OriginalArray(row_index, col_index))
                arr_idx = arr_idx + 1
            Next col_index
        Next row_index
    End If

    ConvertArray2dTo1d = res_arr
End Function

'* Converts a one-dimensional array to a two-dimensional array.
'*
'* @param OriginalArray One-dimensional array.
'* @param RowLBound [Optional] Lower bound of the row index (first dimension). Default is 1.
'* @param ColumnLBound [Optional] Lower bound of the column index (second dimension). Default is 1.
'* @param RowCount [Optional] Number of row elements (first dimension). Default is unlimited.
'* @param ColCount [Optional] Number of column elements (second dimension). Default is 1.
'* @param ColumnDirection [Optional] Whether to fill in the column direction (vertical direction, first dimension). Default is False, which fills in the row direction (horizontal direction, second dimension).
'* @return Two-dimensional array.
'*
'* @details
'* Converts a one-dimensional array to a two-dimensional array.
'* 0-element arrays are not supported. Branch on empty arrays in the caller.
Public Function ConvertArray1dTo2d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal RowLBound As Long = 1, _
        Optional ByVal ColLBound As Long = 1, _
        Optional ByVal RowCount As Long = -1, _
        Optional ByVal ColCount As Long = 1, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()

    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="The argument is not an array. (" & TypeName(OriginalArray) & ")"
    End If

    If RowCount <= 0 And ColCount <= 0 Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="Both the row count and column count were specified as unlimited values (0 or less)."
    End If

    If 0 < RowCount And 0 < ColCount Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="Both the row count and column count were specified as fixed element counts."
    End If

    Dim orig_count As Long
    orig_count = UBound(OriginalArray) - LBound(OriginalArray) + 1

    Dim count_1 As Long, count_2 As Long
    If 0 < RowCount Then
        count_1 = RowCount
        count_2 = (orig_count + RowCount - 1) \ RowCount
    Else
        count_1 = (orig_count + ColCount - 1) \ ColCount
        count_2 = ColCount
    End If

    Dim max_idx_1 As Long, max_idx_2 As Long
    max_idx_1 = RowLBound + count_1 - 1
    max_idx_2 = ColLBound + count_2 - 1

    Dim result_arr() As Variant
    ReDim result_arr(RowLBound To max_idx_1, ColLBound To max_idx_2)

    Dim idx_1 As Long, idx_2 As Long
    idx_1 = RowLBound
    idx_2 = ColLBound

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(OriginalArray)
    Do While enum_obj.MoveNext()
        If IsObject(enum_obj.Current) Then
            Set result_arr(idx_1, idx_2) = enum_obj.Current
        Else
            result_arr(idx_1, idx_2) = enum_obj.Current
        End If

        If ColumnDirection Then
            idx_1 = idx_1 + 1
            If max_idx_1 < idx_1 Then
                idx_1 = RowLBound
                idx_2 = idx_2 + 1
            End If
        Else
            idx_2 = idx_2 + 1
            If max_idx_2 < idx_2 Then
                idx_1 = idx_1 + 1
                idx_2 = ColLBound
            End If
        End If
    Loop

    ConvertArray1dTo2d = result_arr
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Array bounds and state
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Creates a Variant array with no elements.
'*
'* @return 0-based Variant array with 0 elements.
'*
'* @details
'* Returns an empty array with `LBound = 0` / `UBound = -1`, not an uninitialized array.
Public Function EmptyVariantArray() As Variant()
    EmptyVariantArray = Array()
End Function

'* Creates a String array with no elements.
'*
'* @return 0-based String array with 0 elements.
'*
'* @details
'* Returns an empty array with `LBound = 0` / `UBound = -1`, not an uninitialized array.
Public Function EmptyStringArray() As String()
    EmptyStringArray = Filter(Split(vbNullString, vbNullChar), vbNullChar)
End Function

'* Gets the maximum and minimum subscripts of an array.
'*
'* @param LBoundArray [Output] Minimum subscript for each dimension.
'* @param UBoundArray [Output] Maximum subscript for each dimension.
'* @param TargetArray Array to inspect.
'*
'* @details
'* Gets the maximum and minimum subscripts of an array. Raises an error for uninitialized arrays.
Public Sub GetArrayBounds(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal TargetArray As Variant)
    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, "Sub GetArrayBounds", "The value is not an array. (" & TypeName(TargetArray) & ")"
    End If

    Dim result_l() As Long
    Dim result_u() As Long

    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(result_l, result_u, TargetArray)
    If dim_count = 0 Then
        Err.Raise vbObjectError + 1, "Sub GetArrayBounds", "The array is not initialized."
    End If

    LBoundArray = result_l
    UBoundArray = result_u
End Sub

'* Determines whether an argument is an empty array.
'*
'* @param TargetArray Array to test.
'* @return True if the array is empty; otherwise, False.
'*
'* @details
'* Returns True when the specified argument is an array and has no elements.
'* Returns False when the argument is not an array.
'* Treats the array as empty when getting array bounds (`UBound` / `LBound`) fails.
Public Function IsEmptyArray(ByVal TargetArray As Variant) As Boolean
    If Not IsArray(TargetArray) Then
        IsEmptyArray = False
        Exit Function
    End If

    Dim result_value As Boolean
    result_value = False

    Err.Clear
    On Error Resume Next
    result_value = UBound(TargetArray) < LBound(TargetArray)
    If Err.Number <> 0 Then
        result_value = True
        Err.Clear
    End If
    On Error GoTo 0

    IsEmptyArray = result_value
End Function

'* Gets an array enumerator.
'*
'* @param TargetArray Array to enumerate.
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @param IsReadOnly [Optional] Whether to make the enumerator read-only.
'* @return Array enumerator (`IEnumerator` object).
'*
'* @details
'* Returns an `IEnumerator` object for enumerating the specified one-dimensional array.
'* Empty one-dimensional arrays are treated as valid input, and MoveNext returns False.
'* Uninitialized arrays and multidimensional arrays raise an error.
'* Update / Remove cannot be used with array enumerators.
'* If an array element is an object, it is set as a reference; otherwise, it is set as a value.
'*
'* Usage example:
'* @code
'* Dim enum_obj As IEnumerator
'* Set enum_obj = GetArrayEnumerator(some_arr)
'* Do While enum_obj.MoveNext()
'*     Dim item_var as Variable
'*     item_var = enum_obj.Current
'*     'Debug.Print item_var
'* Loop
'* @endcode
Public Function GetArrayEnumerator(ByVal TargetArray As Variant, Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(TargetArray, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetArrayEnumerator = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Array search
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Determines whether an array contains the specified value.
'*
'* @param CheckItem Value to test.
'* @param TargetArray Array to test.
'* @return True if the array contains the value; otherwise, False.
'*
'* @details
'* Determines whether an array contains the specified value.
Public Function ContainsInArray(ByVal CheckItem As Variant, ByRef TargetArray As Variant) As Boolean
    Dim arr_item As Variant

    For Each arr_item In TargetArray
        If pArrayItemsEqual(arr_item, CheckItem) Then
            ContainsInArray = True
            Exit Function
        End If
    Next arr_item

    ContainsInArray = False
End Function

Private Function pArrayItemsEqual(ByVal Item1 As Variant, ByVal Item2 As Variant) As Boolean
    If IsObject(Item1) Or IsObject(Item2) Then
        If IsObject(Item1) And IsObject(Item2) Then
            Dim obj_1 As Object
            Dim obj_2 As Object
            Set obj_1 = Item1
            Set obj_2 = Item2
            pArrayItemsEqual = (obj_1 Is obj_2)
        Else
            pArrayItemsEqual = False
        End If
        Exit Function
    End If

    If IsArray(Item1) Or IsArray(Item2) Then
        pArrayItemsEqual = False
        Exit Function
    End If

    If IsNull(Item1) Or IsNull(Item2) Then
        pArrayItemsEqual = (IsNull(Item1) And IsNull(Item2))
        Exit Function
    End If

    If IsError(Item1) Or IsError(Item2) Then
        If IsError(Item1) And IsError(Item2) Then
            pArrayItemsEqual = (CLng(Item1) = CLng(Item2))
        Else
            pArrayItemsEqual = False
        End If
        Exit Function
    End If

    If IsEmpty(Item1) Or IsEmpty(Item2) Then
        pArrayItemsEqual = (IsEmpty(Item1) And IsEmpty(Item2))
        Exit Function
    End If

    pArrayItemsEqual = (Item1 = Item2)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Array sorting
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Sorts an array.
'*
'* @param TargetArray Array to sort.
'* @param Descending True to sort in descending order. Default is False (ascending order).
'*
'* @details
'* Sorts an array.
'*
'* @note
'* Assumes each array element can be compared with inequality operators.
Public Sub SortArray(ByRef TargetArray As Variant, Optional ByVal Descending As Boolean = False)
    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, Source:="Sub SortArray", Description:="Argument is not an array. (" & TypeName(TargetArray) & ")"
    End If

    If IsEmptyArray(TargetArray) Then Exit Sub
    Call pValidateSortArrayItems(TargetArray)
    Call pSortArrayCore(TargetArray, LBound(TargetArray), UBound(TargetArray), Descending)
End Sub

Private Sub pValidateSortArrayItems(ByRef TargetArray As Variant)
    Dim arr_item As Variant
    For Each arr_item In TargetArray
        If IsObject(arr_item) Then
            Err.Raise vbObjectError + 1, Source:="Sub SortArray", Description:="Object items are not supported."
        End If
    Next arr_item
End Sub

Private Sub pSortArrayCore(ByRef TargetArray As Variant, ByVal MinIndex As Long, ByVal MaxIndex As Long, ByVal Descending As Boolean)
    If MinIndex >= MaxIndex Then Exit Sub
    Dim item_idx As Long
    Dim pos_lng As Long: pos_lng = MinIndex

    Call pSortSwap(TargetArray(MinIndex), TargetArray(Int((MinIndex + MaxIndex) \ 2)))
    For item_idx = MinIndex + 1 To MaxIndex
        If pSortIsLessThan(TargetArray(item_idx), TargetArray(MinIndex), Descending) Then
            pos_lng = pos_lng + 1
            Call pSortSwap(TargetArray(item_idx), TargetArray(pos_lng))
        End If
    Next
    Call pSortSwap(TargetArray(MinIndex), TargetArray(pos_lng))

    Call pSortArrayCore(TargetArray, MinIndex, pos_lng - 1, Descending)
    Call pSortArrayCore(TargetArray, pos_lng + 1, MaxIndex, Descending)
End Sub

Private Sub pSortSwap(ByRef Item1 As Variant, ByRef Item2 As Variant)
    Dim tmp_var As Variant
    Call pAssignVariant(tmp_var, Item1)
    Call pAssignVariant(Item1, Item2)
    Call pAssignVariant(Item2, tmp_var)
End Sub

Private Function pSortIsLessThan(ByVal Item1 As Variant, ByVal Item2 As Variant, ByVal Descending As Boolean) As Boolean
    If Descending Then
        pSortIsLessThan = (Item2 < Item1)
    Else
        pSortIsLessThan = (Item1 < Item2)
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Array joining
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Concatenates arrays.
'*
'* @param Array1 First array.
'* @param Array2 Second array.
'* @param OtherArrays Additional variable-length arrays.
'* @return Concatenated array.
'*
'* @details
'* Joins multiple arrays and returns a single array.
'* If an argument is not an array, it is treated as an array of length 1.
Public Function ConcatArray(ByVal Array1 As Variant, ByVal Array2 As Variant, ParamArray OtherArrays() As Variant) As Variant()
    Dim result_value() As Variant
    Dim arr_length As Long

    arr_length = pGetArrayLengthCore(Array1)
    arr_length = arr_length + pGetArrayLengthCore(Array2)

    Dim param_idx As Long
    Dim param_item As Variant
    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            Call pAssignVariant(param_item, OtherArrays(param_idx))
            arr_length = arr_length + pGetArrayLengthCore(param_item)
        Next param_idx
    End If

    If arr_length = 0 Then
        ConcatArray = Array()
        Exit Function
    End If

    'If IsArray(Array1) And Not IsEmptyArray(Array1) Then
    '    ReDim result_value(LBound(Array1) To LBound(Array1) + arr_length - 1)
    'Else
        ReDim result_value(0 To arr_length - 1)
    'End If

    Dim result_idx As Long
    result_idx = LBound(result_value)

    Call pConcatArrayCore(result_value, result_idx, Array1)
    Call pConcatArrayCore(result_value, result_idx, Array2)

    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            Call pAssignVariant(param_item, OtherArrays(param_idx))
            Call pConcatArrayCore(result_value, result_idx, param_item)
        Next param_idx
    End If

    ConcatArray = result_value
End Function

Private Function pGetArrayLengthCore(ByVal TargetArray As Variant) As Long
    Dim result_value As Long

    If IsEmptyArray(TargetArray) Then
        result_value = 0
    ElseIf IsArray(TargetArray) Then
        result_value = UBound(TargetArray) - LBound(TargetArray) + 1
    Else
        result_value = 1
    End If

    pGetArrayLengthCore = result_value
End Function

Private Sub pConcatArrayCore(ByRef ResultArray() As Variant, ByRef CurrentIndex As Long, ByVal TargetArray As Variant)
    If IsEmptyArray(TargetArray) Then
        Exit Sub
    End If

    Dim item_idx As Long
    If IsArray(TargetArray) Then
        For item_idx = LBound(TargetArray) To UBound(TargetArray)
            Call pAssignVariant(ResultArray(CurrentIndex), TargetArray(item_idx))
            CurrentIndex = CurrentIndex + 1
        Next item_idx
    Else
        Call pAssignVariant(ResultArray(CurrentIndex), TargetArray)
        CurrentIndex = CurrentIndex + 1
    End If
End Sub

Private Sub pAssignVariant(ByRef TargetValue As Variant, ByVal SourceValue As Variant)
    If IsObject(SourceValue) Then
        Set TargetValue = SourceValue
    Else
        TargetValue = SourceValue
    End If
End Sub


' #############################################################################
'
' Strings
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' ID notation
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Returns a display name that joins an ID number and name.
'*
'* @param IDNumber ID number.
'* @param Name Name.
'* @param NumFormat Format applied to the ID number. When omitted, two-digit zero padding is used.
'* @param Separator Separator between the ID number and name.
'* @return Formatted ID display name.
'*
'* @details
'* Converts IDNumber to a string with the Format function and joins it with Separator and Name.
Public Function FormatIDName( _
        ByVal IDNumber As Integer, _
        ByVal Name As String, _
        Optional ByVal NumFormat As String = "00", _
        Optional ByVal Separator As String = ".") As Variant

    FormatIDName = Format(IDNumber, NumFormat) & Separator & Name
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Type strings
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Returns whether a string can be used as an element type contract name.
'*
'* @param ElementTypeKey Element type contract name to check.
'* @return True if it can be used as an element type contract name.
'*
'* @details
'* To make it usable as a class module name, the first character is limited to a letter, the second and later characters to letters, digits, and underscores, and the whole name to 31 characters or fewer.
Public Function IsValidElementTypeKey(ByVal ElementTypeKey As String) As Boolean
    Dim result_value As Boolean
    result_value = False

    Dim key_length As Long
    key_length = Len(ElementTypeKey)
    If key_length < 1 Or 31 < key_length Then
        IsValidElementTypeKey = result_value
        Exit Function
    End If

    If Not pIsElementTypeKeyFirstChar(Mid$(ElementTypeKey, 1, 1)) Then
        IsValidElementTypeKey = result_value
        Exit Function
    End If

    Dim char_idx As Long
    For char_idx = 2 To key_length
        If Not pIsElementTypeKeyFollowingChar(Mid$(ElementTypeKey, char_idx, 1)) Then
            IsValidElementTypeKey = result_value
            Exit Function
        End If
    Next char_idx

    result_value = True
    IsValidElementTypeKey = result_value
End Function

'* Returns the type string of a value.
'*
'* @param Value Value whose type string is retrieved.
'* @param ObjectKeyMode Object keying method.
'* @param IncludeArrayBounds Whether to include array bound information. When True, takes precedence over IncludeArrayRank.
'* @param IncludeArrayRank Whether to include the number of array dimensions.
'* @return Type string of the value.
'*
'* @details
'* Scalar values return the same type name as VBA TypeName.
'* Objects are returned in Object@ClassName, IEquatable@ClassName, or IDuplicateCheckable@ClassName format according to ObjectKeyMode.
'* Arrays have [], [,], or [1:2,1:3] appended to the end.
Public Function GetTypeString( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal IncludeArrayBounds As Boolean = False, _
        Optional ByVal IncludeArrayRank As Boolean = False, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetTypeStringCore(Value, ObjectKeyMode, IncludeArrayBounds, IncludeArrayRank, False, UseElementTypeKey, "Function GetTypeString")

    GetTypeString = result_value
End Function

Private Function pGetTypeStringCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal IncludeArrayBounds As Boolean, _
        ByVal IncludeArrayRank As Boolean, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetArrayTypeName(Value, ObjectKeyMode, IncludeArrayBounds, IncludeArrayRank, UsePrimitiveTypeName, ErrorSource)
    ElseIf IsObject(Value) Then
        If Value Is Nothing Then
            result_value = "Nothing"
        Else
            result_value = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource) & "@" & pGetElementTypeName(Value, UseElementTypeKey, ErrorSource)
        End If
    Else
        result_value = TypeName(Value)
        If UsePrimitiveTypeName And pIsPrimitiveValueKeyTypeName(result_value) Then result_value = "Primitive"
    End If

    pGetTypeStringCore = result_value
End Function

Private Function pGetArrayTypeName( _
        ByVal TargetArray As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal IncludeArrayBounds As Boolean, _
        ByVal IncludeArrayRank As Boolean, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(lbound_arr, ubound_arr, TargetArray)

    Dim item_type_name As String
    item_type_name = pGetArrayItemTypeName(TypeName(TargetArray))

    Dim result_value As String
    If UsePrimitiveTypeName And pIsPrimitiveValueKeyTypeName(item_type_name) Then
        result_value = "Primitive"
    ElseIf pIsObjectArrayType(item_type_name) Then
        Dim object_array_type_name As String
        object_array_type_name = item_type_name
        If object_array_type_name = "Object" Then object_array_type_name = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource)
        result_value = pGetObjectKeyModeName(ObjectKeyMode, ErrorSource) & "@" & object_array_type_name
    Else
        result_value = item_type_name
    End If

    If IncludeArrayBounds And 0 < dim_count Then
        result_value = result_value & "[" & pGetArrayBoundsString(lbound_arr, ubound_arr, dim_count) & "]"
    ElseIf IncludeArrayRank And 0 < dim_count Then
        result_value = result_value & "[" & pGetArrayRankString(dim_count) & "]"
    Else
        result_value = result_value & "[]"
    End If

    pGetArrayTypeName = result_value
End Function

Private Function pGetArrayRankString(ByVal DimCount As Long) As String
    Dim result_value As String

    Dim dim_idx As Long
    For dim_idx = 2 To DimCount
        result_value = result_value & ","
    Next dim_idx

    pGetArrayRankString = result_value
End Function

Private Function pGetArrayBoundsCore( _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long, _
        ByVal TargetArray As Variant) As Long

    Dim dim_count As Long
    dim_count = 0

    Do
        Dim lower_bound As Long
        Dim upper_bound As Long

        Err.Clear
        On Error Resume Next
            lower_bound = LBound(TargetArray, dim_count + 1)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If

            upper_bound = UBound(TargetArray, dim_count + 1)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
        On Error GoTo 0

        ReDim Preserve LBoundArray(0 To dim_count)
        ReDim Preserve UBoundArray(0 To dim_count)
        LBoundArray(dim_count) = lower_bound
        UBoundArray(dim_count) = upper_bound

        dim_count = dim_count + 1
    Loop While True

    pGetArrayBoundsCore = dim_count
End Function

Private Function pGetArrayItemTypeName(ByVal ArrayTypeName As String) As String
    Dim result_value As String
    If Right$(ArrayTypeName, 2) = "()" Then
        result_value = Left$(ArrayTypeName, Len(ArrayTypeName) - 2)
    Else
        result_value = ArrayTypeName
    End If

    pGetArrayItemTypeName = result_value
End Function

Private Function pIsObjectArrayType(ByVal ItemTypeName As String) As Boolean
    Dim result_value As Boolean
    Select Case ItemTypeName
        Case "Boolean", "Byte", "Currency", "Date", "Decimal", "Double", "Integer", "Long", "LongLong", "LongPtr", "Short", "Single", "String", "Variant"
            result_value = False
        Case Else
            result_value = True
    End Select

    pIsObjectArrayType = result_value
End Function

Private Function pGetArrayBoundsString(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As String
    Dim result_value As String

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        If result_value <> "" Then result_value = result_value & ","
        result_value = result_value & CStr(LBoundArray(dim_idx)) & ":" & CStr(UBoundArray(dim_idx))
    Next dim_idx

    pGetArrayBoundsString = result_value
End Function

Private Function pGetElementTypeName( _
        ByVal Value As Object, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If UseElementTypeKey And TypeOf Value Is IElementTypeProvider Then
        Dim element_type_provider As IElementTypeProvider
        Set element_type_provider = Value

        result_value = element_type_provider.ElementTypeKey
        If Not IsValidElementTypeKey(result_value) Then
            Err.Raise vbObjectError + 1, ErrorSource, "ElementTypeKey must contain only characters allowed in class module names. (" & result_value & ")"
        End If
    Else
        result_value = TypeName(Value)
    End If

    pGetElementTypeName = result_value
End Function

Private Function pIsElementTypeKeyFirstChar(ByVal TargetChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TargetChar)

    Dim result_value As Boolean
    result_value = ((AscW("A") <= char_code And char_code <= AscW("Z")) _
            Or (AscW("a") <= char_code And char_code <= AscW("z")))

    pIsElementTypeKeyFirstChar = result_value
End Function

Private Function pIsElementTypeKeyFollowingChar(ByVal TargetChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TargetChar)

    Dim result_value As Boolean
    result_value = (pIsElementTypeKeyFirstChar(TargetChar) _
            Or (AscW("0") <= char_code And char_code <= AscW("9")) _
            Or TargetChar = "_")

    pIsElementTypeKeyFollowingChar = result_value
End Function

Private Function pGetObjectKeyModeName(ByVal ObjectKeyMode As Long, ByVal ErrorSource As String) As String
    Dim result_value As String
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            result_value = "Object"
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            result_value = "IEquatable"
        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            result_value = "IDuplicateCheckable"
        Case Else
            Err.Raise vbObjectError + 1, ErrorSource, "Unsupported object key mode. (" & CStr(ObjectKeyMode) & ")"
    End Select

    pGetObjectKeyModeName = result_value
End Function

Private Function pIsPrimitiveValueKeyTypeName(ByVal TypeNameString As String) As Boolean
    Dim result_value As Boolean
    Select Case TypeNameString
        Case "Boolean", "Byte", "Date", "Decimal", "Double", "Integer", "Long", "LongLong", "LongPtr", "Short", "Single", "String"
            result_value = True
        Case Else
            result_value = False
    End Select

    pIsPrimitiveValueKeyTypeName = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Value keys
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a value to a key string with type information.
'*
'* @param Value Value to key.
'* @param ObjectKeyMode Object keying method.
'* @return Key string with type information.
'*
'* @details
'* Returns an internal representation for use as a key in Dictionary and similar structures.
'* Converts to a string including scalar values, special values, array dimensions / bounds / elements, and object identification method.
'* Scalar values are also distinguished by type name.
Public Function GetTypedValueKey( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetValueKeyCore(Value, ObjectKeyMode, False, UseElementTypeKey, "Function GetTypedValueKey")

    GetTypedValueKey = result_value
End Function

'* Converts a value to a key string.
'*
'* @param Value Value to key.
'* @param ObjectKeyMode Object keying method.
'* @return Key string.
'*
'* @details
'* Keys using the same rules as `GetTypedValueKey`, but primitive values except Error / Currency / Variant / Null / Empty / object types are treated as the same Primitive.
Public Function GetValueKey( _
        ByVal Value As Variant, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    Dim result_value As String
    result_value = pGetValueKeyCore(Value, ObjectKeyMode, True, UseElementTypeKey, "Function GetValueKey")

    GetValueKey = result_value
End Function

Private Function pGetValueKeyCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetArrayValueKeyString(Value, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource)
    ElseIf IsObject(Value) Then
        result_value = pGetObjectValueKeyString(Value, ObjectKeyMode, UseElementTypeKey, ErrorSource)
    Else
        result_value = pGetTypeStringCore(Value, ObjectKeyMode, False, False, UsePrimitiveTypeName, False, ErrorSource) _
                & "(" & pGetScalarValueKeyString(Value) & ")"
    End If

    pGetValueKeyCore = result_value
End Function

Private Function pGetArrayValueKeyString( _
        ByVal TargetArray As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Dim dim_count As Long
    dim_count = pGetArrayBoundsCore(lbound_arr, ubound_arr, TargetArray)

    Dim array_prefix As String
    array_prefix = pGetTypeStringCore(TargetArray, ObjectKeyMode, True, False, UsePrimitiveTypeName, False, ErrorSource)

    Dim result_value As String
    If dim_count = 0 Then
        result_value = array_prefix & "()"
    Else
        Dim total_count As Long
        total_count = pGetArrayItemCount(lbound_arr, ubound_arr, dim_count)

        If total_count = 0 Then
            result_value = array_prefix & "()"
        Else
            Dim item_type_name As String
            item_type_name = pGetArrayItemTypeName(TypeName(TargetArray))

            Dim flat_items() As Variant
            Call pFlattenArray(flat_items, TargetArray, total_count)

            Dim index_arr() As Long
            ReDim index_arr(0 To dim_count - 1)

            result_value = array_prefix & "(" _
                    & pGetArrayItemsString(flat_items, lbound_arr, ubound_arr, index_arr, 0, item_type_name, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource) & ")"
        End If
    End If

    pGetArrayValueKeyString = result_value
End Function

Private Function pIsEmptyArrayBounds(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As Boolean
    Dim result_value As Boolean
    result_value = False

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        If UBoundArray(dim_idx) < LBoundArray(dim_idx) Then
            result_value = True
            Exit For
        End If
    Next dim_idx

    pIsEmptyArrayBounds = result_value
End Function

Private Function pGetArrayItemCount(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal DimCount As Long) As Long
    Dim result_value As Long
    result_value = 1

    Dim dim_idx As Long
    For dim_idx = 0 To DimCount - 1
        result_value = result_value * (UBoundArray(dim_idx) - LBoundArray(dim_idx) + 1)
    Next dim_idx

    pGetArrayItemCount = result_value
End Function

Private Sub pFlattenArray(ByRef FlatItems() As Variant, ByVal TargetArray As Variant, ByVal ItemCount As Long)
    ReDim FlatItems(0 To ItemCount - 1)

    Dim item_idx As Long
    item_idx = 0

    Dim item_value As Variant
    For Each item_value In TargetArray
        If IsObject(item_value) Then
            Set FlatItems(item_idx) = item_value
        Else
            FlatItems(item_idx) = item_value
        End If
        item_idx = item_idx + 1
    Next item_value
End Sub

Private Function pGetArrayItemsString( _
        ByRef FlatItems() As Variant, _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long, _
        ByRef IndexArray() As Long, _
        ByVal CurrentDim As Long, _
        ByVal ItemTypeName As String, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    Dim item_idx As Long

    If CurrentDim = UBound(IndexArray) Then
        For item_idx = LBoundArray(CurrentDim) To UBoundArray(CurrentDim)
            If result_value <> "" Then result_value = result_value & ","
            IndexArray(CurrentDim) = item_idx
            result_value = result_value & pGetArrayItemString(FlatItems(pGetForEachFlatIndex(IndexArray, LBoundArray, UBoundArray)), ItemTypeName, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource)
        Next item_idx
    Else
        For item_idx = LBoundArray(CurrentDim) To UBoundArray(CurrentDim)
            If result_value <> "" Then result_value = result_value & ","
            IndexArray(CurrentDim) = item_idx
            result_value = result_value & "(" & pGetArrayItemsString(FlatItems, LBoundArray, UBoundArray, IndexArray, CurrentDim + 1, ItemTypeName, ObjectKeyMode, UsePrimitiveTypeName, ErrorSource) & ")"
        Next item_idx
    End If

    pGetArrayItemsString = result_value
End Function

Private Function pGetForEachFlatIndex( _
        ByRef IndexArray() As Long, _
        ByRef LBoundArray() As Long, _
        ByRef UBoundArray() As Long) As Long

    Dim result_value As Long
    Dim multiplier As Long
    result_value = 0
    multiplier = 1

    Dim dim_idx As Long
    For dim_idx = LBound(IndexArray) To UBound(IndexArray)
        result_value = result_value + (IndexArray(dim_idx) - LBoundArray(dim_idx)) * multiplier
        multiplier = multiplier * (UBoundArray(dim_idx) - LBoundArray(dim_idx) + 1)
    Next dim_idx

    pGetForEachFlatIndex = result_value
End Function

Private Function pGetArrayItemString( _
        ByVal ItemValue As Variant, _
        ByVal ItemTypeName As String, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If ItemTypeName = "Variant" Or pIsObjectArrayType(ItemTypeName) Then
        result_value = pGetValueKeyCore(ItemValue, ObjectKeyMode, UsePrimitiveTypeName, False, ErrorSource)
    Else
        result_value = pGetScalarValueKeyString(ItemValue)
    End If

    pGetArrayItemString = result_value
End Function

Private Function pGetObjectValueKeyString( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If Value Is Nothing Then
        result_value = "Nothing()"
    Else
        Dim object_type_name As String
        object_type_name = pGetTypeStringCore(Value, ObjectKeyMode, False, False, False, UseElementTypeKey, ErrorSource)

        Select Case ObjectKeyMode
            Case G_OBJECT_KEY_MODE_REFERENCE
                result_value = object_type_name & "(" & Hex$(ObjPtr(Value)) & ")"

            Case G_OBJECT_KEY_MODE_I_EQUATABLE
                If Not TypeOf Value Is IEquatable Then
                    Err.Raise vbObjectError + 1, ErrorSource, "Attempted to key an object that does not implement IEquatable in IEquatable mode. (" & object_type_name & ")"
                End If

                Dim eq_item As IEquatable
                Set eq_item = Value
                result_value = object_type_name & "(" & pEscapeValueKeyString(eq_item.GetIdentityString()) & ")"

            Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
                If Not TypeOf Value Is IDuplicateCheckable Then
                    Err.Raise vbObjectError + 1, ErrorSource, "Attempted to key an object that does not implement IDuplicateCheckable in IDuplicateCheckable mode. (" & object_type_name & ")"
                End If

                Dim dup_item As IDuplicateCheckable
                Set dup_item = Value
                result_value = object_type_name & "(" & pEscapeValueKeyString(dup_item.GetKey()) & ")"

            Case Else
                Err.Raise vbObjectError + 1, ErrorSource, "Unsupported object key mode. (" & CStr(ObjectKeyMode) & ")"
        End Select
    End If

    pGetObjectValueKeyString = result_value
End Function

Private Function pGetScalarValueKeyString(ByVal Value As Variant) As String
    Dim result_value As String
    If IsError(Value) Then
        result_value = CStr(CLng(Value))
    ElseIf IsNull(Value) Or IsEmpty(Value) Then
        result_value = ""
    ElseIf VarType(Value) = vbDate Then
        result_value = pEscapeValueKeyString(Format$(Value, "yyyy-mm-dd\Thh:nn:ss"))
    ElseIf VarType(Value) = vbString Then
        result_value = pEscapeValueKeyString(CStr(Value))
    Else
        result_value = pEscapeValueKeyString(CStr(Value))
    End If

    pGetScalarValueKeyString = result_value
End Function

Private Function pEscapeValueKeyString(ByVal Expression As String) As String
    Dim result_value As String
    result_value = Replace(Expression, "\", "\\")
    result_value = Replace(result_value, vbTab, "\t")
    result_value = Replace(result_value, "(", "\(")
    result_value = Replace(result_value, ")", "\)")
    result_value = Replace(result_value, "[", "\[")
    result_value = Replace(result_value, "]", "\]")
    result_value = Replace(result_value, ",", "\,")
    result_value = Replace(result_value, ":", "\:")
    result_value = Replace(result_value, "@", "\@")
    result_value = Replace(result_value, "=", "\=")

    pEscapeValueKeyString = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Composite keys
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Concatenates multiple keys and returns them as a string with type information.
'*
'* @param DictionaryKeys Keys (variable-length arguments). Returns an empty string when 0 keys are specified.
'* @return Concatenated key string.
'*
'* @details
'* Keys the specified multiple keys using the same basic rules as `GetTypedValueKey` and returns them as one string.
'* The separator used for keys is the tab character (`vbTab`).
'* Tab characters inside each key are escaped by `GetTypedValueKey`.
Public Function GetTypedMultiKey(ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetTypedMultiKey = GetTypedMultiKeyByModeFromArray(G_OBJECT_KEY_MODE_REFERENCE, key_arr)
End Function

'* Concatenates multiple keys as a string with type information, using the specified object keying mode.
'*
'* @param ObjectKeyMode Top-level object keying method.
'* @param DictionaryKeys Keys (variable-length arguments). Returns an empty string when 0 keys are specified.
'* @return Concatenated key string.
'*
'* @details
'* Top-level object values are keyed with ObjectKeyMode taking precedence, and objects that do not support the specified mode fall back to reference keys.
'* Array arguments are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
Public Function GetTypedMultiKeyByMode(ByVal ObjectKeyMode As Long, ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetTypedMultiKeyByMode = GetTypedMultiKeyByModeFromArray(ObjectKeyMode, key_arr)
End Function

'* Concatenates multiple keys from an array as a string with type information, using the specified object keying mode.
'*
'* @param ObjectKeyMode Top-level object keying method.
'* @param DictionaryKeys One-dimensional array in the same format as ParamArray.
'* @return Concatenated key string.
'*
'* @details
'* This function exists as a workaround because VBA cannot expand an array into a ParamArray.
'* Use `GetTypedMultiKeyByMode` for normal calls.
'* Top-level object values are keyed with ObjectKeyMode taking precedence, and objects that do not support the specified mode fall back to reference keys.
'* Array arguments are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
Public Function GetTypedMultiKeyByModeFromArray( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    GetTypedMultiKeyByModeFromArray = pGetMultiKeyByModeFromArrayCore( _
            ObjectKeyMode, _
            DictionaryKeys, _
            False, _
            UseElementTypeKey, _
            "Function GetTypedMultiKeyByModeFromArray")
End Function

'* Concatenates multiple keys and returns the result.
'*
'* @param DictionaryKeys Keys (variable-length arguments). Returns an empty string when 0 keys are specified.
'* @return Concatenated key string.
'*
'* @details
'* Keys the specified multiple keys using the same basic rules as `GetValueKey` and returns them as one string.
'* Primitive values except Error / Currency / Variant / Null / Empty / object types are treated as the same Primitive.
Public Function GetMultiKey(ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetMultiKey = GetMultiKeyByModeFromArray(G_OBJECT_KEY_MODE_REFERENCE, key_arr)
End Function

'* Concatenates multiple keys using the specified object keying mode.
'*
'* @param ObjectKeyMode Top-level object keying method.
'* @param DictionaryKeys Keys (variable-length arguments). Returns an empty string when 0 keys are specified.
'* @return Concatenated key string.
'*
'* @details
'* Top-level object values are keyed with ObjectKeyMode taking precedence, and objects that do not support the specified mode fall back to reference keys.
'* Array arguments are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
'* Primitive values except Error / Currency / Variant / Null / Empty / object types are treated as the same Primitive.
Public Function GetMultiKeyByMode(ByVal ObjectKeyMode As Long, ParamArray DictionaryKeys() As Variant) As String
    Dim key_arr As Variant
    key_arr = DictionaryKeys

    GetMultiKeyByMode = GetMultiKeyByModeFromArray(ObjectKeyMode, key_arr)
End Function

'* Concatenates multiple keys from an array using the specified object keying mode.
'*
'* @param ObjectKeyMode Top-level object keying method.
'* @param DictionaryKeys One-dimensional array in the same format as ParamArray.
'* @return Concatenated key string.
'*
'* @details
'* This function exists as a workaround because VBA cannot expand an array into a ParamArray.
'* Use `GetMultiKeyByMode` for normal calls.
'* Top-level object values are keyed with ObjectKeyMode taking precedence, and objects that do not support the specified mode fall back to reference keys.
'* Array arguments are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
'* Primitive values except Error / Currency / Variant / Null / Empty / object types are treated as the same Primitive.
Public Function GetMultiKeyByModeFromArray( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        Optional ByVal UseElementTypeKey As Boolean = False) As String

    GetMultiKeyByModeFromArray = pGetMultiKeyByModeFromArrayCore( _
            ObjectKeyMode, _
            DictionaryKeys, _
            True, _
            UseElementTypeKey, _
            "Function GetMultiKeyByModeFromArray")
End Function

Private Function pGetMultiKeyByModeFromArrayCore( _
        ByVal ObjectKeyMode As Long, _
        ByVal DictionaryKeys As Variant, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Call pGetObjectKeyModeName(ObjectKeyMode, ErrorSource)

    Dim lower_bound As Long
    Dim upper_bound As Long
    Call pGetArgumentArrayBounds(lower_bound, upper_bound, DictionaryKeys, ErrorSource)

    If upper_bound < lower_bound Then
        pGetMultiKeyByModeFromArrayCore = ""
        Exit Function
    End If

    Dim result_value As String
    Dim item_idx As Long
    For item_idx = lower_bound To upper_bound
        If result_value <> "" Then result_value = result_value & vbTab
        result_value = result_value & pGetFallbackValueKeyCore(DictionaryKeys(item_idx), ObjectKeyMode, UsePrimitiveTypeName, UseElementTypeKey, ErrorSource)
    Next item_idx

    pGetMultiKeyByModeFromArrayCore = result_value
End Function

Private Sub pGetArgumentArrayBounds( _
        ByRef LowerBoundValue As Long, _
        ByRef UpperBoundValue As Long, _
        ByVal TargetArray As Variant, _
        ByVal ErrorSource As String)

    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, ErrorSource, "Specify a one-dimensional array in the same format as ParamArray."
    End If

    On Error GoTo INVALID_ARRAY
    LowerBoundValue = LBound(TargetArray, 1)
    UpperBoundValue = UBound(TargetArray, 1)
    On Error GoTo 0

    Dim second_dim_lower_bound As Long
    On Error Resume Next
    second_dim_lower_bound = LBound(TargetArray, 2)
    If Err.Number = 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 1, ErrorSource, "Specify a one-dimensional array in the same format as ParamArray."
    End If
    Err.Clear
    On Error GoTo 0
    Exit Sub

INVALID_ARRAY:
    On Error GoTo 0
    Err.Raise vbObjectError + 1, ErrorSource, "Specify an initialized one-dimensional array."
End Sub

Private Function pGetFallbackValueKeyCore( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UsePrimitiveTypeName As Boolean, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim result_value As String
    If IsArray(Value) Then
        result_value = pGetValueKeyCore(Value, G_OBJECT_KEY_MODE_REFERENCE, UsePrimitiveTypeName, False, ErrorSource)
    ElseIf IsObject(Value) Then
        result_value = pGetFallbackObjectValueKeyString(Value, ObjectKeyMode, UseElementTypeKey, ErrorSource)
    Else
        result_value = pGetValueKeyCore(Value, G_OBJECT_KEY_MODE_REFERENCE, UsePrimitiveTypeName, False, ErrorSource)
    End If

    pGetFallbackValueKeyCore = result_value
End Function

Private Function pGetFallbackObjectValueKeyString( _
        ByVal Value As Variant, _
        ByVal ObjectKeyMode As Long, _
        ByVal UseElementTypeKey As Boolean, _
        ByVal ErrorSource As String) As String

    Dim effective_object_key_mode As Long
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE

        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IEquatable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_I_EQUATABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IDuplicateCheckable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case Else
            Err.Raise vbObjectError + 1, ErrorSource, "Unsupported object key mode. (" & CStr(ObjectKeyMode) & ")"
    End Select

    pGetFallbackObjectValueKeyString = pGetObjectValueKeyString(Value, effective_object_key_mode, UseElementTypeKey, ErrorSource)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' String arrays
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Diffs arrays of strings.
'*
'* @param OldArray [Input/Output] Old string array. Must be ReDim-capable.
'* @param NewArray [Input/Output] New string array. Must be ReDim-capable.
'* @param ChangeTypeArray [Output] Array of change types. Must be ReDim-capable. The passed array may be uninitialized. If the result has one or more items, the lower bound is always 0. If the result has 0 items, it becomes an empty array.
'* @param EnableReplaceType [Optional] Whether consecutive deletion and addition parts are treated as replacement. Default is True, treating them as replacement.
'* @param ReplaceCost [Optional] Replacement cost when deletion and addition each cost 1 (0 to 2). Default is 1.5.
'*
'* @details
'* Calculates the minimum edit distance for arrays OldArray and NewArray,
'* and overwrites each array with sequences of (Old row, New row, change type).
'* Result arrays are normalized to 0-based regardless of the lower bounds of the input arrays.
'* OldArray / NewArray can be uninitialized or empty arrays. If both are empty, the output arrays also become empty arrays.
'*
'* The change type is one of "" (match), "DEL" (delete), "ADD" (add), or "MOD" (replace).
'* For a matching type, both the OldArray and NewArray index contain the original string.
'* For a deletion, the OldArray index contains the original string and the NewArray index contains an empty string.
'* For an addition, the NewArray index contains the new string and the OldArray index contains an empty string.
'* For a replacement, the OldArray index contains the original string and the NewArray index contains the new string.
Public Sub DiffStringArray( _
        ByRef OldArray() As String, _
        ByRef NewArray() As String, _
        ByRef ChangeTypeArray() As String, _
        Optional ByVal EnableReplaceType As Boolean = True, _
        Optional ByVal ReplaceCost As Double = 1.5)

    Dim diff_old_coll As New Collection
    Dim diff_new_coll As New Collection
    Dim diff_op_coll  As New Collection

    If Not EnableReplaceType Then ReplaceCost = 3

    Dim old_start_idx As Long
    Dim old_end_idx As Long
    Dim new_start_idx As Long
    Dim new_end_idx As Long

    If IsEmptyArray(OldArray) Then
        old_start_idx = 0
        old_end_idx = -1
    Else
        old_start_idx = LBound(OldArray)
        old_end_idx = UBound(OldArray)
    End If

    If IsEmptyArray(NewArray) Then
        new_start_idx = 0
        new_end_idx = -1
    Else
        new_start_idx = LBound(NewArray)
        new_end_idx = UBound(NewArray)
    End If

    ' Recursively calculate the differences (Hirschberg algorithm).
    Call pHirschbergDiffRecursive(OldArray, old_start_idx, old_end_idx, _
                              NewArray, new_start_idx, new_end_idx, _
                              diff_old_coll, diff_new_coll, diff_op_coll, ReplaceCost)

    Dim num_items_val As Long, loop_idx_val As Long
    num_items_val = diff_old_coll.Count
    If num_items_val = 0 Then
        Erase OldArray
        Erase NewArray
        Erase ChangeTypeArray
        Exit Sub
    End If

    ReDim OldArray(0 To num_items_val - 1)
    ReDim NewArray(0 To num_items_val - 1)
    ReDim ChangeTypeArray(0 To num_items_val - 1)

    For loop_idx_val = 1 To num_items_val
        OldArray(loop_idx_val - 1) = diff_old_coll(loop_idx_val)
        NewArray(loop_idx_val - 1) = diff_new_coll(loop_idx_val)
        ChangeTypeArray(loop_idx_val - 1) = diff_op_coll(loop_idx_val)
    Next loop_idx_val
End Sub

Private Sub pHirschbergDiffRecursive(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)

    Dim SMALL_THRESHOLD As Long: SMALL_THRESHOLD = 1

    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1

    If old_count_val = 0 Then
        Dim loop_idx_new As Long
        For loop_idx_new = StartIdxNew To EndIdxNew
            DiffOldColl.Add ""
            DiffNewColl.Add NewArr(loop_idx_new)
            DiffOpColl.Add "ADD"
        Next loop_idx_new
        Exit Sub
    ElseIf new_count_val = 0 Then
        Dim loop_idx_old As Long
        For loop_idx_old = StartIdxOld To EndIdxOld
            DiffOldColl.Add OldArr(loop_idx_old)
            DiffNewColl.Add ""
            DiffOpColl.Add "DEL"
        Next loop_idx_old
        Exit Sub
    ElseIf old_count_val <= SMALL_THRESHOLD Or new_count_val <= SMALL_THRESHOLD Then
        Call pDiffSmall(OldArr, StartIdxOld, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
        Exit Sub
    End If

    Dim mid_idx_val As Long
    mid_idx_val = (StartIdxOld + EndIdxOld) \ 2

    Dim forward_costs_arr As Variant
    forward_costs_arr = pComputeForwardCost(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)

    Dim backward_costs_arr As Variant
    backward_costs_arr = pComputeBackwardCost(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)

    Dim split_j_val As Long, loop_idx_new_val As Long
    Dim min_cost_val As Double
    min_cost_val = 1E+99
    Dim new_size_val As Long
    new_size_val = new_count_val
    Dim cur_cost_val As Double
    For loop_idx_new_val = 0 To new_size_val
        cur_cost_val = forward_costs_arr(loop_idx_new_val) + backward_costs_arr(new_size_val - loop_idx_new_val)
        If cur_cost_val < min_cost_val Then
            min_cost_val = cur_cost_val
            split_j_val = loop_idx_new_val
        End If
    Next loop_idx_new_val

    Call pHirschbergDiffRecursive(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, StartIdxNew + split_j_val - 1, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
    Call pHirschbergDiffRecursive(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew + split_j_val, EndIdxNew, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
End Sub

Private Function pComputeForwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)

    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val

    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = StartIdxOld To EndIdxOld
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(StartIdxNew + inner_idx_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val

    pComputeForwardCost = cost_arr
End Function

Private Function pComputeBackwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)

    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val

    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = EndIdxOld To StartIdxOld Step -1
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(EndIdxNew - inner_idx_val + 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val

    pComputeBackwardCost = cost_arr
End Function

Private Sub pDiffSmall(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)

    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0

    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim dp_arr() As Double, op_arr() As String
    ReDim dp_arr(0 To old_count_val, 0 To new_count_val)
    ReDim op_arr(0 To old_count_val, 0 To new_count_val)

    Dim loop_old_val As Long, loop_new_val As Long
    dp_arr(0, 0) = 0
    op_arr(0, 0) = "DONE"
    For loop_old_val = 1 To old_count_val
        dp_arr(loop_old_val, 0) = dp_arr(loop_old_val - 1, 0) + COST_DEL
        op_arr(loop_old_val, 0) = "DEL"
    Next loop_old_val
    For loop_new_val = 1 To new_count_val
        dp_arr(0, loop_new_val) = dp_arr(0, loop_new_val - 1) + COST_ADD
        op_arr(0, loop_new_val) = "ADD"
    Next loop_new_val

    Dim inner_old_val As Long, inner_new_val As Long
    Dim sub_cost_val As Double, del_cost_val As Double, add_cost_val As Double, mod_cost_val As Double
    For inner_old_val = 1 To old_count_val
        For inner_new_val = 1 To new_count_val
            If OldArr(StartIdxOld + inner_old_val - 1) = NewArr(StartIdxNew + inner_new_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            del_cost_val = dp_arr(inner_old_val - 1, inner_new_val) + COST_DEL
            add_cost_val = dp_arr(inner_old_val, inner_new_val - 1) + COST_ADD
            mod_cost_val = dp_arr(inner_old_val - 1, inner_new_val - 1) + sub_cost_val
            dp_arr(inner_old_val, inner_new_val) = Application.Min(del_cost_val, add_cost_val, mod_cost_val)
            If dp_arr(inner_old_val, inner_new_val) = mod_cost_val Then
                op_arr(inner_old_val, inner_new_val) = IIf(sub_cost_val = COST_EQUAL, "", "MOD")
            ElseIf dp_arr(inner_old_val, inner_new_val) = del_cost_val Then
                op_arr(inner_old_val, inner_new_val) = "DEL"
            Else
                op_arr(inner_old_val, inner_new_val) = "ADD"
            End If
        Next inner_new_val
    Next inner_old_val

    Dim pos_old_val As Long, pos_new_val As Long
    pos_old_val = old_count_val: pos_new_val = new_count_val
    Dim temp_old_arr() As String, temp_new_arr() As String, temp_op_arr() As String
    ReDim temp_old_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_new_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_op_arr(0 To old_count_val + new_count_val - 1)
    Dim step_idx_val As Long: step_idx_val = 0
    Do While pos_old_val > 0 Or pos_new_val > 0
        Dim cur_op_val As String
        cur_op_val = op_arr(pos_old_val, pos_new_val)
        Select Case cur_op_val
            Case ""
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = ""
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "MOD"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "MOD"
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "DEL"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = ""
                temp_op_arr(step_idx_val) = "DEL"
                pos_old_val = pos_old_val - 1
            Case "ADD"
                temp_old_arr(step_idx_val) = ""
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "ADD"
                pos_new_val = pos_new_val - 1
        End Select
        step_idx_val = step_idx_val + 1
    Loop

    Dim num_steps_val As Long: num_steps_val = step_idx_val
    Dim rev_old_arr() As String, rev_new_arr() As String, rev_op_arr() As String
    ReDim rev_old_arr(0 To num_steps_val - 1)
    ReDim rev_new_arr(0 To num_steps_val - 1)
    ReDim rev_op_arr(0 To num_steps_val - 1)
    For loop_old_val = 0 To num_steps_val - 1
        rev_old_arr(loop_old_val) = temp_old_arr(num_steps_val - 1 - loop_old_val)
        rev_new_arr(loop_old_val) = temp_new_arr(num_steps_val - 1 - loop_old_val)
        rev_op_arr(loop_old_val) = temp_op_arr(num_steps_val - 1 - loop_old_val)
    Next loop_old_val

    Dim loop_idx_final_val As Long
    For loop_idx_final_val = 0 To num_steps_val - 1
        DiffOldColl.Add rev_old_arr(loop_idx_final_val)
        DiffNewColl.Add rev_new_arr(loop_idx_final_val)
        DiffOpColl.Add rev_op_arr(loop_idx_final_val)
    Next loop_idx_final_val
End Sub

'* Determines whether a String array is empty.
'*
'* @param StringArray String array to test.
'* @param BlankAsEmpty [Optional] Whether to treat an array with one element whose content is an empty string as empty. Default is True.
'* @return True if empty; otherwise, False.
'*
'* @details
'* Returns True when the array is uninitialized or has fewer than 1 element.
'* When BlankAsEmpty is True, also returns True when there is one element and that only element is an empty string.
Public Function IsEmptyStringArray( _
        ByRef StringArray() As String, _
        Optional ByVal BlankAsEmpty As Boolean = True) As Boolean
    If (Not StringArray) = -1 Then
        ' Uninitialized.
        IsEmptyStringArray = True
    ElseIf UBound(StringArray) < LBound(StringArray) Then
        ' Length is less than 1.
        IsEmptyStringArray = True
    ElseIf LBound(StringArray) < UBound(StringArray) Then
        ' Length is 2 or more.
        IsEmptyStringArray = False
    ElseIf BlankAsEmpty And StringArray(LBound(StringArray)) = "" Then
        ' Length is 1 and the content is an empty string.
        IsEmptyStringArray = True
    Else
        ' Otherwise.
        IsEmptyStringArray = False
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' String replacement
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Performs replacements on a string using multiple conditions.
'*
'* @param Expression Target string.
'* @param Find1 First string to replace.
'* @param Replace1 First replacement string (string or array).
'* @param FindReplacePairs Variable-length pairs of strings to replace and replacement strings.
'* @return Array of strings after replacement.
'*
'* @details
'* Replaces strings under the specified conditions and returns all results as an array.
'* If a replacement string is an array, each combination is included in the result.
'* Empty strings cannot be specified as strings to replace.
'* Raises an error when a replacement string is an empty array.
Public Function ReplaceMulti(ByVal Expression As String, ByVal Find1 As String, ByVal Replace1 As Variant, ParamArray FindReplacePairs() As Variant) As String()
    Dim result_list As ObjectList
    Set result_list = New ObjectList

    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        If (UBound(FindReplacePairs) - LBound(FindReplacePairs) + 1) Mod 2 <> 0 Then
            Err.Raise vbObjectError + 1, "Function ReplaceMulti", "An odd number of variable-length strings was supplied. Replacement targets and replacement strings do not match."
            Exit Function
        End If
    End If

    Call result_list.Add(Expression)
    Call pReplaceMultiCore(result_list, Find1, Replace1)

    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        Dim item_idx As Long
        For item_idx = LBound(FindReplacePairs) To UBound(FindReplacePairs) - 1 Step 2
            Call pReplaceMultiCore(result_list, FindReplacePairs(item_idx), FindReplacePairs(item_idx + 1))
        Next item_idx
    End If

    ReplaceMulti = result_list.ConvertToStringArray()
End Function

Private Sub pReplaceMultiCore(ByRef ResultList As ObjectList, ByVal FindString As String, ByVal Replaces As Variant)
    If FindString = "" Then
        Err.Raise vbObjectError + 1, "Function ReplaceMulti", "The replacement target string cannot be an empty string."
    End If

    If IsArray(Replaces) Then
        If IsEmptyArray(Replaces) Then
            Err.Raise vbObjectError + 1, "Function ReplaceMulti", "The replacement string candidate array is empty."
        End If
    End If

    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Dim enum_obj As IEnumerator
    Set enum_obj = ResultList.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim current_item As String
        current_item = enum_obj.Current
        If IsArray(Replaces) Then
            Dim replace_item As Variant
            For Each replace_item In Replaces
                Call result_value.Add(Replace(current_item, FindString, replace_item))
            Next replace_item
        Else
            Call result_value.Add(Replace(current_item, FindString, Replaces))
        End If
    Loop

    Set ResultList = result_value
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Newline characters
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Sub pValidateLineSeparatorEscSeqChar(ByVal EscSeqChar As String, ByVal ErrorSource As String)
    If Len(EscSeqChar) <> 1 Then
        Err.Raise vbObjectError + 1, ErrorSource, "EscSeqChar must be one character."
    End If
End Sub

'* Escapes newline characters.
'*
'* @param Expression Target string.
'* @param EscSeqChar Escape character (default is `\`).
'* @return String with newline characters escaped.
'*
'* @details
'* Replaces newline characters (`vbCr`, `vbLf`) with the specified escape character.
'* If EscSeqChar is an empty string, it is treated as omitted and `\` is used.
'* EscSeqChar can specify only one character, except for an empty string.
Public Function EscapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "") As String
    If EscSeqChar = "" Then EscSeqChar = "\"
    Call pValidateLineSeparatorEscSeqChar(EscSeqChar, "Function EscapeLineSeparator")

    Dim result_value As String
    result_value = Replace(Expression, EscSeqChar, EscSeqChar & EscSeqChar)
    result_value = Replace(result_value, vbCr, EscSeqChar & "r")
    result_value = Replace(result_value, vbLf, EscSeqChar & "n")
    EscapeLineSeparator = result_value
End Function

'* Unescapes newline characters.
'*
'* @param Expression Target string.
'* @param EscSeqChar Escape character (default is `\`).
'* @return Unescaped string.
'*
'* @details
'* Restores escaped newline sequences (`\n`, `\r`) to the original newline characters (`vbCr`, `vbLf`).
'* If EscSeqChar is an empty string, it is treated as omitted and `\` is used.
'* EscSeqChar can specify only one character, except for an empty string.
'* Unknown escape sequences drop the escape character and keep only the following character.
'* An escape character at the end of the string is an error.
Public Function UnescapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "\") As String
    If EscSeqChar = "" Then EscSeqChar = "\"
    Call pValidateLineSeparatorEscSeqChar(EscSeqChar, "Function UnescapeLineSeparator")

    Dim result_value As String
    Dim item_idx As Long
    Dim str_len As Long
    Dim cur_char As String

    str_len = Len(Expression)
    item_idx = 1
    result_value = ""

    Do While item_idx <= str_len
        cur_char = Mid$(Expression, item_idx, 1)

        If cur_char = EscSeqChar Then
            ' Start of an escape sequence.
            If item_idx < str_len Then
                Dim next_char As String
                next_char = Mid$(Expression, item_idx + 1, 1)

                Select Case next_char
                    Case "n"
                        result_value = result_value & vbLf
                        item_idx = item_idx + 2
                    Case "r"
                        result_value = result_value & vbCr
                        item_idx = item_idx + 2
                    Case Else
                        result_value = result_value & next_char
                        item_idx = item_idx + 2
                End Select
            Else
                ' Escape character at the end of the string.
                Err.Raise vbObjectError + 512, "Function UnescapeLineSeparator", "The string ends with an escape character."
            End If
        Else
            ' If it is not an escape character, append it as-is.
            result_value = result_value & cur_char
            item_idx = item_idx + 1
        End If
    Loop

    UnescapeLineSeparator = result_value
End Function

'* Splits a string by newline characters and returns an array.
'*
'* @param StringList Target string.
'* @return String array split by newline characters.
'*
'* @details
'* Splits a string based on newline characters and returns it as an array.
Public Function SplitByLineSeparator(ByVal StringList As String) As String()
    SplitByLineSeparator = Split(pUnifyLineSeparatorCore(StringList), vbLf)
End Function

'* Normalizes newline characters.
'*
'* @param Expression Target string.
'* @param LineSep Newline character after normalization (default is `vbLf`).
'* @return Normalized string.
'*
'* @details
'* Replaces newline characters with the specified sequence (`vbLf`, `vbCr`, `vbCrLf`).
Public Function UnifyLineSeparator(ByVal Expression As String, Optional ByVal LineSep As String = vbLf) As String
    If LineSep <> vbLf And LineSep <> vbCr And LineSep <> vbCrLf Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function UnifyLineSeparator", Description:="The value is not a newline character. (" & LineSep & ")"
    End If

    UnifyLineSeparator = Replace(pUnifyLineSeparatorCore(Expression), vbLf, LineSep)
End Function

Private Function pUnifyLineSeparatorCore(ByVal Expression As String) As String
    pUnifyLineSeparatorCore = Replace(Replace(Expression, vbCrLf, vbLf), vbCr, vbLf)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' String joining
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Joins an ObjectList containing strings into one String separated by Delimiter.
'*
'* @param SourceList ObjectList containing strings.
'* @param Delimiter Delimiter (default is a half-width space).
'* @return String joined with the delimiter.
'*
'* @details
'* Joins the specified string list with the delimiter and returns it as a single string.
Public Function JoinStringList(ByVal SourceList As ObjectList, Optional ByVal Delimiter As String = " ") As String
    Dim result_value As String

    If SourceList.Count = 0 Then Exit Function

    result_value = SourceList.Item(0)

    Dim item_idx As Long
    For item_idx = 1 To SourceList.Count - 1
        result_value = result_value & Delimiter & SourceList.Item(item_idx)
    Next item_idx

    JoinStringList = result_value
End Function

'* Joins an ObjectSet containing strings into one String separated by Delimiter.
'*
'* @param SourceSet ObjectSet containing strings.
'* @param Delimiter Delimiter (default is a half-width space).
'* @return String joined with the delimiter.
'*
'* @details
'* Joins the specified string set with the delimiter and returns it as a single string.
Public Function JoinStringSet(ByVal SourceSet As ObjectSet, Optional ByVal Delimiter As String = " ") As String
    Dim result_value As String

    If SourceSet.Count = 0 Then Exit Function

    result_value = SourceSet.Item(0)

    Dim item_idx As Long
    For item_idx = 1 To SourceSet.Count - 1
        result_value = result_value & Delimiter & SourceSet.Item(item_idx)
    Next item_idx

    JoinStringSet = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Message splitting
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* MsgBox that splits a message string into pages and displays them.
'*
'* @param MessageString Message string to display.
'* @param Title MsgBox title (optional).
'*
'* @details
'* Splits the specified message into pages and displays each page with MsgBox.
Public Sub MsgBoxPage(ByVal MessageString As String, Optional ByVal Title As String = "")
    Dim msg_pages() As String
    Dim msg_page As Variant 'String
    Dim total_num As Long
    Dim page_num As Long

    msg_pages = SplitMessage(MessageString)

    total_num = UBound(msg_pages) - LBound(msg_pages) + 1

    page_num = 1
    For Each msg_page In msg_pages
        MsgBox msg_page, Title:=Title & "(" & page_num & "/" & total_num & ")"
        page_num = page_num + 1
    Next msg_page
End Sub

'* Splits a message string into pages.
'*
'* @param MessageString Target string.
'* @param PageSize Maximum page size in bytes (default is 1023).
'* @return Array of strings split into pages.
'*
'* @details
'* Splits the specified message string by newline and paginates it within the specified byte count.
Public Function SplitMessage(ByVal MessageString As String, Optional ByVal PageSize As Long = 1023) As Variant
    If PageSize < 4 Then Err.Raise vbObjectError + 1, "Function SplitMessage", "PageSize must be 4 or greater. (" & PageSize & ")"

    Dim msgs_list As ObjectList
    Dim msg_str As String
    Dim has_msg_str As Boolean
    Dim lines_arr() As String
    Dim line_str As Variant 'String
    Dim taken_str As String
    Dim rem_str As String
    Dim result_value() As String
    Dim item_idx As Long

    Set msgs_list = New ObjectList

    If MessageString = "" Then
        ReDim result_value(0 To 0)
        result_value(0) = ""
        SplitMessage = result_value
        Exit Function
    End If

    lines_arr() = SplitByLineSeparator(MessageString)

    For Each line_str In lines_arr
        If (Not has_msg_str And LenB(StrConv(line_str, vbFromUnicode)) <= PageSize) _
                Or (has_msg_str And LenB(StrConv(msg_str, vbFromUnicode)) + 2 + LenB(StrConv(line_str, vbFromUnicode)) <= PageSize) Then
            If Not has_msg_str Then
                msg_str = line_str
                has_msg_str = True
            Else
                msg_str = msg_str & vbCrLf & line_str
            End If
        Else
            If has_msg_str Then
                Call msgs_list.Add(msg_str)
            End If

            Do While LenB(StrConv(line_str, vbFromUnicode)) > PageSize
                Call pTakeString(taken_str, rem_str, line_str, PageSize)
                Call msgs_list.Add(taken_str)
                line_str = rem_str
            Loop

            msg_str = line_str
            has_msg_str = True
        End If
    Next line_str

    If has_msg_str Then Call msgs_list.Add(msg_str)

    ReDim result_value(0 To msgs_list.Count - 1)
    For item_idx = LBound(result_value) To UBound(result_value)
        result_value(item_idx) = msgs_list.Item(item_idx)
    Next item_idx

    SplitMessage = result_value
End Function

Private Sub pTakeString( _
        ByRef TakenString As String, _
        ByRef RemainingString As String, _
        ByVal Expression As String, _
        ByVal LengthByte As Long)

    Dim char_text As String
    Dim item_idx As Long
    Dim result_value As String
    Dim total_bytes As Long
    Dim current_bytes As Long

    For item_idx = 1 To Len(Expression)
        char_text = Mid(Expression, item_idx, 1)
        current_bytes = LenB(StrConv(char_text, vbFromUnicode))

        If total_bytes + current_bytes > LengthByte Then Exit For

        result_value = result_value + char_text
        total_bytes = total_bytes + current_bytes
    Next item_idx

    TakenString = result_value
    RemainingString = Mid(Expression, item_idx)
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Trimming
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Removes leading and trailing half-width spaces, tabs, and newline characters from a string.
'*
'* @param Expression Target string.
'* @param IgnoreHead True to ignore leading spaces (default is False).
'* @param IgnoreTail True to ignore trailing spaces (default is False).
'* @param RemoveFullWidthSpace True to also remove full-width spaces (default is False).
'* @return String with leading and trailing spaces removed.
'*
'* @details
'* Removes whitespace from the beginning and end of a string based on the specified conditions.
Public Function Strip( _
        ByVal Expression As String, _
        Optional ByVal IgnoreHead As Boolean = False, _
        Optional ByVal IgnoreTail As Boolean = False, _
        Optional ByVal RemoveFullWidthSpace As Boolean = False) As String

    Dim head_pos As Long
    Dim tail_pos As Long
    Dim test_char As String

    head_pos = 1
    tail_pos = Len(Expression)

    If tail_pos = 0 Then
        Strip = Expression
        Exit Function
    End If

    If Not IgnoreHead Then
        ' Remove leading spaces, tabs, and newlines.
        Do While head_pos <= tail_pos
            test_char = Mid(Expression, head_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If

            head_pos = head_pos + 1
        Loop
    End If

    If Not IgnoreTail Then
        ' Remove trailing spaces, tabs, and newlines.
        Do While tail_pos >= head_pos
            test_char = Mid(Expression, tail_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If

            tail_pos = tail_pos - 1
        Loop
    End If

    ' Return the final string.
    Strip = Mid(Expression, head_pos, tail_pos - head_pos + 1)
End Function

Private Function pIsWhitespace(ByVal Character As String, ByVal IncludeFullWidthSpace As Boolean) As Boolean
    Dim result_value As Boolean

    If Character = " " Then
        result_value = True
    ElseIf Character = vbTab Then
        result_value = True
    ElseIf Character = vbCrLf Then
        result_value = True
    ElseIf Character = vbCr Then
        result_value = True
    ElseIf Character = vbLf Then
        result_value = True
    ElseIf (Character = "Å@" And IncludeFullWidthSpace) Then
        result_value = True
    Else
        result_value = False
    End If

    pIsWhitespace = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' String checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Determines whether a string starts with the specified string.
'*
'* @param Expression String to test.
'* @param SearchString String to search for.
'* @return True if the string starts with the specified string; otherwise, False.
'*
'* @details
'* Determines whether the string to test starts with the specified string. Returns True when SearchString is an empty string.
Public Function StartsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If SearchString = "" Then
        StartsWith = True
        Exit Function
    End If

    If Len(Expression) < Len(SearchString) Then
        StartsWith = False
        Exit Function
    End If

    If Left(Expression, Len(SearchString)) = SearchString Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

'* Determines whether a string ends with the specified string.
'*
'* @param Expression String to test.
'* @param SearchString String to search for.
'* @return True if the string ends with the specified string; otherwise, False.
'*
'* @details
'* Determines whether the string to test ends with the specified string. Returns True when SearchString is an empty string.
Public Function EndsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If SearchString = "" Then
        EndsWith = True
        Exit Function
    End If

    If Len(Expression) < Len(SearchString) Then
        EndsWith = False
        Exit Function
    End If

    If Right(Expression, Len(SearchString)) = SearchString Then
        EndsWith = True
    Else
        EndsWith = False
    End If
End Function

'* Determines whether a string is enclosed by the specified strings.
'*
'* @param Expression String to test.
'* @param QuoteString Start string.
'* @param EndString End string (default is `QuoteString`). If the end string is omitted, the same value as the start string is used.
'* @return True if enclosed by the specified strings; otherwise, False.
'*
'* @details
'* Determines whether the string to test is enclosed by the specified strings.
Public Function IsQuotedWith(ByVal Expression As String, ByVal QuoteString As String, Optional ByVal EndString As String = "") As Boolean
    If QuoteString = "" Then
        IsQuotedWith = False
        Exit Function
    End If

    If EndString = "" Then
        EndString = QuoteString
    End If

    If Len(Expression) < 2 Or Len(Expression) < (Len(QuoteString) + Len(EndString)) Then
        IsQuotedWith = False
        Exit Function
    End If

    If Left(Expression, Len(QuoteString)) = QuoteString And Right(Expression, Len(EndString)) = EndString Then
        IsQuotedWith = True
    Else
        IsQuotedWith = False
    End If
End Function


' #############################################################################
'
' Numbers
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Maximum and minimum values
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Returns the largest Long value among the arguments passed.
'*
'* @param Number1 First number.
'* @param Number2 Second number.
'* @param Numbers Other numbers (variable-length arguments).
'* @return Largest number.
'*
'* @details
'* Specify at least two arguments, plus any number of additional numbers as variable-length arguments.
Public Function MaxLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result_value As Long

    result_value = Number1
    If result_value < Number2 Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If result_value < Numbers(item_idx) Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MaxLng = result_value
End Function

'* Returns the largest Double value among the arguments passed.
'*
'* @param Number1 First number.
'* @param Number2 Second number.
'* @param Numbers Other numbers (variable-length arguments).
'* @return Largest number.
'*
'* @details
'* Specify at least two arguments, plus any number of additional numbers as variable-length arguments.
Public Function MaxDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result_value As Double

    result_value = Number1
    If result_value < Number2 Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If result_value < Numbers(item_idx) Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MaxDbl = result_value
End Function

'* Returns the smallest Long value among the arguments passed.
'*
'* @param Number1 First number.
'* @param Number2 Second number.
'* @param Numbers Other numbers (variable-length arguments).
'* @return Smallest number.
'*
'* @details
'* Specify at least two arguments, plus any number of additional numbers as variable-length arguments.
Public Function MinLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result_value As Long

    result_value = Number1
    If Number2 < result_value Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(item_idx) < result_value Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MinLng = result_value
End Function

'* Returns the smallest Double value among the arguments passed.
'*
'* @param Number1 First number.
'* @param Number2 Second number.
'* @param Numbers Other numbers (variable-length arguments).
'* @return Smallest number.
'*
'* @details
'* Specify at least two arguments, plus any number of additional numbers as variable-length arguments.
Public Function MinDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result_value As Double

    result_value = Number1
    If Number2 < result_value Then result_value = Number2

    If UBound(Numbers) <> -1 Then
        Dim item_idx As Long
        For item_idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(item_idx) < result_value Then result_value = Numbers(item_idx)
        Next item_idx
    End If

    MinDbl = result_value
End Function


' #############################################################################
'
' Integer and Long
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Long bit strings
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Converts a Long value to a binary notation string.
'*
'* @param LongValue Long value to convert.
'* @return Binary notation string.
'*
'* @details
'* Converts the specified Long value to a binary notation string.
Public Function LongToBin(ByVal LongValue As Long) As String
    Dim high_bit As String
    Dim long_value As Long
    Dim item_idx As Long
    Dim result_value As String

    If 0 <= LongValue Then
        high_bit = "0"
        long_value = LongValue
    Else
        high_bit = "1"
        long_value = LongValue And &H7FFFFFFF
    End If

    For item_idx = 1 To 31
        result_value = (long_value Mod 2) & result_value
        long_value = long_value \ 2
    Next item_idx

    result_value = high_bit & result_value

    LongToBin = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Bit shifting
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Sub Test_BitShift()
'    Dim test_values(0 To 7) As Long
'    Dim arr_idx As Integer
'    'For arr_idx = 0 To 3
'    '    test_values(arr_idx) = 32767& * Rnd
'    '    test_values(arr_idx) = test_values(arr_idx) * (2 ^ 16) + 32767& * Rnd
'    '    If Rnd < 0.5 Then test_values(arr_idx) = -test_values(arr_idx) - 1
'    'next
'
'    test_values(0) = &H80010003
'    test_values(1) = &H80010002
'    test_values(2) = &HE0010003
'    test_values(3) = &HE0010002
'    test_values(4) = &H40010003
'    test_values(5) = &H40010002
'    test_values(6) = &H70010003
'    test_values(7) = &H70010002
'
'    arr_idx = 0 ' Change within the range 0 to 7.
'
'    Dim test_value As Long
'    test_value = test_values(arr_idx)
'
'    'Debug.Print "----------------------"
'    'Debug.Print "left"
'    'Debug.Print test_value & "(" & LongToBin(test_value) & ")"
'
'    Dim item_idx As Long
'    For item_idx = 33 To -33 Step -1
'        'Debug.Print LongToBin(BitLeft(test_value, item_idx)) & " | " & item_idx
'    Next
'    'Debug.Print "right (arithmetic)"
'    For item_idx = 0 To 33
'        'Debug.Print LongToBin(BitRight(test_value, item_idx, Arithmetic:=True)) & " | " & item_idx
'    Next
'End Sub

'* Function that performs a bit left shift operation.
'*
'* @param TargetValue Value to shift left.
'* @param ShiftCount Number of bits to shift.
'* @return Result value after left shift.
'*
'* @details
'* Performs a left bit shift operation on the Long value specified as an argument.
'* If a negative shift count is specified, it is processed as a logical right shift.
'* Returns 0 when ShiftCount is 32 or greater, or -32 or less.
Public Function BitLeft(ByVal TargetValue As Long, ByVal ShiftCount As Long) As Long
    ' If the shift count is 0, return as-is.
    If ShiftCount = 0 Then
        BitLeft = TargetValue
        Exit Function
    End If

    ' If the shift count is 32 or greater, everything disappears.
    If 31 < ShiftCount Then
        BitLeft = &H0&
        Exit Function
    End If

    ' If the shift count is -32 or less, everything disappears due to a logical right shift in the opposite direction.
    If ShiftCount < -31 Then
        BitLeft = &H0&
        Exit Function
    End If

    ' If the shift count is negative, process it as a logical right shift.
    If ShiftCount < 0 Then
        BitLeft = BitRight(TargetValue, -ShiftCount)
        Exit Function
    End If

    ' If ShiftCount is 31, handle it first because it would overflow.
    If ShiftCount = 31 Then
        ' When ShiftCount is 31, the rightmost bit of TargetValue becomes the sign bit (leftmost).
        ' The others become 0.
        If TargetValue Mod 2 = 0 Then
            ' If even, the rightmost bit is 0, so the sign bit is 0.
            BitLeft = &H0&
        Else
            ' If odd, the rightmost bit is 1, so the sign bit is 1.
            BitLeft = &H80000000
        End If
        Exit Function
    End If

    ' Calculate the subsequent shift counts from 1 to 30.

    Dim result_value As Long

   ' Prepare a mask to prevent overflow.
   ' Processing to get the same value as &HFFFFFFFF >> (ShiftCount + 1) in other languages.
   ' Used to zero out the upper bits that should disappear and the bit that becomes the sign bit after processing.
    Dim mask_value As Long
    mask_value = (2& ^ (31 - ShiftCount)) - 1

    ' Shift processing.
    ' Simulate TargetValue << ShiftCount in other languages by multiplying by a power of 2.
    result_value = (TargetValue And mask_value) * (2 ^ ShiftCount)

    ' Prepare a mask for getting the most significant bit.
    ' Processing to get the same value as &H1& << (32 - ShiftCount) in other languages.
    ' Used to get the bit that should become the sign bit after processing.
    Dim high_bit_mask As Long
    high_bit_mask = &H1& * (2& ^ (31 - ShiftCount))

    ' Process the sign bit (most significant bit).
    ' If nonzero, set the sign bit.
    If (TargetValue And high_bit_mask) <> 0 Then
        result_value = result_value Or &H80000000
    End If

    BitLeft = result_value
End Function

'* Function that performs a bit right shift operation.
'*
'* @param TargetValue Value to shift right.
'* @param ShiftCount Number of bits to shift.
'* @param Arithmetic True to perform an arithmetic shift. Default is False.
'* @return Result value after right shift.
'*
'* @details
'* Performs a right bit shift operation on the Long value specified as an argument.
'* If a negative shift count is specified, it is processed as a left shift.
'* If an arithmetic shift is requested, the left side is filled with the sign bit.
'* Returns 0 when ShiftCount is -32 or less.
'* When ShiftCount is 32 or greater, logical shift returns 0 and arithmetic shift returns a value filled with the sign bit.
Public Function BitRight(ByVal TargetValue As Long, ByVal ShiftCount As Long, Optional ByVal Arithmetic As Boolean = False) As Long
    ' If the shift count is 0, return as-is.
    If ShiftCount = 0 Then
        BitRight = TargetValue
        Exit Function
    End If

    ' If the shift count is 32 or greater.
    If 31 < ShiftCount Then
        If Arithmetic And TargetValue < 0 Then
            ' Everything is filled with the sign bit.
            BitRight = &HFFFFFFFF
        Else
            ' Everything disappears.
            BitRight = 0
        End If
        Exit Function
    End If

    ' If the shift count is -32 or less, everything disappears due to a left shift in the opposite direction.
    If ShiftCount < -31 Then
        BitRight = &H0&
        Exit Function
    End If

    ' If the shift count is negative, process it as a left shift.
    If ShiftCount < 0 Then
        BitRight = BitLeft(TargetValue, -ShiftCount)
        Exit Function
    End If

    ' If ShiftCount is 31, handle it first because it would overflow.
    If ShiftCount = 31 Then
        If TargetValue < 0 Then
            If Arithmetic Then
                ' 31 bits are filled with the sign bit (1), and the 32nd bit is the sign bit itself (1).
                BitRight = &HFFFFFFFF
            Else
                ' 31 bits are filled with 0, and the 32nd bit is the sign bit itself (1).
                BitRight = &H1&
            End If
        Else
            ' 31 bits are filled with the sign bit (0), and the 32nd bit is the sign bit itself (0).
            ' 31 bits are filled with 0, and the 32nd bit is the sign bit itself (0).
            ' In any case, all bits are 0.
            BitRight = &H0&
        End If
        Exit Function
    End If

    ' Calculate the subsequent shift counts from 1 to 30.

    Dim result_value As Long
    ' First, clear the sign bit.
    result_value = TargetValue And &H7FFFFFFF

    ' Shift once.
    ' Simulate a right shift by dividing by 2 once within the integer range.
    result_value = result_value \ 2

    ' Shift the sign bit once.
    ' When the sign bit is 1 (that is, negative), simulate it by setting the second bit.
    If TargetValue < 0 Then
        result_value = result_value Or &H40000000
    End If

    ' Remaining shifts.
    ' Simulate by dividing by 2 raised to the power of (ShiftCount - 1) within the integer range.
    ' When the shift count is 1 this work is wasted, but it avoids adding a comparison that is unnecessary for other counts.
    result_value = result_value \ (2 ^ (ShiftCount - 1))

    If Arithmetic And TargetValue < 0 Then
        ' Processing to get the same value as &HFFFFFFFF << (32 - ShiftCount - 1) in other languages.
        Dim sign_mask As Long
        sign_mask = Not ((2& ^ (32 - ShiftCount)) - 1)

        ' Fill with 1s the area that became 0 because of the logical shift.
        result_value = result_value Or sign_mask
    End If

    BitRight = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Unsigned Long comparison
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Function that compares Long values as unsigned integers.
'*
'* @param ValueA First value to compare.
'* @param ValueB Second value to compare.
'* @return Comparison result. 0 when A = B, -1 when A < B, and 1 when A > B.
'*
'* @details
'* Compares value A and value B as unsigned integers.
Public Function CompareAsUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Integer
    If ValueA = ValueB Then
        CompareAsUnsignedLong = 0
    Else
        If 0 <= ValueA Then
            If 0 <= ValueB Then
                ' When both are positive, compare normally.
                If ValueA < ValueB Then
                    CompareAsUnsignedLong = -1
                Else
                    CompareAsUnsignedLong = 1
                End If
            Else
                ' When A is positive and B is negative, A < B.
                CompareAsUnsignedLong = -1
            End If
        Else
            If 0 <= ValueB Then
                ' When A is negative and B is positive, A > B.
                CompareAsUnsignedLong = 1
            Else
                ' When both are negative, compare normally.
                If ValueA < ValueB Then
                    CompareAsUnsignedLong = -1
                Else
                    CompareAsUnsignedLong = 1
                End If
            End If
        End If
    End If
End Function

'* Determines the A < B relationship as unsigned integers.
'*
'* @param ValueA First value to compare.
'* @param ValueB Second value to compare.
'* @return True when A is less than B; otherwise, False.
'*
'* @details
'* Compares two Long values as unsigned integers and returns True when A < B.
'* The result of this function matches when the `CompareAsUnsignedLong` function returns a negative value.
Public Function IsLessThanUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Boolean
    If CompareAsUnsignedLong(ValueA, ValueB) < 0 Then
        IsLessThanUnsignedLong = True
    Else
        IsLessThanUnsignedLong = False
    End If
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Unsigned Long addition and subtraction
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Adds two Long values as unsigned integers.
'*
'* @param ValueA First value to add.
'* @param ValueB Second value to add.
'* @return Addition result treated as unsigned.
'*
'* @details
'* Treats two Long values as unsigned integers and adds them.
'* Raises an error if the addition result exceeds `FFFFFFFF`.
Public Function AddUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    Dim a_high As Long
    Dim a_low As Long
    Call pSeparateUnsignedLong(a_high, a_low, ValueA)

    Dim b_high As Long
    Dim b_low As Long
    Call pSeparateUnsignedLong(b_high, b_low, ValueB)

    Dim carry_bit_l As Long
    Dim add_result_l As Long
    Call pAddUnsignedLongCore(add_result_l, carry_bit_l, a_low, b_low)

    Dim carry_bit_ha As Long
    Dim add_result_ha As Long
    Call pAddUnsignedLongCore(add_result_ha, carry_bit_ha, a_high, carry_bit_l)

    If 0 < carry_bit_ha Then Err.Raise vbObjectError + 1, "Function AddUnsignedLong", "The addition result exceeds FFFFFFFF. (" & Hex(ValueA) & " + " & Hex(ValueB) & ")"

    Dim carry_bit_h As Long
    Dim add_result_h As Long
    Call pAddUnsignedLongCore(add_result_h, carry_bit_h, add_result_ha, b_high)
    If 0 < carry_bit_h Then Err.Raise vbObjectError + 1, "Function AddUnsignedLong", "The addition result exceeds FFFFFFFF. (" & Hex(ValueA) & " + " & Hex(ValueB) & ")"

    Dim high_bit As Boolean
    If 32767 < add_result_h Then
        high_bit = True
        add_result_h = add_result_h And &H7FFF&
    Else
        high_bit = False
    End If

    add_result_h = add_result_h * 65536

    Dim result_value As Long
    result_value = add_result_h + add_result_l

    If high_bit Then
        result_value = result_value Or &H80000000
    End If

    AddUnsignedLong = result_value
End Function

Private Function pCombineUnsignedLong(ByVal HighPart As Long, ByVal LowPart As Long) As Long
    Dim high_bit As Boolean
    If 32767 < HighPart Then
        high_bit = True
        HighPart = HighPart And &H7FFF&
    Else
        high_bit = False
    End If

    Dim result_value As Long
    result_value = HighPart * 65536 + LowPart
    If high_bit Then
        result_value = result_value Or &H80000000
    End If

    pCombineUnsignedLong = result_value
End Function

Private Sub pSeparateUnsignedLong(ByRef HighPart As Long, ByRef LowPart As Long, ByVal TargetValue As Long)
    Dim high_bit As Boolean
    If TargetValue < 0 Then
        high_bit = True
        TargetValue = TargetValue And &H7FFFFFFF
    Else
        high_bit = False
    End If

    LowPart = TargetValue And &HFFFF&

    TargetValue = TargetValue \ 65536

    If high_bit Then
        HighPart = TargetValue Or &H8000&
    Else
        HighPart = TargetValue
    End If
End Sub

Private Sub pAddUnsignedLongCore(ByRef ResultValue As Long, ByRef CarryBit As Long, ByVal ValueA As Long, ByVal ValueB As Long)
    Dim result_value As Long
    result_value = ValueA + ValueB
    If 65535 < result_value Then
        ResultValue = result_value - 65536
        CarryBit = 1
    Else
        ResultValue = result_value
        CarryBit = 0
    End If
End Sub

'* Subtracts two Long values as unsigned integers.
'*
'* @param ValueA Value to subtract from.
'* @param ValueB Value to subtract.
'* @return Subtraction result treated as unsigned.
'*
'* @details
'* Treats two Long values as unsigned integers and subtracts them.
'* Raises an error when the subtraction result is negative (`A < B`).
Public Function SubtractUnsignedLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    If ValueA = ValueB Then
        SubtractUnsignedLong = 0
        Exit Function
    End If

    If IsLessThanUnsignedLong(ValueA, ValueB) Then
        Err.Raise vbObjectError + 1, "Function SubtractUnsignedLong", "The second argument is greater than the first argument, so the result would be negative. (a: " & Hex(ValueA) & ", b: " & Hex(ValueB) & ")"
        Exit Function
    End If

    Dim a_high As Long
    Dim a_low As Long
    Call pSeparateUnsignedLong(a_high, a_low, ValueA)

    Dim b_high As Long
    Dim b_low As Long
    Call pSeparateUnsignedLong(b_high, b_low, ValueB)

    Dim borrow_bit_l As Long
    Dim subtract_result_l As Long
    If a_low < b_low Then
        subtract_result_l = a_low + 65536 - b_low
        borrow_bit_l = 1
    Else
        subtract_result_l = a_low - b_low
        borrow_bit_l = 0
    End If

    Dim subtract_result_h As Long
    subtract_result_h = a_high - b_high - borrow_bit_l
    If subtract_result_h < 0 Then
        Err.Raise vbObjectError + 1, "Function SubtractUnsignedLong", "The second argument is greater than the first argument, so the result would be negative. (a: " & Hex(ValueA) & ", b: " & Hex(ValueB) & ")"
    End If

    SubtractUnsignedLong = pCombineUnsignedLong(subtract_result_h, subtract_result_l)
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Integer type checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Determines whether the specified value can be converted to an integer (`Integer` type).
'*
'* @param Value Value to test.
'* @return True if the value can be converted to an integer; otherwise, False.
'*
'* @details
'* Returns True when the specified value is numeric and within the range convertible to the `Integer` type.
'* Returns False when the value is not numeric or is out of range.
Public Function IsInteger(ByVal Value As Variant) As Boolean
    Dim numeric_value As Double
    If Not pTryConvertToDouble(numeric_value, Value) Then
        IsInteger = False
        Exit Function
    End If

    If numeric_value < C_INTEGER_MIN Or C_INTEGER_MAX < numeric_value Then
        IsInteger = False
        Exit Function
    End If

    IsInteger = (numeric_value = Fix(numeric_value))
End Function

'* Determines whether the specified value can be converted to a long integer (`Long` type).
'*
'* @param Value Value to test.
'* @return True if the value can be converted to a long integer; otherwise, False.
'*
'* @details
'* Returns True when the specified value is numeric and within the range convertible to the `Long` type.
'* Returns False when the value is not numeric or is out of range.
Public Function IsLong(ByVal Value As Variant) As Boolean
    Dim numeric_value As Double
    If Not pTryConvertToDouble(numeric_value, Value) Then
        IsLong = False
        Exit Function
    End If

    If numeric_value < C_LONG_MIN Or C_LONG_MAX < numeric_value Then
        IsLong = False
        Exit Function
    End If

    IsLong = (numeric_value = Fix(numeric_value))
End Function

Private Function pTryConvertToDouble(ByRef ResultValue As Double, ByVal Value As Variant) As Boolean
    If Not IsNumeric(Value) Then
        pTryConvertToDouble = False
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    ResultValue = CDbl(Value)
    If Err.Number <> 0 Then
        Err.Clear
        pTryConvertToDouble = False
    Else
        pTryConvertToDouble = True
    End If
    On Error GoTo 0
End Function


' #############################################################################
'
' Excel
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel file format checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Gets the file format from the extension of a file path.
'*
'* @param FileNameOrPath File path string.
'* @return File format. If it cannot be determined, returns xlOpenXMLWorkbook.
'*
'* @details
'* Gets the file format from the extension of a file path.
'* Extension case is ignored.
Public Function GetExcelFileFormat(ByVal FileNameOrPath As String) As Long
    Select Case LCase$(GetLeafFromPath(FileNameOrPath, BaseName:=False, Extension:=True))
     Case ".xlsm"
        GetExcelFileFormat = xlOpenXMLWorkbookMacroEnabled
     Case ".xltm"
        GetExcelFileFormat = xlOpenXMLTemplateMacroEnabled
     Case ".xls"
        GetExcelFileFormat = xlExcel8
     Case ".xla"
        GetExcelFileFormat = xlAddIn8
     Case ".xlam"
        GetExcelFileFormat = xlOpenXMLAddIn
     Case ".xlsb"
        GetExcelFileFormat = xlExcel12
     Case ".xlt"
        GetExcelFileFormat = xlTemplate8
     Case ".xltx"
        GetExcelFileFormat = xlOpenXMLTemplate
     Case ".xlw"
        GetExcelFileFormat = xlExcel4Workbook
     Case ".csv"
        GetExcelFileFormat = xlCSV
     Case ".txt"
        GetExcelFileFormat = xlCurrentPlatformText
     Case ".dbf"
        GetExcelFileFormat = xlDBF4
     Case ".dif"
        GetExcelFileFormat = xlDIF
     Case ".htm"
        GetExcelFileFormat = xlHtml
     Case ".html"
        GetExcelFileFormat = xlHtml
     Case ".ods"
        GetExcelFileFormat = xlOpenDocumentSpreadsheet
     Case ".slk"
        GetExcelFileFormat = xlSYLK
     Case Else
        GetExcelFileFormat = xlOpenXMLWorkbook
    End Select
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel address type checks
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Checks whether a Range address is a multiple selection (for example, A1, B2:C3, D4).
'*
'* @param AddressString Excel address string to test.
'* @return True when multiple ranges are selected; otherwise, False.
'*
'* @details
'* Checks whether a Range address is a multiple selection (for example, A1, B2:C3, D4).
Public Function IsMultiRange(ByVal AddressString As String) As Boolean
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, AddressString)

    IsMultiRange = 0 < InStr(cell_address, ",")
End Function

'* Checks whether a Range address is an Area (a contiguous multi-cell range excluding a single cell).
'*
'* @param AddressString Excel address string to test.
'* @return True if it is an Area; otherwise, False.
'*
'* @details
'* In this common module, Area means a contiguous multi-cell range excluding a single cell.
'* Non-rectangular multiple-selection ranges return False.
Public Function IsArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsArea = range_bounds.IsArea
End Function

'* Checks whether a Range address is a Cell (single-cell range).
'*
'* @param AddressString Excel address string to test.
'* @return True for a single-cell range; otherwise, False.
'*
'* @details
'* Checks whether a Range address is a Cell (for example, A1).
'* Non-rectangular multiple-selection ranges return False.
Public Function IsCell(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsCell = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsCell = range_bounds.IsCell
End Function

'* Checks whether a Range address is an entire row (for example, 1:2).
'*
'* @param AddressString Excel address string to test.
'* @return True when it is an entire row; otherwise, False.
'*
'* @details
'* Checks whether a Range address is an entire row (for example, 1:2).
'* Non-rectangular multiple-selection ranges return False.
Public Function IsEntireRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireRow = range_bounds.IsEntireRow
End Function

'* Checks whether a Range address is an entire column (for example, A:B).
'*
'* @param AddressString Excel address string to test.
'* @return True when it is an entire column; otherwise, False.
'*
'* @details
'* Checks whether a Range address is an entire column (for example, A:B).
'* Non-rectangular multiple-selection ranges return False.
Public Function IsEntireColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireColumn = range_bounds.IsEntireColumn
End Function

'* Checks whether a Range address has a one-row shape (for example, A1, A1:B1, or 1:1).
'*
'* @param AddressString Excel address string to test.
'* @return True for a one-row shape; otherwise, False.
'*
'* @details
'* A single cell is also True as a one-row shape.
'* Non-rectangular multiple-selection ranges return False.
Public Function IsOneRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRow = range_bounds.IsOneRow
End Function

'* Checks whether a Range address has a one-column shape (for example, A1, A1:A2, or A:A).
'*
'* @param AddressString Excel address string to test.
'* @return True for a one-column shape; otherwise, False.
'*
'* @details
'* A single cell is also True as a one-column shape.
'* Non-rectangular multiple-selection ranges return False.
Public Function IsOneColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumn = range_bounds.IsOneColumn
End Function

'* Checks whether a Range address is an Area consisting of only one row.
'*
'* @param AddressString Excel address string to test.
'* @return True for an Area consisting of only one row; otherwise, False.
'*
'* @details
'* A1:B1 and 1:1 are True; A1 is False because it is a Cell.
'* Non-rectangular multiple-selection ranges return False.
Public Function IsOneRowArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRowArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRowArea = range_bounds.IsOneRowArea
End Function

'* Checks whether a Range address is an Area consisting of only one column.
'*
'* @param AddressString Excel address string to test.
'* @return True for an Area consisting of only one column; otherwise, False.
'*
'* @details
'* A1:A2 and A:A are True; A1 is False because it is a Cell.
'* Non-rectangular multiple-selection ranges return False.
Public Function IsOneColumnArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumnArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumnArea = range_bounds.IsOneColumnArea
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel address generation
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Gets a range address string with specified parameters.
'*
'* @param StartRow Start row number (unspecified when omitted).
'* @param StartColumn Start column number (unspecified when omitted).
'* @param FinishRow Finish row number (unspecified when omitted).
'* @param FinishColumn Finish column number (unspecified when omitted).
'* @param IsAbsoluteStartRow True to specify the start row as an absolute reference (default is False).
'* @param IsAbsoluteStartColumn True to specify the start column as an absolute reference (default is False).
'* @param IsAbsoluteFinishRow True to specify the finish row as an absolute reference (default is False).
'* @param IsAbsoluteFinishColumn True to specify the finish column as an absolute reference (default is False).
'* @param ReferenceRow Reference row number for relative references (default is 1).
'* @param ReferenceColumn Reference column number for relative references (default is 1).
'* @param AddressType Address format ("A1" or "R1C1") (default is "A1").
'* @param SheetName Sheet name (unspecified when omitted).
'* @param BookName Workbook name (unspecified when omitted).
'* @return Range address string generated based on the specified conditions.
'*
'* @details
'* Generates an Excel range address as a string by specifying parameters such as row, column, and absolute / relative references.
'*
'* If row specification is omitted, it is treated as a column range. If column specification is omitted, it is treated as a row range.
'* Row specification and column specification cannot both be omitted.
'* When BookName is specified, SheetName must also be specified.
Public Function RangeAddress( _
        Optional ByVal StartRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal StartColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal IsAbsoluteStartRow As Boolean = False, _
        Optional ByVal IsAbsoluteStartColumn As Boolean = False, _
        Optional ByVal IsAbsoluteFinishRow As Boolean = False, _
        Optional ByVal IsAbsoluteFinishColumn As Boolean = False, _
        Optional ByVal ReferenceRow As Long = 0, _
        Optional ByVal ReferenceColumn As Long = 0, _
        Optional ByVal AddressType As String = "A1", _
        Optional ByVal SheetName As String = "", _
        Optional ByVal BookName As String = "") As String

    Call pValidateBookAndSheetAddress(BookName, SheetName, "Function RangeAddress")

    ' Check row and column specifications.
    If StartRow = G_OMIT_CELL_INDEX And StartColumn = G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Start row and start column cannot both be omitted."
        Exit Function
    End If

    If StartRow = G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Start row cannot be omitted by itself. (finish_row: " & FinishRow & ")"
        Exit Function
    ElseIf StartRow <> G_OMIT_CELL_INDEX And FinishRow = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartRow = IsAbsoluteFinishRow Then
            ' If absolute address settings match, fill in FinishRow.
            FinishRow = StartRow
        ElseIf Not IsAbsoluteFinishRow Then
            ' If IsAbsoluteFinishRow remains at its initial value (False), fill in FinishRow and IsAbsoluteFinishRow.
            FinishRow = StartRow
            IsAbsoluteFinishRow = IsAbsoluteStartRow
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Finish row cannot be omitted. (start_row: " & StartRow & ")"
            Exit Function
        End If
    End If

    If StartColumn = G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Start column cannot be omitted by itself. (finish_col: " & FinishColumn & ")"
        Exit Function
    ElseIf StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
            ' If absolute address settings match, fill in FinishColumn.
            FinishColumn = StartColumn
        ElseIf Not IsAbsoluteFinishColumn Then
            ' If IsAbsoluteFinishColumn remains at its initial value (False), fill in FinishColumn and IsAbsoluteFinishColumn.
            FinishColumn = StartColumn
            IsAbsoluteFinishColumn = IsAbsoluteStartColumn
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Finish column cannot be omitted. (start_col: " & StartColumn & ")"
            Exit Function
        End If
    End If

    Dim is_row_range As Boolean
    If StartColumn = G_OMIT_CELL_INDEX Then
        is_row_range = True
    Else
        is_row_range = False
    End If

    Dim is_col_range As Boolean
    If StartRow = G_OMIT_CELL_INDEX Then
        is_col_range = True
    Else
        is_col_range = False
    End If

    Dim is_cell_range As Boolean
    If StartRow <> G_OMIT_CELL_INDEX And FinishRow = StartRow And IsAbsoluteStartRow = IsAbsoluteFinishRow _
            And StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = StartColumn And IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
        is_cell_range = True
    Else
        is_cell_range = False
    End If

    Dim result_value As String
    result_value = ExcelBookAndSheetAddress(BookName, SheetName)

    Select Case AddressType
     Case "A1"
        If is_cell_range Then
            result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow)
        Else
            If is_row_range Then
                result_value = result_value & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & ":" & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            ElseIf is_col_range Then
                result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteFinishColumn, ReferenceColumn)
            Else
                result_value = result_value & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & _
                        ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteFinishColumn, ReferenceColumn) & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            End If
        End If
     Case "R1C1"
        If is_cell_range Then
            result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn)
        Else
            If is_row_range Then
                result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & ":R" & pR1C1RowAddressCore(FinishRow, IsAbsoluteFinishRow)
            ElseIf is_col_range Then
                result_value = result_value & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn) & ":C" & pR1C1ColumnAddressCore(FinishColumn, IsAbsoluteFinishColumn)
            Else
                result_value = result_value & "R" & pR1C1RowAddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1ColumnAddressCore(StartColumn, IsAbsoluteStartColumn) & _
                        ":" & "R" & pR1C1RowAddressCore(FinishRow, IsAbsoluteFinishRow) & "C" & pR1C1ColumnAddressCore(FinishColumn, IsAbsoluteFinishColumn)
            End If
        End If
     Case Else
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="AddressType must be A1 or R1C1. (" & AddressType & ")"
        Exit Function
    End Select

    RangeAddress = result_value
End Function

Private Function pA1columnAddressCore(ByVal ColumnIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceColumn As Long) As String
    Dim actual_column As Double
    If IsAbsolute Then
        actual_column = ColumnIndex
        If actual_column < 1 Or G_COL_MAX < actual_column Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Column index is out of range. (" & ColumnIndex & ")"
        End If

        pA1columnAddressCore = "$" & ExcelA1ColumnAddress(CLng(actual_column))
    Else
        actual_column = CDbl(ReferenceColumn) + CDbl(ColumnIndex)
        If actual_column < 1 Or G_COL_MAX < actual_column Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Column index is out of range. (" & ReferenceColumn & " + " & ColumnIndex & " = " & CStr(actual_column) & ")"
        End If

        pA1columnAddressCore = ExcelA1ColumnAddress(CLng(actual_column))
    End If
End Function

Private Function pA1RowAddressCore(ByVal RowIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceRow As Long) As String
    Dim actual_row As Double
    If IsAbsolute Then
        actual_row = RowIndex
        If actual_row < 1 Or G_ROW_MAX < actual_row Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Row index is out of range. (" & RowIndex & ")"
        End If

        pA1RowAddressCore = "$" & CStr(CLng(actual_row))
    Else
        actual_row = CDbl(ReferenceRow) + CDbl(RowIndex)
        If actual_row < 1 Or G_ROW_MAX < actual_row Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Row index is out of range. (" & ReferenceRow & " + " & RowIndex & " = " & CStr(actual_row) & ")"
        End If

        pA1RowAddressCore = CStr(CLng(actual_row))
    End If
End Function

Private Function pR1C1RowAddressCore(ByVal RowIndex As Long, ByVal IsAbsolute As Boolean) As String
    If IsAbsolute Then
        If RowIndex < 1 Or G_ROW_MAX < RowIndex Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Row index is out of range. (" & RowIndex & ")"
        End If

        pR1C1RowAddressCore = CStr(RowIndex)
    Else
        If RowIndex <> 0 Then
            pR1C1RowAddressCore = "[" & CStr(RowIndex) & "]"
        Else
            pR1C1RowAddressCore = ""
        End If
    End If
End Function

Private Function pR1C1ColumnAddressCore(ByVal ColumnIndex As Long, ByVal IsAbsolute As Boolean) As String
    If IsAbsolute Then
        If ColumnIndex < 1 Or G_COL_MAX < ColumnIndex Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="Column index is out of range. (" & ColumnIndex & ")"
        End If

        pR1C1ColumnAddressCore = CStr(ColumnIndex)
    Else
        If ColumnIndex <> 0 Then
            pR1C1ColumnAddressCore = "[" & CStr(ColumnIndex) & "]"
        Else
            pR1C1ColumnAddressCore = ""
        End If
    End If
End Function

'* Gets a string with workbook name and sheet name in Excel address format.
'*
'* @param BookName Workbook name (empty string when omitted).
'* @param SheetName Sheet name (empty string when omitted).
'* @return Excel address-format string containing the specified workbook name and sheet name.
'*
'* @details
'* Combines workbook name and sheet name to generate a string in an address format usable by Excel.
'* When BookName is specified, SheetName must also be specified.
Public Function ExcelBookAndSheetAddress(Optional ByVal BookName As String = "", Optional ByVal SheetName As String = "") As String
    Call pValidateBookAndSheetAddress(BookName, SheetName, "Function ExcelBookAndSheetAddress")

    If pIsNeedQuoteBookAndSheetAddressCore(BookName, SheetName) Then
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "'[" & pEscapeBookAndSheetAddressCore(BookName) & "]" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = "'" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        Else
            'Debug.Print "Function ExcelBookAndSheetAddress: unexpected"
        End If
    Else
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "[" & BookName & "]" & SheetName & "!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = SheetName & "!"
        Else
            ExcelBookAndSheetAddress = ""
        End If
    End If
End Function

Private Sub pValidateBookAndSheetAddress(ByVal BookName As String, ByVal SheetName As String, ByVal ErrorSource As String)
    If BookName <> "" And SheetName = "" Then
        Err.Raise vbObjectError + 1, ErrorSource, "If BookName is specified, SheetName must also be specified. (" & BookName & ")"
    End If
End Sub

Private Function pIsNeedQuoteBookAndSheetAddressCore(ByRef BookName As String, ByRef SheetName As String) As Boolean
    If BookName = "" And SheetName = "" Then
        pIsNeedQuoteBookAndSheetAddressCore = False
    ElseIf pBookNameNeedsQuote(BookName) Or pSheetNameNeedsQuote(SheetName) Then
        pIsNeedQuoteBookAndSheetAddressCore = True
    Else
        pIsNeedQuoteBookAndSheetAddressCore = False
    End If
End Function

Private Function pBookNameNeedsQuote(ByVal BookName As String) As Boolean
    If BookName = "" Then Exit Function

    pBookNameNeedsQuote = (0 < InStr(BookName, " ") Or 0 < InStr(BookName, "'") Or 0 < InStr(BookName, "!") _
            Or 0 < InStr(BookName, "[") Or 0 < InStr(BookName, "]") Or 0 < InStr(BookName, "(") Or 0 < InStr(BookName, ")"))
End Function

Private Function pSheetNameNeedsQuote(ByVal SheetName As String) As Boolean
    If SheetName = "" Then Exit Function

    If Not pIsUnquotedSheetName(SheetName) Then
        pSheetNameNeedsQuote = True
        Exit Function
    End If

    If pLooksLikeA1Reference(SheetName) Or pLooksLikeR1C1Reference(SheetName) Then
        pSheetNameNeedsQuote = True
    Else
        pSheetNameNeedsQuote = False
    End If
End Function

Private Function pIsUnquotedSheetName(ByVal SheetName As String) As Boolean
    If SheetName = "" Then Exit Function
    If Not pIsUnquotedSheetNameStartChar(Left$(SheetName, 1)) Then Exit Function

    Dim char_idx As Long
    For char_idx = 2 To Len(SheetName)
        If Not pIsUnquotedSheetNameChar(Mid$(SheetName, char_idx, 1)) Then Exit Function
    Next char_idx

    pIsUnquotedSheetName = True
End Function

Private Function pIsUnquotedSheetNameStartChar(ByVal TestChar As String) As Boolean
    pIsUnquotedSheetNameStartChar = (pIsAsciiAlphabet(TestChar) Or TestChar = "_" Or pIsNonAsciiChar(TestChar))
End Function

Private Function pIsUnquotedSheetNameChar(ByVal TestChar As String) As Boolean
    pIsUnquotedSheetNameChar = (pIsUnquotedSheetNameStartChar(TestChar) Or pIsAsciiDigit(TestChar) Or TestChar = ".")
End Function

Private Function pIsNonAsciiChar(ByVal TestChar As String) As Boolean
    Dim char_code As Long
    char_code = AscW(TestChar)
    pIsNonAsciiChar = (char_code < 0 Or 127 < char_code)
End Function

Private Function pLooksLikeA1Reference(ByVal SheetName As String) As Boolean
    Err.Clear
    On Error Resume Next
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, SheetName)
    pLooksLikeA1Reference = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function pLooksLikeR1C1Reference(ByVal SheetName As String) As Boolean
    Dim upper_name As String
    upper_name = UCase$(SheetName)
    If Left$(upper_name, 1) <> "R" Then Exit Function

    Dim c_pos As Long
    c_pos = InStr(2, upper_name, "C")
    If c_pos <= 2 Or c_pos = Len(upper_name) Then Exit Function
    If 0 < InStr(c_pos + 1, upper_name, "C") Then Exit Function

    Dim row_text As String
    Dim col_text As String
    row_text = Mid$(upper_name, 2, c_pos - 2)
    col_text = Mid$(upper_name, c_pos + 1)

    If Not pIsAsciiDigitString(row_text) Or Not pIsAsciiDigitString(col_text) Then Exit Function

    Dim row_num As Double
    Dim col_num As Double
    row_num = CDbl(row_text)
    col_num = CDbl(col_text)

    pLooksLikeR1C1Reference = (1 <= row_num And row_num <= G_ROW_MAX And 1 <= col_num And col_num <= G_COL_MAX)
End Function

Private Function pIsAsciiDigitString(ByVal TestString As String) As Boolean
    If TestString = "" Then Exit Function

    Dim char_idx As Long
    For char_idx = 1 To Len(TestString)
        If Not pIsAsciiDigit(Mid$(TestString, char_idx, 1)) Then Exit Function
    Next char_idx

    pIsAsciiDigitString = True
End Function

Private Function pEscapeBookAndSheetAddressCore(ByVal BookOrSheetName As String) As String
    pEscapeBookAndSheetAddressCore = Replace(Replace(Replace(BookOrSheetName, "'", "''"), "[", "("), "]", ")")
End Function

'* Converts a column number to an A1-format column name (for example, A, B, ... Z, AA, AB...).
'*
'* @param ColumnIndex Column number (value from 1 through G_COL_MAX).
'* @return A1-format column name corresponding to the column number.
'*
'* @details
'* Generates an Excel A1-format column name based on the specified column number.
'* Raises an error if the column number is less than 1 or greater than G_COL_MAX.
Public Function ExcelA1ColumnAddress(ByVal ColumnIndex As Long) As String
    If ColumnIndex < 1 Or G_COL_MAX < ColumnIndex Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExcelA1ColumnAddress", Description:="Column index is out of range. (" & ColumnIndex & ")"
    End If

    Dim result_value As String
    Dim temp_num As Long

    Do While ColumnIndex > 0
        temp_num = (ColumnIndex - 1) Mod 26
        result_value = Chr(temp_num + 65) & result_value
        ColumnIndex = (ColumnIndex - 1) \ 26
    Loop

    ExcelA1ColumnAddress = result_value
End Function


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel address splitting
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Extracts each piece of information from Excel address notation.
'*
'* @param FolderPath [Output] Folder path obtained as the result.
'* @param BookName [Output] Workbook name obtained as the result.
'* @param SheetName [Output] Sheet name obtained as the result.
'* @param CellAddress [Output] Cell address obtained as the result.
'* @param AddressString Excel address string to decompose.
'*
'* @details
'* Decomposes the specified Excel address string and outputs folder path, workbook name, sheet name, and cell address separately.
'* Raises an error if the format is invalid.
Public Sub SplitExcelAddress(ByRef FolderPath As String, ByRef BookName As String, ByRef SheetName As String, ByRef CellAddress As String, ByVal AddressString As String)
    FolderPath = ""
    BookName = ""
    SheetName = ""
    CellAddress = ""

    If AddressString = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim delimiter_idx As Long
    delimiter_idx = pFindExcelAddressDelimiter(AddressString)
    If delimiter_idx < 0 Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim location_part As String
    If delimiter_idx = 0 Then
        CellAddress = AddressString
        If 0 < InStr(CellAddress, "[") Or 0 < InStr(CellAddress, "]") Or 0 < InStr(CellAddress, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If
    Else
        location_part = Left(AddressString, delimiter_idx - 1)
        CellAddress = Mid(AddressString, delimiter_idx + 1)

        If location_part = "" Or CellAddress = "" Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        If StartsWith(location_part, "'") Then
            If Not EndsWith(location_part, "'") Or Len(location_part) < 2 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            location_part = Mid(location_part, 2, Len(location_part) - 2)
            location_part = Replace(location_part, "''", "'")
        ElseIf 0 < InStr(location_part, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        Dim close_book_idx As Long
        close_book_idx = InStrRev(location_part, "]")
        If 0 < close_book_idx Then
            Dim open_book_idx As Long
            open_book_idx = InStrRev(Left(location_part, close_book_idx - 1), "[")
            If open_book_idx < 1 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If

            FolderPath = Left(location_part, open_book_idx - 1)
            BookName = Mid(location_part, open_book_idx + 1, close_book_idx - open_book_idx - 1)
            SheetName = Mid(location_part, close_book_idx + 1)
            If BookName = "" Or SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        Else
            If 0 < InStr(location_part, "[") Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            SheetName = location_part
            If SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        End If
    End If

    If CellAddress = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If
End Sub

Private Function pFindExcelAddressDelimiter(ByVal AddressString As String) As Long
    Dim result_value As Long
    Dim is_in_quote As Boolean
    Dim char_idx As Long
    For char_idx = 1 To Len(AddressString)
        Dim current_char As String
        current_char = Mid$(AddressString, char_idx, 1)
        If current_char = "'" Then
            If is_in_quote And char_idx < Len(AddressString) And Mid$(AddressString, char_idx + 1, 1) = "'" Then
                char_idx = char_idx + 1
            Else
                is_in_quote = Not is_in_quote
            End If
        ElseIf current_char = "!" And Not is_in_quote Then
            If result_value <> 0 Then
                pFindExcelAddressDelimiter = -1
                Exit Function
            End If
            result_value = char_idx
        End If
    Next

    If is_in_quote Then
        pFindExcelAddressDelimiter = -1
    Else
        pFindExcelAddressDelimiter = result_value
    End If
End Function

Private Sub pRaiseInvalidExcelAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitExcelAddress", Description:="The Excel address string is not in a valid format. (" & AddressString & ")"
End Sub

'* Parses an A1-format single rectangular range address into start and finish indexes.
'*
'* @param StartRow [Output] Start row number. G_OMIT_CELL_INDEX for column ranges.
'* @param StartColumn [Output] Start column number. G_OMIT_CELL_INDEX for row ranges.
'* @param FinishRow [Output] Finish row number. G_OMIT_CELL_INDEX for column ranges.
'* @param FinishColumn [Output] Finish column number. G_OMIT_CELL_INDEX for row ranges.
'* @param AddressString A1-format address to parse. Does not include workbook name or sheet name.
'*
'* @details
'* Handles A1, A1:B2, 1:3, A:C, and $A$1:$B$2.
'* Multiple ranges, R1C1 format, addresses with workbook name or sheet name, and incomplete addresses are errors.
Public Sub SplitA1RangeAddress( _
        ByRef StartRow As Long, _
        ByRef StartColumn As Long, _
        ByRef FinishRow As Long, _
        ByRef FinishColumn As Long, _
        ByVal AddressString As String)

    StartRow = G_OMIT_CELL_INDEX
    StartColumn = G_OMIT_CELL_INDEX
    FinishRow = G_OMIT_CELL_INDEX
    FinishColumn = G_OMIT_CELL_INDEX

    Dim normalized_address As String
    normalized_address = Trim(AddressString)
    If normalized_address = "" Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If 0 < InStr(normalized_address, ",") Or 0 < InStr(normalized_address, "!") _
            Or 0 < InStr(normalized_address, "[") Or 0 < InStr(normalized_address, "]") _
            Or 0 < InStr(normalized_address, "'") Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    Dim address_parts() As String
    address_parts = Split(normalized_address, ":")

    Dim start_type As Long
    Dim finish_type As Long
    If UBound(address_parts) = 0 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        If start_type <> C_A1_TOKEN_CELL Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If
        FinishRow = StartRow
        FinishColumn = StartColumn
    ElseIf UBound(address_parts) = 1 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        Call pSplitA1AddressToken(FinishRow, FinishColumn, finish_type, address_parts(1), AddressString)
        If start_type <> finish_type Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If

        If start_type = C_A1_TOKEN_ROW Then
            StartColumn = G_OMIT_CELL_INDEX
            FinishColumn = G_OMIT_CELL_INDEX
        ElseIf start_type = C_A1_TOKEN_COLUMN Then
            StartRow = G_OMIT_CELL_INDEX
            FinishRow = G_OMIT_CELL_INDEX
        End If
    Else
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If StartRow <> G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX And FinishRow < StartRow Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
    If StartColumn <> G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX And FinishColumn < StartColumn Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
End Sub

Private Sub pSplitA1AddressToken( _
        ByRef RowIndex As Long, _
        ByRef ColumnIndex As Long, _
        ByRef TokenType As Long, _
        ByVal AddressToken As String, _
        ByVal OriginalAddressString As String)

    RowIndex = G_OMIT_CELL_INDEX
    ColumnIndex = G_OMIT_CELL_INDEX
    TokenType = 0

    Dim normalized_token As String
    normalized_token = UCase(AddressToken)
    If normalized_token = "" Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim char_idx As Long
    char_idx = 1

    Dim has_leading_dollar As Boolean
    If Mid$(normalized_token, char_idx, 1) = "$" Then
        has_leading_dollar = True
        char_idx = char_idx + 1
        If Len(normalized_token) < char_idx Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    End If

    Dim col_text As String
    Do While char_idx <= Len(normalized_token)
        Dim char_code As Long
        char_code = Asc(Mid$(normalized_token, char_idx, 1))
        If char_code < Asc("A") Or Asc("Z") < char_code Then Exit Do
        col_text = col_text & Chr(char_code)
        char_idx = char_idx + 1
    Loop

    If col_text = "" And has_leading_dollar Then
        If char_idx <= Len(normalized_token) And Mid$(normalized_token, char_idx, 1) = "$" Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    ElseIf col_text <> "" Then
        If char_idx <= Len(normalized_token) And Mid$(normalized_token, char_idx, 1) = "$" Then
            char_idx = char_idx + 1
            If Len(normalized_token) < char_idx Then
                Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
            End If
        End If
    End If

    Dim row_text As String
    Do While char_idx <= Len(normalized_token)
        char_code = Asc(Mid$(normalized_token, char_idx, 1))
        If char_code < Asc("0") Or Asc("9") < char_code Then Exit Do
        row_text = row_text & Chr(char_code)
        char_idx = char_idx + 1
    Loop

    If char_idx <= Len(normalized_token) Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    If col_text <> "" Then
        ColumnIndex = pA1ColumnIndex(col_text, OriginalAddressString)
    End If
    If row_text <> "" Then
        RowIndex = pA1RowIndex(row_text, OriginalAddressString)
    End If

    If col_text <> "" And row_text <> "" Then
        TokenType = C_A1_TOKEN_CELL
    ElseIf row_text <> "" Then
        TokenType = C_A1_TOKEN_ROW
    ElseIf col_text <> "" Then
        TokenType = C_A1_TOKEN_COLUMN
    Else
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
End Sub

Private Function pA1ColumnIndex(ByVal ColumnAddress As String, ByVal OriginalAddressString As String) As Long
    Dim result_value As Long
    Dim char_idx As Long
    For char_idx = 1 To Len(ColumnAddress)
        Dim char_code As Long
        char_code = Asc(Mid(ColumnAddress, char_idx, 1))
        result_value = result_value * 26 + char_code - Asc("A") + 1
        If G_COL_MAX < result_value Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    Next

    If result_value < 1 Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1ColumnIndex = result_value
End Function

Private Function pA1RowIndex(ByVal RowAddress As String, ByVal OriginalAddressString As String) As Long
    If Len(CStr(G_ROW_MAX)) < Len(RowAddress) Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim result_value As Long
    result_value = CLng(RowAddress)
    If result_value < 1 Or G_ROW_MAX < result_value Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1RowIndex = result_value
End Function

Private Sub pRaiseInvalidA1RangeAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitA1RangeAddress", Description:="The value is not a single-rectangle A1 range address. (" & AddressString & ")"
End Sub


' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Range value conversion
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Returns an ObjectList storing Range values.
'*
'* @param TargetRange Target range.
'* @param IgnoreEmpty True to ignore empty cells (default is False).
'* @param GetText True to get the Text property (default is False).
'* @return ObjectList storing cells in the target range converted to strings.
'*
'* @details
'* Converts each cell in the specified range to a string as Value or Text, stores them in an ObjectList, and returns it.
'* When IgnoreEmpty is True, empty cells are not included in the list.
Public Function ConvertRangeToStringList(ByVal TargetRange As WorksheetRangeBounds, Optional ByVal IgnoreEmpty As Boolean = False, Optional ByVal GetText As Boolean = False) As ObjectList
    Dim result_value As ObjectList
    Set result_value = New ObjectList

    Dim enum_obj As IEnumerator
    Set enum_obj = TargetRange.GetEnumerator( _
            EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, _
            Descending:=False)
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current

        Dim cell_text As String
        Call WsSrv.ReadCell(cell_item, cell_text, GetText:=GetText)

        If Not IgnoreEmpty Or cell_text <> "" Then
            Call result_value.Add(cell_text)
        End If
    Loop

    Set ConvertRangeToStringList = result_value
End Function

