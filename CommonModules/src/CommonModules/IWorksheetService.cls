VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorksheetService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for a service that manages worksheet access.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IWorksheetService", Description:="This class is intended for interfaces."
End Sub

'* Searches within a range.
'*
'* @param What String to search for. Empty strings are allowed and treated the same as an Excel empty-string search.
'* @param RangeBounds Range to search.
'* @param LookInValues [Optional] Whether to search values. The default is True, which searches values. When False, searches formulas.
'* @param LookAtWhole [Optional] Whether to search for cells whose contents match exactly. The default is True, which searches exact matches. When False, searches partial matches.
'* @param MatchCase [Optional] Whether to distinguish uppercase and lowercase. The default is True. When True, searches case-sensitively. When False, searches case-insensitively.
'* @param MatchByte [Optional] Whether to distinguish full-width and half-width characters. The default is True. When True, distinguishes them. When False, does not distinguish them.
'* @return ObjectList containing WorksheetRangeBounds for found cells.
'*
'* @details
'* Searches within a range. When there are no search results or the range is empty, returns an ObjectList with Count = 0.
'* Each returned element is a WorksheetRangeBounds.
'* After the call, resets the search state shared by the Excel UI and subsequent Range.Find calls to the standard value-search and partial-match state.
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList
End Function

'* Sorts within a range.
'*
'* @param RangeBounds Range to sort.
'* @param SortKeyAndOrder [ParamArray] Pairs of relative column number to use as a sort key and sort order. Column numbers start at 1 from the left edge of RangeBounds and must be from 1 through RangeBounds.ColumnCount. Sort order is xlAscending or xlDescending.
'*
'* @details
'* Sorts within a range. SortKeyAndOrder is specified as pairs of column number and sort order.
Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
End Sub

'* Activates a range.
'*
'* @details
'* Activates a range.
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* Writes a string to a cell.
'*
'* @param RangeBounds Cell range to write to. Ranges that are not a single cell are errors.
'* @param Expression String to write.
'* @param NumberFormat Display format. (Default is no setting change.)
'* @param AsFormula True to write as a formula (default is False).
'* @param ClearWhenEmpty Whether to clear the cell when Expression is an empty string (default is True).
'* @param IgnoreEmpty True to do nothing when Expression is an empty string (default is False).
'* @param TypeConvert True to convert to an appropriate type, such as number or date, before writing (default is True).
'*
'* @details
'* Sets a string or formula in the specified cell.
'* When TypeConvert is True, numeric strings are treated as integers if they can be represented as integers, and as decimals otherwise.
'* When Expression is an empty string, behavior changes based on `IgnoreEmpty` and `ClearWhenEmpty`.
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
End Sub

'* Writes values to a range.
'*
'* @param RangeBounds Destination range information. Empty ranges and single cells are treated as the top-left anchor and expanded to the same size as ValuesArray. Multi-cell ranges must have the same row and column counts as ValuesArray.
'* @param ValuesArray Two-dimensional array to write.
'*
'* @details
'* Writes a two-dimensional array to the specified range. The destination range accepts only an empty range, a single cell, or a range with the same shape as ValuesArray.
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
End Sub

'* Writes an array (CSE) formula to a range.
'*
'* @param RangeBounds Destination range information.
'* @param FormulaArray Array (CSE) formula.
'*
'* @details
'* Writes an array (CSE) formula to a range.
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
End Sub

'* Reads a value from a cell.
'*
'* @param RangeBounds Cell range to read. Empty ranges return an empty string, and ranges that are not a single cell are errors.
'* @param Expression [Output][Optional] String that was read. Same as the return value. Empty string for empty ranges.
'* @param NumberFormat [Output][Optional] Display format (NumberFormatLocal). Empty string for empty ranges.
'* @param GetText [Optional] Whether to read Text instead of Value.
'* @return String that was read. Same as Expression.
'*
'* @details
'* Reads a value from a cell. For an empty range, sets Expression and NumberFormat to empty strings and returns an empty string.
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
End Function

'* Reads values from a range.
'*
'* @param RangeBounds Range information to read.
'* @return Array that was read.
'*
'* @details
'* Reads a two-dimensional array from the specified range.
'* Internally, this reads from Range.Value with =, so detailed behavior follows Excel.
'* Unlike Excel behavior, even when a cell range is specified, this returns a two-dimensional array. This enables unified processing.
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
End Function

'* Evaluates a formula.
'*
'* @param ContextRangeBounds Range information indicating the sheet on which to evaluate the formula.
'* @param Formula Formula to evaluate.
'* @return Evaluation result.
'*
'* @details
'* Evaluates Formula on the sheet of ContextRangeBounds and returns the result.
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
End Function

'* Gets the result of XLOOKUP.
'*
'* @param LookupValue Lookup value.
'* @param LookupRangeBounds Lookup range.
'* @param ReturnRangeBounds Return range.
'* @param IfNotFound [Optional] Value to return when not found.
'* @param MatchMode [Optional] Match mode.
'* @param SearchMode [Optional] Search mode.
'* @return Search result.
'*
'* @details
'* Evaluates the Excel XLOOKUP function against the specified ranges and returns the result.
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
End Function

'* Copies a cell.
'*
'* @param SourceRangeBounds Source cell range. Ranges that are not a single cell are errors.
'* @param DestinationRangeBounds Destination cell range. Ranges that are not a single cell are errors.
'* @param AsValue [Optional] Whether to copy the value as a string. (Default is False.)
'* @param CopyNumberFormat [Optional] Whether to copy the display format. (Default is False.)
'*
'* @details
'* Copies the value or formula of the specified single cell.
'* Formula copying also supports array formulas.
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
End Sub

'* Copies a range.
'*
'* @param SourceRangeBounds Source range.
'* @param DestinationRangeBounds Destination range. If empty or a single cell, it is treated as the top-left anchor and expanded to the same size as the source. Multi-cell ranges must have the same row and column counts as the source.
'* @param AsValue [Optional] Whether to copy values as strings. (Default is False.)
'* @param CopyNumberFormat [Optional] Whether to copy display formats. (Default is False.)
'*
'* @details
'* Copies the values or formulas of the specified range.
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
End Sub

'* Checks whether a cell is empty.
'*
'* @param RangeBounds Target cell or range.
'* @param IgnoreEmptyString [Optional] Whether to ignore blank strings. Default value is False.
'* @return Whether the cell is empty.
'*
'* @details
'* Checks whether a cell is empty. When IgnoreEmptyString is True, a cell with a formula is judged empty if the calculated result is an empty string.
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
End Function

'* Checks whether a cell contains a formula.
'*
'* @param RangeBounds Target cell.
'* @return Whether a formula is entered.
'*
'* @details
'* Returns whether the specified cell contains a formula. Implementations should raise an error when the range is not a cell.
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
End Function

'* Gets the range that is actually used.
'*
'* @param RangeBounds Target range from which to get the used range.
'* @param IgnoreEmpty [Optional] Whether to ignore empty strings. Default value is True.
'* @param IsUsedStartRow [Optional] Whether to treat the start row of RangeBounds as used. Default value is True.
'* @param IsUsedStartColumn [Optional] Whether to treat the start column of RangeBounds as used. Default value is True.
'* @param GetRawRange [Optional] Whether to return a range based on the overlap between RangeBounds and UsedRange, without trimming by values.
'* @return Range that is actually used.
'*
'* @details
'* Returns the range that is actually used (contains values) within the specified range.
'* Cells that have only formatting and no value are treated as unused.
'* When IgnoreEmpty is True, cells whose formulas return empty strings are ignored even if formulas exist.
'* When GetRawRange is True, returns a range based on the overlap between Excel UsedRange and RangeBounds, without trimming empty rows or columns by value.
'* When IsUsedStartRow / IsUsedStartColumn is True, the returned start row / start column retains the start row / start column of RangeBounds.
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
End Function

'* Clears a specified range.
'*
'* @param RangeBounds Range information to clear.
'* @param ClearAll [Optional] Clears everything. When this option is True, False settings for other options are ignored.
'* @param ClearContents [Optional] Clears cell contents. Default is True.
'* @param ClearNumberFormats [Optional] Clears cell display formats.
'* @param ClearColors [Optional] Clears cell font colors and background colors.
'* @param ClearComments [Optional] Clears comments.
'* @param ClearHyperlinks [Optional] Clears hyperlinks.
'*
'* @details
'* Clears a specified range.
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
End Sub

'* Removes duplicate rows within a range.
'*
'* @param RangeBounds Range from which to remove duplicates.
'* @param DuplicateColumns Column number or array of column numbers to use for duplicate checks. Specify relative column numbers where the left edge of RangeBounds is 1.
'*
'* @details
'* Removes duplicate rows within the specified range using the same criteria as Excel RemoveDuplicates. DuplicateColumns must specify one or more values from 1 through RangeBounds.ColumnCount.
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
End Sub

'* Inserts rows.
'*
'* @param RangeBounds Row range to insert. Only whole-row ranges are accepted.
'*
'* @details
'* Inserts sheet rows from RangeBounds.Row through RangeBounds.FinishRow. Empty ranges and cell ranges are errors.
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* Deletes rows.
'*
'* @param RangeBounds Row range to delete. Only whole-row ranges are accepted.
'*
'* @details
'* Deletes sheet rows from RangeBounds.Row through RangeBounds.FinishRow. Empty ranges and cell ranges are errors.
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* Shows all data.
'*
'* @param RangeBounds Information for the sheet to clear. Column and row range information is ignored.
'*
'* @details
'* Performs Sheet.ShowAllData, Rows.Hidden = False, and Columns.Hidden = False.
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
End Sub

'* Sets the outline level to display.
'*
'* @param RangeBounds Information for the sheet whose outline level is set. Column and row range information is ignored.
'* @param RowLevels [Optional] Row outline level. 0 means not specified; 1 through 8 are display levels.
'* @param ColumnLevels [Optional] Column outline level. 0 means not specified; 1 through 8 are display levels.
'*
'* @details
'* Sets the outline level to display for an existing outline. Does nothing when both RowLevels and ColumnLevels are 0.
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
End Sub

'* Sets the sheet tab color.
'*
'* @param RangeBounds Information for the sheet whose tab color is set. Column and row range information is ignored.
'* @param TabColorIndex [Optional] Tab color index. Default is no color.
'* @param TabColor [Optional] Tab color value. Set the value with the RGB(,,) function.
'*
'* @details
'* Sets the sheet tab color. If both an index and a value are specified, the value takes precedence.
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
End Sub

'* Sets range colors.
'*
'* @param RangeBounds Information for the sheet whose colors are set.
'* @param FontColorIndex [Optional] Font color index. Default is no color.
'* @param FontColor [Optional] Font color value. Set the value with the RGB(,,) function.
'* @param InteriorColorIndex [Optional] Background color index. Default is no color.
'* @param InteriorColor [Optional] Background color value. Set the value with the RGB(,,) function.
'*
'* @details
'* Sets range colors. If both an index and a value are specified, the value takes precedence.
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
End Sub

'* Sets whether text is displayed wrapped.
'*
'* @param RangeBounds Target range information.
'* @param WrapText Whether to display text wrapped.
'*
'* @details
'* Sets whether text is displayed wrapped.
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
End Sub

'* Sets whether text is shrunk to fit the whole content.
'*
'* @param RangeBounds Target range information.
'* @param ShrinkToFit Whether to shrink text to fit the whole content.
'*
'* @details
'* Sets whether text is shrunk to fit the whole content.
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
End Sub

'* Sets cell alignment.
'*
'* @param RangeBounds Target range information.
'* @param HorizontalAlignment Horizontal alignment. No change when G_ALIGNMENT_NO_CHANGE.
'* @param VerticalAlignment Vertical alignment. No change when G_ALIGNMENT_NO_CHANGE.
'* @param Orientation Text orientation. No change when G_ALIGNMENT_NO_CHANGE.
'* @param IndentLevel Indent level. No change when G_ALIGNMENT_NO_CHANGE.
'*
'* @details
'* Sets cell alignment.
'* Items specified as G_ALIGNMENT_NO_CHANGE do not change existing settings.
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)
End Sub
