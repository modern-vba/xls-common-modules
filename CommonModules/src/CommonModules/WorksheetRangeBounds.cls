VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetRangeBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStringable
Implements IEquatable

' #############################################################################
'!
'! @brief
'! Class module that holds Excel range information.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pRow As Long
Private pColumn As Long
Private pFinishRow As Long
Private pFinishColumn As Long
Private pWorkbookName As String
Private pWorksheetName As String
Private pNewEnumSnapshot As Collection

'* Row number.
Public Property Get Row() As Long
    Call pCheckInit
    Row = pRow
End Property

'* Column number.
Public Property Get Column() As Long
    Call pCheckInit
    Column = pColumn
End Property

'* Last row number.
Public Property Get FinishRow() As Long
    Call pCheckInit
    FinishRow = pFinishRow
End Property

'* Last column number.
Public Property Get FinishColumn() As Long
    Call pCheckInit
    FinishColumn = pFinishColumn
End Property

'* Worksheet name.
Public Property Get WorksheetName() As String
    Call pCheckInit
    WorksheetName = pWorksheetName
End Property

'* Workbook name.
Public Property Get WorkbookName() As String
    Call pCheckInit
    WorkbookName = pWorkbookName
End Property

'* Row count.
Public Property Get RowCount() As Long
    Call pCheckInit
    Dim result_value As Long
    result_value = pFinishRow - pRow + 1
    If 0 < result_value Then
        RowCount = result_value
    Else
        RowCount = 0
    End If
End Property

'* Column count.
Public Property Get ColumnCount() As Long
    Call pCheckInit
    Dim result_value As Long
    result_value = pFinishColumn - pColumn + 1
    If 0 < result_value Then
        ColumnCount = result_value
    Else
        ColumnCount = 0
    End If
End Property

'* Cell count.
Public Property Get Count() As Long
    Call pCheckInit

    Count = RowCount * ColumnCount

'    Dim row_count As Long, col_count As Long
'    row_count = RowCount
'    col_count = ColumnCount
'
'    If row_count = 0 And col_count = 0 Then
'        Count = 0
'    Else
'        Count = MaxLng(RowCount, 1) * MaxLng(ColumnCount, 1)
'    End If
End Property

'* Whether the range is empty.
Public Property Get IsEmpty() As Boolean
    Call pCheckInit
    If pFinishRow < pRow Or pFinishColumn < pColumn Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Property

'* Whether the range is the entire sheet.
Public Property Get IsEntireSheet() As Boolean
    Call pCheckInit
    If IsEntireRow And IsEntireColumn Then
        IsEntireSheet = True
    Else
        IsEntireSheet = False
    End If
End Property

'* Whether the range is an entire row (range such as 1:2).
Public Property Get IsEntireRow() As Boolean
    Call pCheckInit
    If pColumn = 1 And pFinishColumn = G_COL_MAX Then
        IsEntireRow = True
    Else
        IsEntireRow = False
    End If
End Property

'* Whether the range is an entire column (range such as A:B).
Public Property Get IsEntireColumn() As Boolean
    Call pCheckInit
    If pRow = 1 And pFinishRow = G_ROW_MAX Then
        IsEntireColumn = True
    Else
        IsEntireColumn = False
    End If
End Property

'* Whether the range is a single cell. Equivalent to a non-empty range that is not an Area.
Public Property Get IsCell() As Boolean
    Call pCheckInit
    If Not IsEmpty And pRow = pFinishRow And pColumn = pFinishColumn Then
        IsCell = True
    Else
        IsCell = False
    End If
End Property

'* Whether the range is an Area. A contiguous multi-cell range excluding a single cell is treated as an Area.
Public Property Get IsArea() As Boolean
    Call pCheckInit
    If Not IsEmpty And Not IsCell() Then
        IsArea = True
    Else
        IsArea = False
    End If
End Property

'* Whether the range has a width of one row (such as 1:1 or A1:B1). A single cell is also True.
Public Property Get IsOneRow() As Boolean
    Call pCheckInit
    If IsEmpty Then
        IsOneRow = False
        Exit Property
    End If

    If pRow = pFinishRow Then
        IsOneRow = True
    Else
        IsOneRow = False
    End If
End Property

'* Whether the range has a width of one column (such as A:A or A1:A2). A single cell is also True.
Public Property Get IsOneColumn() As Boolean
    Call pCheckInit
    If IsEmpty Then
        IsOneColumn = False
        Exit Property
    End If

    If pColumn = pFinishColumn Then
        IsOneColumn = True
    Else
        IsOneColumn = False
    End If
End Property

'* Whether the range is an Area consisting of only one row. A single cell is False.
Public Property Get IsOneRowArea() As Boolean
    Call pCheckInit
    If IsArea And IsOneRow Then
        IsOneRowArea = True
    Else
        IsOneRowArea = False
    End If
End Property

'* Whether the range is an Area consisting of only one column. A single cell is False.
Public Property Get IsOneColumnArea() As Boolean
    Call pCheckInit
    If IsArea And IsOneColumn Then
        IsOneColumnArea = True
    Else
        IsOneColumnArea = False
    End If
End Property

'* Gets the enumerator for standard For Each.
'*
'* @return Enumerator used by standard For Each.
'*
'* @details
'* Enumerates a snapshot of cell elements from the start of enumeration, in row-major and ascending order.
'* The enumeration mode for standard For Each is fixed and independent of arguments specified for GetEnumerator.
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* Indexer for this collection. Read-only. Gets a cell range by number.
'*
'* @param ItemIndex Number.
'* @param ColumnDirection [Optional] Whether to scan in the column (vertical) direction. Default is False, scanning in the row (horizontal) direction.
Public Property Get Item(ByVal ItemIndex As Long, Optional ByVal ColumnDirection As Boolean = False) As WorksheetRangeBounds
Attribute Item.VB_UserMemId = 0
    Call pCheckInit
    If ItemIndex < 0 Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds.Item", "Index is out of range."
    End If

    If ColumnDirection Then
        Set Item = pGetCellVertical(ItemIndex)
    Else
        Set Item = pGetCellHorizontal(ItemIndex)
    End If
End Property

Private Function pGetCellVertical(ByVal ItemIndex As Long) As WorksheetRangeBounds
    If IsEmpty Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds.pGetCellVertical", "The range is empty."
    End If

    Dim row_count As Long, row_idx As Long, col_idx As Long
    row_count = RowCount
    row_idx = ItemIndex Mod row_count
    col_idx = (ItemIndex - row_idx) / row_count

    If col_idx < ColumnCount Then
        Set pGetCellVertical = GetCell(row_idx + 1, col_idx + 1)
    Else
        Err.Raise vbObjectError, "Class WorksheetRangeBounds.pGetCellVertical", "Index is out of range."
    End If
End Function

Private Function pGetCellHorizontal(ByVal ItemIndex As Long) As WorksheetRangeBounds
    If IsEmpty Then
        Err.Raise vbObjectError, "Class WorksheetRangeBounds.pGetCellHorizontal", "The range is empty."
    End If

    Dim col_count As Long, row_idx As Long, col_idx As Long
    col_count = ColumnCount
    col_idx = ItemIndex Mod col_count
    row_idx = (ItemIndex - col_idx) / col_count

    If row_idx < RowCount Then
        Set pGetCellHorizontal = GetCell(row_idx + 1, col_idx + 1)
    Else
        Err.Raise vbObjectError, "Class WorksheetRangeBounds.pGetCellHorizontal", "Index is out of range."
    End If
End Function

'* Sets each value.
'*
'* @param RowIndex Row number.
'* @param ColumnIndex Column number.
'* @param FinishRowIndex Last row number.
'* @param FinishColumnIndex Last column number.
'* @param SheetName Worksheet name.
'* @param BookName Workbook name.
'*
'* @details
'* Sets each value.
'*
'* Usage example:
'* @code
'* ' One cell (B5).
'* Dim cell_bounds As WorksheetRangeBounds
'* Set cell_bounds = New WorksheetRangeBounds
'* Call cell_bounds.Initialize(Row:=5, Column:=2)
'*
'* ' Entire row (5:5).
'* Dim row_bounds As WorksheetRangeBounds
'* Set row_bounds = New WorksheetRangeBounds
'* Call row_bounds.Initialize(Row:=5)
'*
'* ' Two entire columns (B:C).
'* Dim two_columns_bounds As WorksheetRangeBounds
'* Set two_columns_bounds = New WorksheetRangeBounds
'* Call two_columns_bounds.Initialize(Column:=2, FinishColumn:=3)
'*
'* ' Entire table (B2:E20).
'* Dim table_bounds As WorksheetRangeBounds
'* Set table_bounds = New WorksheetRangeBounds
'* Call table_bounds.Initialize(Row:=2, Column:=2, FinishRow:=20, FinishColumn:=5)
'* @endcode
Public Sub Initialize( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Sheet As String = "Sheet1", _
        Optional ByVal Book As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Initialize", "The object is already initialized."

'    If 0 < FinishRow <> G_OMIT_CELL_INDEX And FinishRow < Row Then
'        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "The start row and finish row are reversed. (" & Row & ", " & FinishRow & ")"
'        Exit Sub
'    End If
'
'    If 0 < FinishColumn And FinishColumn < Column Then
'        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds", "The start column and finish column are reversed. (" & Column & ", " & FinishColumn & ")"
'        Exit Sub
'    End If

    If Sheet = "" Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Initialize", "Worksheet name cannot be omitted."
        Exit Sub
    End If

    If Book = "" Then
        Book = pGetDefaultWorkbookName()
    End If

    Call pWellFormBounds(Row, Column, FinishRow, FinishColumn)

    pRow = Row
    pColumn = Column
    pFinishRow = FinishRow
    pFinishColumn = FinishColumn
    pWorksheetName = Sheet
    pWorkbookName = Book
    Set pNewEnumSnapshot = Nothing
    pIsInitialized = True
End Sub

'* Initializes from an Excel address string.
'*
'* @param AddressString Excel address string to use for initialization.
'*
'* @details
'* Handles only A1-format single rectangular ranges. Multiple ranges cannot be represented by WorksheetRangeBounds and are errors.
Public Sub InitializeFromAddress(ByVal AddressString As String)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.InitializeFromAddress", "The object is already initialized."

    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, AddressString)

    If folder_path <> "" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WorksheetRangeBounds.InitializeFromAddress", Description:="Addresses with folder paths are not supported. (" & AddressString & ")"
    End If

    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, cell_address)

    If sheet_name = "" Then
        sheet_name = "Sheet1"
    End If

    Call Initialize( _
            Row:=start_row, _
            Column:=start_col, _
            FinishRow:=finish_row, _
            FinishColumn:=finish_col, _
            Sheet:=sheet_name, _
            Book:=book_name)
End Sub

Private Function pGetDefaultWorkbookName() As String
    Call pRequireWorkbookService("pGetDefaultWorkbookName")
    pGetDefaultWorkbookName = WbSrv.GetThisWorkbookName()
End Function

Private Sub pWellFormBounds(ByRef Row As Long, ByRef Column As Long, ByRef FinishRow As Long, ByRef FinishColumn As Long)
    If Row = G_OMIT_CELL_INDEX Then
        Row = 1
        If FinishRow = G_OMIT_CELL_INDEX Then
            FinishRow = G_ROW_MAX
        End If
    ElseIf Row < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Start row must be 1 or greater. (" & Row & ")"
    ElseIf G_ROW_MAX < Row Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Start row exceeds " & G_ROW_MAX & ". (" & Row & ")"
    End If

    If Column = G_OMIT_CELL_INDEX Then
        Column = 1
        If FinishColumn = G_OMIT_CELL_INDEX Then
            FinishColumn = G_COL_MAX
        End If
    ElseIf Column < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Start column must be 1 or greater. (" & Column & ")"
    ElseIf G_COL_MAX < Column Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Start column exceeds " & G_COL_MAX & ". (" & Column & ")"
    End If

    If FinishRow = G_OMIT_CELL_INDEX Then
        FinishRow = Row
    ElseIf FinishRow < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Finish row cannot be negative. (" & FinishRow & ")"
    ElseIf FinishRow < Row Then
        FinishRow = 0
    End If

    If FinishColumn = G_OMIT_CELL_INDEX Then
        FinishColumn = Column
    ElseIf FinishColumn < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pWellFormBounds", "Finish column cannot be negative. (" & FinishColumn & ")"
    ElseIf FinishColumn < Column Then
        FinishColumn = 0
    End If

    If G_ROW_MAX < FinishRow Then FinishRow = G_ROW_MAX
    If G_COL_MAX < FinishColumn Then FinishColumn = G_COL_MAX

'    If FinishRow = 0 Or FinishColumn = 0 Then
'        FinishRow = 0
'        FinishColumn = 0
'    End If
End Sub

'* Gets a one-row range by specifying a row index from the top row.
'*
'* @param RowIndex Row index. Positive numbers are 1-origin from the top; negative numbers are 1-origin from the bottom. 0 is an out-of-range error.
'* @return Specified range (one row).
'*
'* @details
'* Gets a one-row range by specifying a row index from the top for positive numbers or from the bottom for negative numbers.
Public Function GetRow(ByVal RowIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_row As Long: target_row = pGetTargetRow(RowIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=target_row, Column:=pColumn, FinishRow:=target_row, FinishColumn:=pFinishColumn, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetRow = result_value
End Function

'* Gets a one-column range by specifying a column index from the leftmost column.
'*
'* @param ColumnIndex Column index. Positive numbers are 1-origin from the left; negative numbers are 1-origin from the right. 0 is an out-of-range error.
'* @return Specified range (one column).
'*
'* @details
'* Gets a one-column range by specifying a column index from the leftmost side for positive numbers or from the rightmost side for negative numbers.
Public Function GetColumn(ByVal ColumnIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_col As Long: target_col = pGetTargetColumn(ColumnIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=pRow, Column:=target_col, FinishRow:=pFinishRow, FinishColumn:=target_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetColumn = result_value
End Function

'* Gets a one-cell range by specifying row and column indexes from the upper-left cell.
'*
'* @param RowIndex Row index. Positive numbers are 1-origin from the top; negative numbers are 1-origin from the bottom. 0 is an out-of-range error.
'* @param ColumnIndex Column index. Positive numbers are 1-origin from the left; negative numbers are 1-origin from the right. 0 is an out-of-range error.
'* @return Specified range (one cell).
'*
'* @details
'* Gets a one-cell range by specifying row and column indexes based on the upper-left for positive numbers and the lower-right for negative numbers.
Public Function GetCell(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As WorksheetRangeBounds
    Call pCheckInit
    Dim target_row As Long: target_row = pGetTargetRow(RowIndex)
    Dim target_col As Long: target_col = pGetTargetColumn(ColumnIndex)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=target_row, Column:=target_col, FinishRow:=target_row, FinishColumn:=target_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set GetCell = result_value
End Function

'* Gets a list created by splitting this range into one-row units.
'*
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @return ObjectList storing one-row WorksheetRangeBounds ranges.
'*
'* @details
'* Splits this range into one-row WorksheetRangeBounds units and returns them as a typed ObjectList.
Public Function GetRows(Optional ByVal Descending As Boolean = False) As ObjectList
    Set GetRows = pGetBoundsList(EnumerationMode:=G_RANGE_ENUM_MODE_ROWS, Descending:=Descending)
End Function

'* Gets a list created by splitting this range into one-column units.
'*
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @return ObjectList storing one-column WorksheetRangeBounds ranges.
'*
'* @details
'* Splits this range into one-column WorksheetRangeBounds units and returns them as a typed ObjectList.
Public Function GetColumns(Optional ByVal Descending As Boolean = False) As ObjectList
    Set GetColumns = pGetBoundsList(EnumerationMode:=G_RANGE_ENUM_MODE_COLUMNS, Descending:=Descending)
End Function

'* Gets a new range transformed from this range.
'*
'* @param Row Row number.
'* @param Column Column number.
'* @param FinishRow Last row number.
'* @param FinishColumn Last column number.
'* @return New transformed object.
'*
'* @details
'* Gets a new range transformed from this range.
Public Function TransformAbsolute( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX) As WorksheetRangeBounds
    Call pCheckInit

    Dim s_row As Long, f_row As Long
    Dim s_col As Long, f_col As Long

    If Row = G_OMIT_CELL_INDEX Then
        s_row = pRow
    ElseIf Row < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Start row must be 1 or greater. (" & Row & ")"
    ElseIf G_ROW_MAX < Row Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Start row exceeds " & G_ROW_MAX & ". (" & Row & ")"
    Else
        s_row = Row
    End If

    If FinishRow = G_OMIT_CELL_INDEX Then
        f_row = pFinishRow
    ElseIf FinishRow < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Finish row cannot be negative. (" & FinishRow & ")"
    Else
        f_row = FinishRow
    End If

    If Column = G_OMIT_CELL_INDEX Then
        s_col = pColumn
    ElseIf Column < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Start column must be 1 or greater. (" & Column & ")"
    ElseIf G_COL_MAX < Column Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Start column exceeds " & G_COL_MAX & ". (" & Column & ")"
    Else
        s_col = Column
    End If

    If FinishColumn = G_OMIT_CELL_INDEX Then
        f_col = pFinishColumn
    ElseIf FinishColumn < 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.TransformAbsolute", "Finish column cannot be negative. (" & FinishColumn & ")"
    Else
        f_col = FinishColumn
    End If

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=s_row, Column:=s_col, FinishRow:=f_row, FinishColumn:=f_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set TransformAbsolute = result_value
End Function

'* Gets a new range transformed from this range.
'*
'* @param AddRow Number of rows to add.
'* @param AddColumn Number of columns to add.
'* @return New transformed object.
'*
'* @details
'* Gets a new range whose row count and column count are increased or decreased while preserving the start position of this range.
Public Function Transform( _
        Optional ByVal AddRow As Long = 0, _
        Optional ByVal AddColumn As Long = 0) As WorksheetRangeBounds
    Call pCheckInit

    Dim new_finish_row As Long
    Dim new_finish_col As Long
    If AddRow = 0 Then
        new_finish_row = pFinishRow
    Else
        Dim new_finish_row_value As Double
        If pFinishRow = 0 Then
            new_finish_row_value = CDbl(pRow) + CDbl(AddRow) - 1#
            If new_finish_row_value < 0# Then
                new_finish_row = 0
            ElseIf G_ROW_MAX < new_finish_row_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Transform", "The transformed row exceeds " & G_ROW_MAX & ". (row: " & pRow & ", add: " & AddRow & ")"
            Else
                new_finish_row = CLng(new_finish_row_value)
            End If
        Else
            new_finish_row_value = CDbl(pFinishRow) + CDbl(AddRow)
            If new_finish_row_value < 0# Then
                new_finish_row = 0
            ElseIf G_ROW_MAX < new_finish_row_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Transform", "The transformed row exceeds " & G_ROW_MAX & ". (row: " & pFinishRow & ", add: " & AddRow & ")"
            Else
                new_finish_row = CLng(new_finish_row_value)
            End If
        End If
    End If

    If AddColumn = 0 Then
        new_finish_col = pFinishColumn
    Else
        Dim new_finish_col_value As Double
        If pFinishColumn = 0 Then
            new_finish_col_value = CDbl(pColumn) + CDbl(AddColumn) - 1#
            If new_finish_col_value < 0# Then
                new_finish_col = 0
            ElseIf G_COL_MAX < new_finish_col_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Transform", "The transformed column exceeds " & G_COL_MAX & ". (col: " & pColumn & ", add: " & AddColumn & ")"
            Else
                new_finish_col = CLng(new_finish_col_value)
            End If
        Else
            new_finish_col_value = CDbl(pFinishColumn) + CDbl(AddColumn)
            If new_finish_col_value < 0# Then
                new_finish_col = 0
            ElseIf G_COL_MAX < new_finish_col_value Then
                Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Transform", "The transformed column exceeds " & G_COL_MAX & ". (col: " & pFinishColumn & ", add: " & AddColumn & ")"
            Else
                new_finish_col = CLng(new_finish_col_value)
            End If
        End If
    End If

    Dim result_value As WorksheetRangeBounds
    Set result_value = New_RangeBounds(Row:=pRow, Column:=pColumn, FinishRow:=new_finish_row, FinishColumn:=new_finish_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
    Set Transform = result_value
End Function

'* Gets a new range shifted from this range.
'*
'* @param ShiftRow Row shift.
'* @param ShiftColumn Column shift.
'* @return New shifted object.
'*
'* @details
'* Creates a duplicate of this object.
'* Gets a new range shifted from this range.
Public Function Shift(Optional ByVal Row As Long = 0, Optional ByVal Column As Long = 0) As WorksheetRangeBounds
    Call pCheckInit
    Dim new_s_row As Long, new_s_col As Long
    Dim new_f_row As Long, new_f_col As Long
    Dim new_s_row_value As Double, new_s_col_value As Double
    Dim new_f_row_value As Double, new_f_col_value As Double

    new_s_row_value = CDbl(pRow) + CDbl(Row)
    new_s_col_value = CDbl(pColumn) + CDbl(Column)
    If pFinishRow = 0 Then
        new_f_row_value = 0#
    Else
        new_f_row_value = CDbl(pFinishRow) + CDbl(Row)
    End If
    If pFinishColumn = 0 Then
        new_f_col_value = 0#
    Else
        new_f_col_value = CDbl(pFinishColumn) + CDbl(Column)
    End If

    If new_s_row_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The row after shifting would be less than 1. (row: " & pRow & ", shift: " & Row & ")"
    If new_s_col_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The column after shifting would be less than 1. (col: " & pColumn & ", shift: " & Column & ")"
    If new_f_row_value <> 0# And new_f_row_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The row after shifting would be less than 1. (row: " & pFinishRow & ", shift: " & Row & ")"
    If new_f_col_value <> 0# And new_f_col_value < 1# Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The column after shifting would be less than 1. (col: " & pFinishColumn & ", shift: " & Column & ")"
    If G_ROW_MAX < new_s_row_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The row after shifting exceeds " & G_ROW_MAX & ". (row: " & pRow & ", shift: " & Row & ")"
    If G_ROW_MAX < new_f_row_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The row after shifting exceeds " & G_ROW_MAX & ". (row: " & pFinishRow & ", shift: " & Row & ")"
    If G_COL_MAX < new_s_col_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The column after shifting exceeds " & G_COL_MAX & ". (col: " & pColumn & ", shift: " & Column & ")"
    If G_COL_MAX < new_f_col_value Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Shift", "The column after shifting exceeds " & G_COL_MAX & ". (col: " & pFinishColumn & ", shift: " & Column & ")"

    new_s_row = CLng(new_s_row_value)
    new_s_col = CLng(new_s_col_value)
    new_f_row = CLng(new_f_row_value)
    new_f_col = CLng(new_f_col_value)

    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds

    Call result_value.Initialize( _
            Row:=new_s_row, _
            Column:=new_s_col, _
            FinishRow:=new_f_row, _
            FinishColumn:=new_f_col, _
            Sheet:=pWorksheetName, Book:=pWorkbookName)

    Set Shift = result_value
End Function

'* Returns the range where two ranges overlap.
'*
'* @param OtherRangeBounds Other range.
'* @return Overlapping range.
'*
'* @details
'* Returns a new object indicating the range where this range and another range overlap.
'* Raises an error if the ranges are not in the same worksheet.
Public Function Intersect(ByVal OtherRangeBounds As WorksheetRangeBounds) As WorksheetRangeBounds
    Call pCheckInit

    If StrComp(pWorksheetName, OtherRangeBounds.WorksheetName, vbTextCompare) <> 0 Or StrComp(pWorkbookName, OtherRangeBounds.WorkbookName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.Intersect", "Worksheets are different. (this: " & pWorksheetName & "@" & pWorkbookName & " / other: " & OtherRangeBounds.WorksheetName & "@" & OtherRangeBounds.WorkbookName & ")"
    End If

    Dim s_row As Long, f_row As Long
    Dim s_col As Long, f_col As Long
    s_row = MaxLng(pRow, OtherRangeBounds.Row)
    f_row = MinLng(pFinishRow, OtherRangeBounds.FinishRow)
    s_col = MaxLng(pColumn, OtherRangeBounds.Column)
    f_col = MinLng(pFinishColumn, OtherRangeBounds.FinishColumn)

'    If s_row <= f_row And s_col <= f_col Then
        Dim result_value As WorksheetRangeBounds
        Set result_value = New_RangeBounds(Row:=s_row, Column:=s_col, FinishRow:=f_row, FinishColumn:=f_col, Sheet:=pWorksheetName, Book:=pWorkbookName)
        Set Intersect = result_value
'    Else
'        Set Intersect = New_RangeBounds( _
'                Row:=MaxLng(MinLng(s_row, f_row, pFinishRow), pRow), _
'                Column:=MaxLng(MinLng(s_col, f_col, pFinishColumn), pColumn), _
'                Sheet:=pWorksheetName, Book:=pWorkbookName)
'    End If
End Function

'* Creates a duplicate of this object.
'*
'* @return New duplicated object.
'*
'* @details
'* Creates a duplicate of this object.
Public Function CopyObject() As WorksheetRangeBounds
    Call pCheckInit
    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds
    Call result_value.Initialize(pRow, pColumn, pFinishRow, pFinishColumn, pWorksheetName, pWorkbookName)
    Set CopyObject = result_value
End Function

'* Gets an address notation string.
'*
'* @param IsAbsoluteStartRow True to make the start row an absolute reference (default is False).
'* @param IsAbsoluteStartColumn True to make the start column an absolute reference (default is False).
'* @param IsAbsoluteFinishRow True to make the finish row an absolute reference (default is False).
'* @param IsAbsoluteFinishColumn True to make the finish column an absolute reference (default is False).
'* @param AddressType Address format ("A1" or "R1C1") (default is "A1").
'* @return Address string.
'*
'* @details
'* Generates from the address notation string of this range.
Public Function ToString( _
        Optional ByVal IsAbsoluteStartRow As Boolean = False, _
        Optional ByVal IsAbsoluteStartColumn As Boolean = False, _
        Optional ByVal IsAbsoluteFinishRow As Boolean = False, _
        Optional ByVal IsAbsoluteFinishColumn As Boolean = False, _
        Optional ByVal AddressType As String = "A1", _
        Optional ByVal CellOnly As Boolean = False) As String
    If Not pIsInitialized Then
        ToString = "UNINITIALIZED(" & Hex(ObjPtr(Me)) & ")"
        Exit Function
    End If
    Dim book_name As String
    Dim sheet_name As String
    If CellOnly Then
        book_name = ""
        sheet_name = ""
    Else
        book_name = WorkbookName
        sheet_name = WorksheetName
    End If

    If IsEmpty Then
        If IsEntireRow Then
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    FinishRow:=Row, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY"
        ElseIf IsEntireColumn Then
            ToString = RangeAddress( _
                    StartColumn:=Column, _
                    FinishColumn:=Column, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY"
        Else
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    StartColumn:=Column, _
                    FinishRow:=MaxLng(FinishRow, Row), _
                    FinishColumn:=MaxLng(FinishColumn, Column), _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name) & ":EMPTY(R" & RowCount & ",C" & ColumnCount & ")"
        End If
    Else
        If IsEntireRow Then
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    FinishRow:=FinishRow, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        ElseIf IsEntireColumn Then
            ToString = RangeAddress( _
                    StartColumn:=Column, _
                    FinishColumn:=FinishColumn, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        Else
            ToString = RangeAddress( _
                    StartRow:=Row, _
                    StartColumn:=Column, _
                    FinishRow:=FinishRow, _
                    FinishColumn:=FinishColumn, _
                    IsAbsoluteStartRow:=IsAbsoluteStartRow, _
                    IsAbsoluteFinishRow:=IsAbsoluteFinishRow, _
                    IsAbsoluteStartColumn:=IsAbsoluteStartColumn, _
                    IsAbsoluteFinishColumn:=IsAbsoluteFinishColumn, _
                    ReferenceRow:=0, _
                    ReferenceColumn:=0, _
                    AddressType:=AddressType, _
                    SheetName:=sheet_name, _
                    BookName:=book_name)
        End If
    End If
End Function
Private Function IStringable_ToString() As String
    IStringable_ToString = ToString()
End Function

'
' Returns whether this object and another object are equal.
'
Public Function Equals(ByVal OtherObject As Object) As Boolean
    Call pCheckInit

    If Not TypeOf OtherObject Is WorksheetRangeBounds Then
        Equals = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = OtherObject
    If Me.GetIdentityString() = range_bounds.GetIdentityString() Then
        Equals = True
    Else
        Equals = False
    End If
End Function
Private Function IEquatable_Equals(ByVal OtherObject As Object) As Boolean
    IEquatable_Equals = Equals(OtherObject)
End Function

'
' Returns a string representation used to determine the identity of this object.
'
Public Function GetIdentityString() As String
    Call pCheckInit
    GetIdentityString = LCase$(ToString(CellOnly:=False))
End Function
Private Function IEquatable_GetIdentityString() As String
    IEquatable_GetIdentityString = GetIdentityString()
End Function

'* Gets the enumerator for this range.
'*
'* @param EnumerationMode [Optional] Enumeration mode. Default is row-by-row.
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @return Enumerator object.
'*
'* @details
'* Gets the enumerator for this range. The enumeration mode is determined only by the arguments specified on this call.
Public Function GetEnumerator( _
        Optional ByVal EnumerationMode As Long = G_RANGE_ENUM_MODE_ROWS, _
        Optional ByVal Descending As Boolean = False) As IEnumerator
    Call pCheckInit
    Dim result_value As WorksheetRangeBoundsEnumerator
    Set result_value = New WorksheetRangeBoundsEnumerator
    Call result_value.Initialize(Me, EnumerationMode:=EnumerationMode, Descending:=Descending)
    Set GetEnumerator = result_value
End Function

Private Function pGetBoundsList(ByVal EnumerationMode As Long, ByVal Descending As Boolean) As ObjectList
    Call pCheckInit

    Dim result_value As ObjectList
    Set result_value = New_ObjectList("WorksheetRangeBounds")

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator(EnumerationMode:=EnumerationMode, Descending:=Descending)
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pGetBoundsList = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pCheckInit", "The object is not initialized."
End Sub

Private Function pGetTargetRow(ByVal RowIndex As Long) As Long
    Dim row_count_value As Long
    row_count_value = RowCount

    If RowIndex = 0 Or row_count_value < RowIndex Or RowIndex < -row_count_value Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pGetTargetRow", "The row index is outside the row range. (index:" & RowIndex & ", start: " & pRow & ", finish: " & pFinishRow & ")"
    End If

    Dim target_row_value As Double
    If 0 < RowIndex Then
        target_row_value = CDbl(pRow) + CDbl(RowIndex) - 1#
    Else
        target_row_value = CDbl(pFinishRow) + CDbl(RowIndex) + 1#
    End If

    Dim result_value As Long
    result_value = CLng(target_row_value)

    pGetTargetRow = result_value
End Function

Private Function pGetTargetColumn(ByVal ColumnIndex As Long) As Long
    Dim column_count_value As Long
    column_count_value = ColumnCount

    If ColumnIndex = 0 Or column_count_value < ColumnIndex Or ColumnIndex < -column_count_value Then
        Err.Raise vbObjectError + 1, "Class WorksheetRangeBounds.pGetTargetColumn", "The column index is outside the column range. (index:" & ColumnIndex & ", start: " & pColumn & ", finish: " & pFinishColumn & ")"
    End If

    Dim target_column_value As Double
    If 0 < ColumnIndex Then
        target_column_value = CDbl(pColumn) + CDbl(ColumnIndex) - 1#
    Else
        target_column_value = CDbl(pFinishColumn) + CDbl(ColumnIndex) + 1#
    End If

    Dim result_value As Long
    result_value = CLng(target_column_value)

    pGetTargetColumn = result_value
End Function

Private Sub pRequireWorkbookService(ByVal MethodName As String)
    If WbSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class WorksheetRangeBounds" & "." & MethodName, Description:="InitializeCommonService has not been called."
    End If
End Sub

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator(EnumerationMode:=G_RANGE_ENUM_MODE_CELLS_HORIZONTAL, Descending:=False)
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function
