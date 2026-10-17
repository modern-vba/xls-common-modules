VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorksheetService

' #############################################################################
'!
'! @brief
'! Implementation of the service that manages worksheet access.
'!
' #############################################################################

'
' Searches a range.
'
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Dim result_value As ObjectList
    Set result_value = New ObjectList

    If RangeBounds.IsEmpty Then
        Set Find = result_value
        Exit Function
    End If

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Dim look_in As Long
    If LookInValues Then
        look_in = xlValues
    Else
        look_in = xlFormulas
    End If

    Dim look_at As Long
    If LookAtWhole Then
        look_at = xlWhole
    Else
        look_at = xlPart
    End If

    Dim target_range As Range
    If RangeBounds.IsEntireSheet Then
        ' Search the entire sheet.
        Set target_range = target_sheet.Cells
    Else
        ' Search the range.
        Set target_range = target_sheet.Range(RangeBounds.ToString(CellOnly:=True))
    End If

    ' Search for the first item.
    Dim first_cell As Range
    Set first_cell = target_range.Find( _
            What:=What, _
            After:=target_range.Cells(target_range.Rows.Count, target_range.Columns.Count), _
            LookIn:=look_in, _
            LookAt:=look_at, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte, _
            SearchFormat:=False)

    ' If nothing is found, return an empty list.
    If first_cell Is Nothing Then
        Call pResetFindState(target_sheet)
        Set Find = result_value
        Exit Function
    End If

    Dim found_cell As Range
    Set found_cell = first_cell
    Do
        Dim found_range As WorksheetRangeBounds
        Set found_range = New_RangeBounds(Row:=found_cell.Row, Column:=found_cell.Column, Sheet:=found_cell.Worksheet.Name, Book:=found_cell.Worksheet.Parent.Name)
        Call result_value.Add(found_range)

        Set found_cell = target_range.FindNext(After:=found_cell)
    Loop While found_cell.Address(External:=True) <> first_cell.Address(External:=True)

    ' Disable exact matching. (Exact matching is hard to use in normal usage.)
    Call pResetFindState(target_sheet)

    Set Find = result_value
End Function
Private Function IWorksheetService_Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As ObjectList

    Set IWorksheetService_Find = Find( _
            What, RangeBounds, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub
Private Sub IWorksheetService_Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub

Private Sub pSortCore(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    If IsEmptyArray(SortKeyAndOrder) Then Exit Sub
    If RangeBounds.IsEmpty Then Exit Sub
    If (UBound(SortKeyAndOrder) - LBound(SortKeyAndOrder)) Mod 2 = 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pSortCore", "SortKeyAndOrder must contain an even number of elements. (Odd-numbered elements are column numbers used as sort keys; even-numbered elements are xlAscending or xlDescending.)"
    End If
    Call pRaiseIfSortKeyAndOrderInvalid(RangeBounds, SortKeyAndOrder)

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Dim sort_range As Range
    Set sort_range = target_sheet.Range(RangeBounds.ToString(CellOnly:=True))

    With target_sheet.Sort
        .SortFields.Clear

        Dim item_idx As Long
        For item_idx = LBound(SortKeyAndOrder) To UBound(SortKeyAndOrder) - 1 Step 2
            Dim sort_key As Long
            sort_key = CLng(SortKeyAndOrder(item_idx))
            If IsEmpty(SortKeyAndOrder(item_idx + 1)) Then
                .SortFields.Add Key:=target_sheet.Cells(RangeBounds.Row, RangeBounds.Column + sort_key - 1)
            Else
                .SortFields.Add Key:=target_sheet.Cells(RangeBounds.Row, RangeBounds.Column + sort_key - 1), Order:=CLng(SortKeyAndOrder(item_idx + 1))
            End If
        Next item_idx

        .SetRange sort_range
        .Header = xlNo
        .Apply
    End With
End Sub

' Activates a range.
'
' @details
' Activates a range.
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    If RangeBounds.IsEmpty Then Err.Raise vbObjectError + 1, "Class WorksheetService.ActivateRange", "RangeBounds is an empty range. (" & RangeBounds.ToString() & ")"

    Dim target_book As Workbook
    Set target_book = pGetWorkbook(RangeBounds)

    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(RangeBounds.WorksheetName)

    Call target_book.Activate
    Call target_sheet.Activate
    target_sheet.Range(RangeBounds.ToString(CellOnly:=True)).Activate
End Sub
Private Sub IWorksheetService_ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call ActivateRange(RangeBounds)
End Sub

Private Function pGetWorkbook(ByVal RangeBounds As WorksheetRangeBounds) As Workbook
    If RangeBounds.WorkbookName <> "" Then
        Set pGetWorkbook = Workbooks(RangeBounds.WorkbookName)
    Else
        Set pGetWorkbook = ThisWorkbook
    End If
End Function

'
' Sets a string or formula in the specified cell.
' When TypeConvert is True, numeric strings are treated as integers if they can be represented as integers, or as decimals otherwise.
'
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)

    If RangeBounds.IsEmpty Then Exit Sub
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService.WriteCell", "RangeBounds is not a cell. (" & RangeBounds.ToString() & ")"

    Dim target_cell As Range
    Set target_cell = pGetCell(RangeBounds)

    If Expression = "" And IgnoreEmpty Then Exit Sub

    ' Set the number format.
    If NumberFormat <> "" Then
        target_cell.NumberFormatLocal = NumberFormat
    End If

    ' Process blanks.
    If Expression = "" Then
        If ClearWhenEmpty Then
            target_cell.ClearContents
        Else
            target_cell.Formula = "="""""
            Call target_cell.Copy
            Call target_cell.PasteSpecial(Paste:=xlPasteValues, SkipBlanks:=False)
            Application.CutCopyMode = False
        End If
        Exit Sub
    End If

    If AsFormula Then
        ' Process normal formulas.
        target_cell.Formula = Expression
    Else
        If TypeConvert Then
            ' When type conversion is enabled.
            Dim err_val As Variant
            If TryConvertExcelErrorStringToCVErr(Expression, err_val) Then
                target_cell.Value = err_val  ' Set CVErr(...).
                Exit Sub
            End If

            Dim long_value As Long
            Dim double_value As Double
            Dim long_error As Long
            Dim double_error As Long

            On Error Resume Next

            Err.Clear
            long_value = CLng(Expression)
            long_error = Err.Number

            Err.Clear
            double_value = CDbl(Expression)
            double_error = Err.Number

            On Error GoTo 0

            If long_error = 0 Or double_error = 0 Then
                If long_error = 0 And double_error = 0 Then
                    If double_value - CDbl(long_value) = 0# Then
                        target_cell.Value = long_value
                    Else
                        target_cell.Value = double_value
                    End If
                ElseIf double_error = 0 Then
                    target_cell.Value = double_value
                Else
                    target_cell.Value = long_value
                End If
                Exit Sub
            End If

            Dim bool_value As Boolean
            Dim bool_error As Long
            On Error Resume Next
            Err.Clear
            bool_value = CBool(Expression)
            bool_error = Err.Number
            Err.Clear
            On Error GoTo 0
            If bool_error = 0 Then
                target_cell.Value = bool_value
                Exit Sub
            End If

            Dim date_value As Date
            Dim date_error As Long
            On Error Resume Next
            Err.Clear
            date_value = CDate(Expression)
            date_error = Err.Number
            Err.Clear
            On Error GoTo 0
            If date_error = 0 Then
                target_cell.Value = date_value
                Exit Sub
            End If

            Dim currency_value As Currency
            Dim currency_error As Long
            On Error Resume Next
            Err.Clear
            currency_value = CCur(Expression)
            currency_error = Err.Number
            Err.Clear
            On Error GoTo 0
            If currency_error = 0 Then
                target_cell.Value = currency_value
                Exit Sub
            End If
        End If

        ' Process as text.
        Dim target_was_protected As Boolean
        target_was_protected = target_cell.Worksheet.ProtectContents

        Dim target_had_formula As Boolean
        target_had_formula = target_cell.HasFormula

        Dim target_formula_r1c1 As String
        If target_had_formula Then target_formula_r1c1 = target_cell.FormulaR1C1

        Dim target_value As Variant
        target_value = target_cell.Value2

        Dim current_format As String
        current_format = target_cell.NumberFormatLocal

        Dim current_wrap_text As Boolean
        current_wrap_text = target_cell.WrapText

        On Error GoTo RESTORE_STRING_WRITE_CELL_ON_ERROR
        target_cell.NumberFormatLocal = "@"
        target_cell.Value = Expression
        target_cell.NumberFormatLocal = current_format
        target_cell.WrapText = current_wrap_text
        On Error GoTo 0
    End If
    Exit Sub

RESTORE_STRING_WRITE_CELL_ON_ERROR:
    Dim err_num As Long
    err_num = Err.Number

    Dim err_source As String
    err_source = Err.Source

    Dim err_desc As String
    err_desc = Err.Description

    Call pRestoreCellContentsAndNumberFormat( _
            target_cell, _
            target_was_protected, _
            target_had_formula, _
            target_formula_r1c1, _
            target_value, _
            current_format, _
            current_wrap_text)

    Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc
End Sub
Private Sub IWorksheetService_WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)

    Call WriteCell( _
            RangeBounds, Expression, _
            NumberFormat:=NumberFormat, AsFormula:=AsFormula, ClearWhenEmpty:=ClearWhenEmpty, IgnoreEmpty:=IgnoreEmpty, TypeConvert:=TypeConvert)
End Sub

Private Sub pRestoreCellContentsAndNumberFormat( _
        ByVal TargetCell As Range, _
        ByVal WorksheetWasProtected As Boolean, _
        ByVal HadFormula As Boolean, _
        ByVal FormulaR1C1 As String, _
        ByVal CellValue As Variant, _
        ByVal NumberFormatLocal As String, _
        ByVal WrapText As Boolean)

    Dim original_enable_events As Boolean
    original_enable_events = Application.EnableEvents

    On Error Resume Next
    Application.EnableEvents = False

    If Not WorksheetWasProtected And TargetCell.Worksheet.ProtectContents Then
        Call TargetCell.Worksheet.Unprotect
    End If

    TargetCell.NumberFormatLocal = NumberFormatLocal
    If HadFormula Then
        TargetCell.FormulaR1C1 = FormulaR1C1
    ElseIf IsEmpty(CellValue) Then
        TargetCell.ClearContents
    Else
        TargetCell.Value2 = CellValue
    End If
    TargetCell.NumberFormatLocal = NumberFormatLocal
    TargetCell.WrapText = WrapText

    Application.EnableEvents = original_enable_events
    On Error GoTo 0
End Sub

'
' Writes a two-dimensional array to a range.
'
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Dim value_row_count As Long
    Dim value_column_count As Long
    Call pGetTwoDimensionalArraySize(value_row_count, value_column_count, ValuesArray)
    Call pRaiseIfDestinationShapeIsInvalid(RangeBounds, value_row_count, value_column_count, "RangeBounds")
    Call pRaiseIfDestinationOutOfSheet(RangeBounds, value_row_count, value_column_count, "write")

    Dim target_range As Range
    Set target_range = pGetDestinationRange(RangeBounds, value_row_count, value_column_count)

    target_range.Value = ValuesArray
End Sub
Private Sub IWorksheetService_WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call WriteRange(RangeBounds, ValuesArray)
End Sub

Private Function pGetDestinationRange( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long) As Range

    If DestinationRangeBounds.IsEmpty Or DestinationRangeBounds.IsCell Then
        Set pGetDestinationRange = pGetCell(DestinationRangeBounds).Resize(SourceRowCount, SourceColumnCount)
    Else
        Set pGetDestinationRange = pGetRange(DestinationRangeBounds)
    End If
End Function

Private Sub pGetTwoDimensionalArraySize( _
        ByRef RowCount As Long, _
        ByRef ColumnCount As Long, _
        ByVal TargetArray As Variant)

    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pGetTwoDimensionalArraySize", "ValuesArray must be a two-dimensional array. (" & TypeName(TargetArray) & ")"
    End If

    Dim lower_bounds() As Long
    Dim upper_bounds() As Long
    Call GetArrayBounds(lower_bounds, upper_bounds, TargetArray)

    If IsEmptyArray(lower_bounds) Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pGetTwoDimensionalArraySize", "ValuesArray must be a two-dimensional array."
    End If
    If UBound(lower_bounds) - LBound(lower_bounds) + 1 <> 2 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pGetTwoDimensionalArraySize", "ValuesArray must be a two-dimensional array."
    End If

    RowCount = upper_bounds(LBound(upper_bounds)) - lower_bounds(LBound(lower_bounds)) + 1
    ColumnCount = upper_bounds(LBound(upper_bounds) + 1) - lower_bounds(LBound(lower_bounds) + 1) + 1

    If RowCount < 1 Or ColumnCount < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pGetTwoDimensionalArraySize", "ValuesArray must be a two-dimensional array with at least one element."
    End If
End Sub

Private Sub pRaiseIfDestinationOutOfSheet( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long, _
        ByVal OperationName As String)

    Dim destination_finish_row As Long
    destination_finish_row = DestinationRangeBounds.Row + SourceRowCount - 1
    If G_ROW_MAX < destination_finish_row Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pRaiseIfDestinationOutOfSheet", "The " & OperationName & " destination range finish row exceeds the sheet limit. (row: " & DestinationRangeBounds.Row & ", row_count: " & SourceRowCount & ", finish_row: " & destination_finish_row & ")"
    End If

    Dim destination_finish_column As Long
    destination_finish_column = DestinationRangeBounds.Column + SourceColumnCount - 1
    If G_COL_MAX < destination_finish_column Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pRaiseIfDestinationOutOfSheet", "The " & OperationName & " destination range finish column exceeds the sheet limit. (column: " & DestinationRangeBounds.Column & ", column_count: " & SourceColumnCount & ", finish_column: " & destination_finish_column & ")"
    End If
End Sub

'
' Writes an array (CSE) formula to a range.
'
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    target_range.FormulaArray = FormulaArray
End Sub
Private Sub IWorksheetService_WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call WriteArrayFormula(RangeBounds, FormulaArray)
End Sub

'
' Reads a value from a cell.
'
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    Expression = ""
    NumberFormat = ""

    If RangeBounds.IsEmpty Then
        ReadCell = ""
        Exit Function
    End If
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService.ReadCell", "RangeBounds is not a cell. (" & RangeBounds.ToString() & ")"

    Dim target_cell As Range
    Set target_cell = pGetCell(RangeBounds)

    If GetText Then
        Call pReadCellText(target_cell, Expression, NumberFormat)
    Else
        Call pReadCellValue(target_cell, Expression, NumberFormat)
    End If
    ReadCell = Expression
End Function
Private Function IWorksheetService_ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    IWorksheetService_ReadCell = ReadCell(RangeBounds, Expression:=Expression, NumberFormat:=NumberFormat, GetText:=GetText)
End Function

Private Sub pReadCellValue(ByVal TargetCell As Range, ByRef Expression As String, ByRef NumberFormat As String)
    If IsError(TargetCell.Value) Then
        Expression = ExcelErrorToString(TargetCell.Value)
    Else
        Expression = CStr(TargetCell.Value)
    End If
    NumberFormat = TargetCell.NumberFormatLocal
End Sub

Private Sub pReadCellText(ByVal TargetCell As Range, ByRef Expression As String, ByRef NumberFormat As String)
    Dim result_string As String
    result_string = TargetCell.Text

    Dim target_column As Range
    Dim current_width As Double
    Dim width_changed As Boolean

    If pIsHashOnlyString(result_string) And Not pIsDisplayedLiteralString(TargetCell, result_string) Then
        Set target_column = TargetCell.Parent.Columns(TargetCell.Column)
        current_width = target_column.ColumnWidth

        On Error Resume Next
        Err.Clear
        target_column.ColumnWidth = 254.88

        Dim width_error_number As Long
        Dim width_error_source As String
        Dim width_error_description As String
        width_error_number = Err.Number
        width_error_source = Err.Source
        width_error_description = Err.Description
        On Error GoTo 0

        If width_error_number <> 0 Then
            Call pRaiseReadCellTextColumnWidthError( _
                    TargetCell, _
                    result_string, _
                    width_error_number, _
                    width_error_source, _
                    width_error_description)
        End If

        On Error GoTo RESTORE_COLUMN_WIDTH
        width_changed = True
        result_string = TargetCell.Text
        target_column.ColumnWidth = current_width
        width_changed = False
        On Error GoTo 0

        If pIsHashOnlyString(result_string) Then
            If IsError(TargetCell.Value) Then
                result_string = ExcelErrorToString(TargetCell.Value)
            Else
                result_string = CStr(TargetCell.Value)
            End If
        End If
    End If

    Expression = result_string
    NumberFormat = TargetCell.NumberFormatLocal
    Exit Sub

RESTORE_COLUMN_WIDTH:
    Dim err_number As Long
    Dim err_source As String
    Dim err_description As String
    err_number = Err.Number
    err_source = Err.Source
    err_description = Err.Description

    On Error Resume Next
    If width_changed Then target_column.ColumnWidth = current_width
    On Error GoTo 0
    Err.Raise err_number, err_source, err_description
End Sub

'
' Reads values from a range.
'
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    If RangeBounds.IsEmpty Then Exit Function

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    If RangeBounds.IsCell Then
        Dim result_value(1 To 1, 1 To 1) As Variant
        result_value(1, 1) = target_range.Value
        ReadRange = result_value
    Else
        ReadRange = target_range.Value
    End If
End Function
Private Function IWorksheetService_ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    IWorksheetService_ReadRange = ReadRange(RangeBounds)
End Function

'
' Evaluates a formula.
'
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    If ContextRangeBounds.IsEmpty Then Exit Function

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(ContextRangeBounds)

    EvaluateFormula = target_sheet.Evaluate(Formula)
End Function
Private Function IWorksheetService_EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    IWorksheetService_EvaluateFormula = EvaluateFormula(ContextRangeBounds, Formula)
End Function

'
' Gets the XLOOKUP result.
'
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant

    If LookupRangeBounds.IsEmpty Then Exit Function
    If ReturnRangeBounds.IsEmpty Then Exit Function

    Dim xlookup_formula As String
    xlookup_formula = "=XLOOKUP(" & _
            pGetFormulaLiteral(LookupValue) & "," & _
            pGetFormulaRangeAddress(LookupRangeBounds) & "," & _
            pGetFormulaRangeAddress(ReturnRangeBounds) & "," & _
            pGetFormulaLiteral(IfNotFound) & "," & _
            CStr(MatchMode) & "," & _
            CStr(SearchMode) & ")"

    XLookup = EvaluateFormula(LookupRangeBounds, xlookup_formula)
End Function
Private Function IWorksheetService_XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant

    IWorksheetService_XLookup = XLookup( _
            LookupValue, _
            LookupRangeBounds, _
            ReturnRangeBounds, _
            IfNotFound:=IfNotFound, _
            MatchMode:=MatchMode, _
            SearchMode:=SearchMode)
End Function

Private Function pGetFormulaRangeAddress(ByVal RangeBounds As WorksheetRangeBounds) As String
    pGetFormulaRangeAddress = pGetRange(RangeBounds).Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1, _
            External:=True)
End Function

Private Function pGetFormulaLiteral(ByVal Value As Variant) As String
    If IsError(Value) Then
        pGetFormulaLiteral = ExcelErrorToString(Value)
        Exit Function
    End If
    If IsEmpty(Value) Or IsNull(Value) Then
        pGetFormulaLiteral = """"""
        Exit Function
    End If

    Select Case VarType(Value)
     Case vbBoolean
        If CBool(Value) Then
            pGetFormulaLiteral = "TRUE"
        Else
            pGetFormulaLiteral = "FALSE"
        End If
     Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        pGetFormulaLiteral = pGetFormulaNumber(CDbl(Value))
     Case vbDate
        pGetFormulaLiteral = pGetFormulaNumber(CDbl(CDate(Value)))
     Case Else
        pGetFormulaLiteral = """" & Replace(CStr(Value), """", """""") & """"
    End Select
End Function

'
' Copies the value or formula in the specified cell.
' Formula copying also supports array formulas.
'
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    If SourceRangeBounds.IsEmpty Then Exit Sub
    If DestinationRangeBounds.IsEmpty Then Exit Sub
    If Not SourceRangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService.CopyCell", "SourceRangeBounds is not a cell. (" & SourceRangeBounds.ToString() & ")"
    If Not DestinationRangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService.CopyCell", "DestinationRangeBounds is not a cell. (" & DestinationRangeBounds.ToString() & ")"

    Dim src_sheet As Worksheet
    Set src_sheet = pGetWorksheet(SourceRangeBounds)

    Dim dst_sheet As Worksheet
    Set dst_sheet = pGetWorksheet(DestinationRangeBounds)

    Call pCopyCellCore( _
        SourceRangeBounds.Row, SourceRangeBounds.Column, _
        DestinationRangeBounds.Row, DestinationRangeBounds.Column, _
        src_sheet, dst_sheet, AsValue, CopyNumberFormat, RemainRowCount:=1, RemainColumnCount:=1, ForcePasteIfArrayFormula:=True)
End Sub
Private Sub IWorksheetService_CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call CopyCell(SourceRangeBounds, DestinationRangeBounds, AsValue:=AsValue, CopyNumberFormat:=CopyNumberFormat)
End Sub

'
' Copies the value or formula in the specified range.
' Formula copying also supports array formulas.
'
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    If SourceRangeBounds.IsEmpty Then Exit Sub
    Call pRaiseIfDestinationShapeIsInvalid(DestinationRangeBounds, SourceRangeBounds.RowCount, SourceRangeBounds.ColumnCount, "DestinationRangeBounds")
    Call pRaiseIfCopyRangeDestinationOutOfSheet(SourceRangeBounds, DestinationRangeBounds)

    Dim src_sheet As Worksheet
    Set src_sheet = pGetWorksheet(SourceRangeBounds)

    Dim dst_sheet As Worksheet
    Set dst_sheet = pGetWorksheet(DestinationRangeBounds)

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = GetUsedRangeBounds(SourceRangeBounds, IgnoreEmpty:=False, IsUsedStartRow:=True, IsUsedStartColumn:=True)

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds( _
            Row:=DestinationRangeBounds.Row, Column:=DestinationRangeBounds.Column, _
            FinishRow:=DestinationRangeBounds.Row + SourceRangeBounds.RowCount - 1, _
            FinishColumn:=DestinationRangeBounds.Column + SourceRangeBounds.ColumnCount - 1, _
            Sheet:=DestinationRangeBounds.WorksheetName, Book:=DestinationRangeBounds.WorkbookName)
    Set dst_bounds = GetUsedRangeBounds(dst_bounds, IgnoreEmpty:=False, IsUsedStartRow:=True, IsUsedStartColumn:=True)

    Dim last_row_idx As Long, last_col_idx As Long
    last_row_idx = MaxLng(src_bounds.RowCount, dst_bounds.RowCount) - 1
    last_col_idx = MaxLng(src_bounds.ColumnCount, dst_bounds.ColumnCount) - 1

    Dim row_start_idx As Long, row_end_idx As Long, row_step As Long
    row_start_idx = 0
    row_end_idx = last_row_idx
    row_step = 1

    Dim col_start_idx As Long, col_end_idx As Long, col_step As Long
    col_start_idx = 0
    col_end_idx = last_col_idx
    col_step = 1

    If src_sheet Is dst_sheet Then
        If src_bounds.Row < dst_bounds.Row Then
            row_start_idx = last_row_idx
            row_end_idx = 0
            row_step = -1
        End If

        If src_bounds.Column < dst_bounds.Column Then
            col_start_idx = last_col_idx
            col_end_idx = 0
            col_step = -1
        End If
    End If

    ' Prepare an array to record processed array formulas.
    Dim procd_arr() As String: ReDim procd_arr(0)
    Dim row_idx As Long, col_idx As Long
    For row_idx = row_start_idx To row_end_idx Step row_step
        For col_idx = col_start_idx To col_end_idx Step col_step
            ' Get the anchor cell of the array formula. (Nothing if it is not an array formula.)
            Dim arr_anchor As Range
            Set arr_anchor = pGetArrayFormulaAnchor(src_sheet, src_bounds.Row + row_idx, src_bounds.Column + col_idx)
            If pIsArrayFormulaToBeProcessed(arr_anchor, procd_arr) Then
                ' When this is an array formula and it has not been processed yet,
                Dim arr_rows As Long, arr_cols As Long
                ' calculate the size of the array formula at the destination.
                Call pCalculateArrayCount(arr_rows, arr_cols, arr_anchor, SourceRangeBounds)
                ' Force processing as an array formula.
                Call pCopyCellCore( _
                    src_bounds.Row + row_idx, src_bounds.Column + col_idx, _
                    dst_bounds.Row + row_idx, dst_bounds.Column + col_idx, _
                    src_sheet, dst_sheet, AsValue, CopyNumberFormat, _
                    RemainRowCount:=arr_rows, _
                    RemainColumnCount:=arr_cols, _
                    ForcePasteIfArrayFormula:=True)
                ' Record it in the processed array.
                Call pAddItemToStringArray(procd_arr, arr_anchor.Address)
            Else
                ' Normal processing.
                Call pCopyCellCore( _
                    src_bounds.Row + row_idx, src_bounds.Column + col_idx, _
                    dst_bounds.Row + row_idx, dst_bounds.Column + col_idx, _
                    src_sheet, dst_sheet, AsValue, CopyNumberFormat)
            End If
        Next col_idx
    Next row_idx
End Sub
Private Sub IWorksheetService_CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)

    Call CopyRange(SourceRangeBounds, DestinationRangeBounds, AsValue:=AsValue, CopyNumberFormat:=CopyNumberFormat)
End Sub

Private Sub pRaiseIfCopyRangeDestinationOutOfSheet( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds)

    Call pRaiseIfDestinationOutOfSheet(DestinationRangeBounds, SourceRangeBounds.RowCount, SourceRangeBounds.ColumnCount, "copy")
End Sub

Private Sub pAddItemToStringArray(ByRef TargetArray() As String, ByVal ItemToBeAdded As String)
    Dim new_ubound As Long

    If (Not TargetArray) = -1 Then
        ReDim Preserve TargetArray(0 To 0) As String
    Else
        new_ubound = UBound(TargetArray) + 1
        ReDim Preserve TargetArray(LBound(TargetArray) To new_ubound) As String
    End If
    TargetArray(new_ubound) = ItemToBeAdded
End Sub

Private Function pGetArrayFormulaAnchor(ByVal TargetSheet As Worksheet, ByVal RowIndex As Long, ByVal ColumnIndex As Long) As Range
    If TargetSheet.Cells(RowIndex, ColumnIndex).HasArray Then
        Set pGetArrayFormulaAnchor = TargetSheet.Cells(RowIndex, ColumnIndex).CurrentArray.Cells(1, 1)
    Else
        Set pGetArrayFormulaAnchor = Nothing
    End If
End Function

Private Function pIsArrayFormulaToBeProcessed(ByVal TargetCell As Range, ByRef ProcessedArray() As String) As Boolean
    If TargetCell Is Nothing Then
        pIsArrayFormulaToBeProcessed = False
        Exit Function
    End If

    If pContainsItemInStringArray(ProcessedArray, TargetCell.Address) Then
        pIsArrayFormulaToBeProcessed = False
    Else
        pIsArrayFormulaToBeProcessed = True
    End If
End Function

Private Sub pCalculateArrayCount(ByRef ArrayRowCount As Long, ByRef ArrayColumnCount As Long, ByVal TargetArrayAnchor As Range, ByVal RangeBounds As WorksheetRangeBounds)
    Dim array_last_row As Long, array_last_col As Long
    array_last_row = TargetArrayAnchor.Row + TargetArrayAnchor.CurrentArray.Rows.Count - 1
    array_last_col = TargetArrayAnchor.Column + TargetArrayAnchor.CurrentArray.Columns.Count - 1

    Dim paste_start_row As Long, paste_start_col As Long
    paste_start_row = MaxLng(TargetArrayAnchor.Row, RangeBounds.Row)
    paste_start_col = MaxLng(TargetArrayAnchor.Column, RangeBounds.Column)

    Dim paste_finish_row As Long, paste_finish_col As Long
    paste_finish_row = MinLng(array_last_row, RangeBounds.FinishRow)
    paste_finish_col = MinLng(array_last_col, RangeBounds.FinishColumn)

    ArrayRowCount = MaxLng(paste_finish_row - paste_start_row + 1, 0)
    ArrayColumnCount = MaxLng(paste_finish_col - paste_start_col + 1, 0)
End Sub

'
' Checks whether the cell is empty. When IgnoreEmptyString is True, the cell is considered empty if it contains a formula whose calculated result is an empty string.
'
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    If RangeBounds.IsEmpty Then
        IsEmptyCell = True
        Exit Function
    End If

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    Dim cell_item As Range
    For Each cell_item In target_range
        If Not pIsEmptyCellCore(cell_item, IgnoreEmptyString:=IgnoreEmptyString) Then
            IsEmptyCell = False
            Exit Function
        End If
    Next cell_item

    IsEmptyCell = True
End Function
Private Function IWorksheetService_IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    IWorksheetService_IsEmptyCell = IsEmptyCell(RangeBounds, IgnoreEmptyString:=IgnoreEmptyString)
End Function

Private Function pIsEmptyCellCore(ByVal TargetCell As Range, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    Dim cell_value As Variant
    cell_value = TargetCell.Value

    If IsEmpty(cell_value) Then
        pIsEmptyCellCore = True
        Exit Function
    End If

    If Not IgnoreEmptyString Then
        pIsEmptyCellCore = False
        Exit Function
    End If

    If VarType(cell_value) = vbString Then
        pIsEmptyCellCore = (Len(cell_value) = 0)
    Else
        pIsEmptyCellCore = False
    End If
End Function

'
' Checks whether a formula is entered in the cell.
'
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    If RangeBounds.IsEmpty Then Exit Function
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService.HasFormula", "RangeBounds is not a cell. (" & RangeBounds.ToString() & ")"
    HasFormula = pGetCell(RangeBounds).HasFormula
End Function
Private Function IWorksheetService_HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    IWorksheetService_HasFormula = HasFormula(RangeBounds)
End Function

'
' Returns the actual used range within the specified range.
'
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds

    ' Return an empty range as-is.
    If RangeBounds.IsEmpty Then
        Set GetUsedRangeBounds = RangeBounds
        Exit Function
    End If

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    ' Get the raw UsedRange.
    Dim raw_range As WorksheetRangeBounds
    Set raw_range = pGetRawUsedRange(target_sheet, RangeBounds, IsUsedStartRow, IsUsedStartColumn)

    ' If GetRawRange is True or UsedRange is empty at this point, finish processing here.
    If GetRawRange Or raw_range.IsEmpty Then
        Set GetUsedRangeBounds = raw_range
        Exit Function
    End If

    Dim row_start As Long, row_finish As Long
    Dim col_start As Long, col_finish As Long

    ' Get the last row.
    row_finish = pGetLastRowCore(target_sheet, raw_range.Row, raw_range.Column, raw_range.FinishRow, raw_range.FinishColumn, IgnoreEmpty)

    ' Get the first row.
    If IsUsedStartRow Then
        row_start = RangeBounds.Row
    Else
        row_start = pGetFirstRowCore(target_sheet, raw_range.Row, raw_range.Column, row_finish, raw_range.FinishColumn, IgnoreEmpty)
    End If

    ' Get the last column.
    col_finish = pGetLastColumnCore(target_sheet, row_start, raw_range.Column, row_finish, raw_range.FinishColumn, IgnoreEmpty)

    ' Get the first column.
    If IsUsedStartColumn Then
        col_start = RangeBounds.Column
    Else
        col_start = pGetFirstColumnCore(target_sheet, row_start, raw_range.Column, row_finish, col_finish, IgnoreEmpty)
    End If

    ' Normalize the processing result.
    If row_finish < row_start Or col_finish < col_start Then
        Set GetUsedRangeBounds = New_RangeBounds(Row:=RangeBounds.Row, Column:=RangeBounds.Column, FinishRow:=0, FinishColumn:=0, Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    Else
        Set GetUsedRangeBounds = New_RangeBounds(Row:=row_start, Column:=col_start, FinishRow:=row_finish, FinishColumn:=col_finish, Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    End If

    ' Disable exact matching. (Exact matching is hard to use in normal usage.)
    Call pResetFindState(target_sheet)
End Function
Private Function IWorksheetService_GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds

    Set IWorksheetService_GetUsedRangeBounds = GetUsedRangeBounds(RangeBounds, IgnoreEmpty:=IgnoreEmpty, IsUsedStartRow:=IsUsedStartRow, IsUsedStartColumn:=IsUsedStartColumn, GetRawRange:=GetRawRange)
End Function

Private Function pGetRawUsedRange( _
        ByVal TargetSheet As Worksheet, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal IsUsedStartRow As Boolean, _
        ByVal IsUsedStartColumn As Boolean) As WorksheetRangeBounds

    Dim used_s_row As Long, used_f_row As Long
    Dim used_s_col As Long, used_f_col As Long
    If 0 < TargetSheet.UsedRange.Rows.Count Then
        used_s_row = TargetSheet.UsedRange.Row
        used_f_row = TargetSheet.UsedRange.Rows(TargetSheet.UsedRange.Rows.Count).Row
    Else
        used_s_row = 1
        used_s_col = 1
    End If
    If 0 < TargetSheet.UsedRange.Columns.Count Then
        used_s_col = TargetSheet.UsedRange.Column
        used_f_col = TargetSheet.UsedRange.Columns(TargetSheet.UsedRange.Columns.Count).Column
    Else
        used_s_row = 1
        used_s_col = 1
    End If

    Dim used_bounds As WorksheetRangeBounds
    If used_s_row = used_f_row And used_s_col = used_f_col _
            And IsEmpty(TargetSheet.Cells(used_s_row, used_s_col).Value) _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeTop).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeRight).LineStyle = xlLineStyleNone Then
        ' Empty region.
        Set used_bounds = New_RangeBounds( _
                Row:=used_s_row, _
                Column:=used_s_col, _
                FinishRow:=0, _
                FinishColumn:=0, _
                Sheet:=TargetSheet.Name, _
                Book:=TargetSheet.Parent.Name)
    Else
        ' Normal region.
        Set used_bounds = New_RangeBounds( _
                Row:=used_s_row, _
                Column:=used_s_col, _
                FinishRow:=used_f_row, _
                FinishColumn:=used_f_col, _
                Sheet:=TargetSheet.Name, _
                Book:=TargetSheet.Parent.Name)
    End If

    ' Intersect with the original range.
    Dim intersect_bounds As WorksheetRangeBounds
    Set intersect_bounds = RangeBounds.Intersect(used_bounds)

    If IsUsedStartRow And IsUsedStartColumn Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=RangeBounds.Row, Column:=RangeBounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    ElseIf IsUsedStartRow Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=RangeBounds.Row, Column:=intersect_bounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    ElseIf IsUsedStartColumn Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=intersect_bounds.Row, Column:=RangeBounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    Else
        Set pGetRawUsedRange = intersect_bounds
    End If
End Function

Private Function pGetLastRowCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long

    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetLastRowCore = 0
        Exit Function
    End If

    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )

    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If

    Dim fnd_cell As Range
    ' Last row -> SearchOrder=xlByRows, SearchDirection=xlPrevious.
    ' After is the upper-left cell of the search range.
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(1), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)

    If fnd_cell Is Nothing Then
        pGetLastRowCore = 0
    Else
        pGetLastRowCore = fnd_cell.Row
    End If
End Function

Private Function pGetLastColumnCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long

    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetLastColumnCore = 0
        Exit Function
    End If

    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )

    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If

    Dim fnd_cell As Range
    ' Last column -> SearchOrder=xlByColumns, SearchDirection=xlPrevious.
    ' After is the upper-left cell of the search range.
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(1), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)

    If fnd_cell Is Nothing Then
        pGetLastColumnCore = 0
    Else
        pGetLastColumnCore = fnd_cell.Column
    End If
End Function

Private Function pGetFirstRowCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long

    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetFirstRowCore = 2147483647#
        Exit Function
    End If

    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )

    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If

    Dim fnd_cell As Range
    ' First row -> SearchOrder=xlByRows, SearchDirection=xlNext.
    ' After is the lower-right cell of the search range.
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(srch_rng.Rows.Count, srch_rng.Columns.Count), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlNext, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)

    If fnd_cell Is Nothing Then
        pGetFirstRowCore = 2147483647#
    Else
        pGetFirstRowCore = fnd_cell.Row
    End If
End Function

Private Function pGetFirstColumnCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long

    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetFirstColumnCore = 2147483647#
        Exit Function
    End If

    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )

    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If

    Dim fnd_cell As Range
    ' First column -> SearchOrder=xlByColumns, SearchDirection=xlNext.
    ' After is the lower-right cell of the search range.
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(srch_rng.Rows.Count, srch_rng.Columns.Count), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlNext, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)

    If fnd_cell Is Nothing Then
        pGetFirstColumnCore = 2147483647#
    Else
        pGetFirstColumnCore = fnd_cell.Column
    End If
End Function

'
' Clears the specified range.
'
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)

    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    ' Process the entire range for now.
    Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)

    ' Get the used range.
    Dim used_row As Long, used_col As Long
    If 0 < target_sheet.UsedRange.Rows.Count Then
        used_row = target_sheet.UsedRange.Rows(target_sheet.UsedRange.Rows.Count).Row
    End If
    If 0 < target_sheet.UsedRange.Columns.Count Then
        used_col = target_sheet.UsedRange.Columns(target_sheet.UsedRange.Columns.Count).Column
    End If

    Dim item_idx As Long
    For item_idx = RangeBounds.Row To MinLng(used_row, RangeBounds.FinishRow)
        If target_sheet.Rows(item_idx).Hidden Then
            ' Process hidden rows within the range.
            Set target_range = target_sheet.Range(target_sheet.Cells(item_idx, RangeBounds.Column), target_sheet.Cells(item_idx, RangeBounds.FinishColumn))
            Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
        End If
    Next item_idx

    For item_idx = RangeBounds.Column To MinLng(used_col, RangeBounds.FinishColumn)
        If target_sheet.Columns(item_idx).Hidden Then
            ' Process hidden columns within the range.
            Set target_range = target_sheet.Range(target_sheet.Cells(RangeBounds.Row, item_idx), target_sheet.Cells(RangeBounds.FinishRow, item_idx))
            Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
        End If
    Next item_idx
End Sub
Private Sub IWorksheetService_ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)

    Call ClearRange( _
            RangeBounds, ClearAll:=ClearAll, _
            ClearContents:=ClearContents, ClearNumberFormats:=ClearNumberFormats, ClearColors:=ClearColors, _
            ClearComments:=ClearComments, ClearHyperlinks:=ClearHyperlinks)
End Sub

Private Sub pClearRangeCore( _
        ByVal TargetRange As Range, _
        ByVal ClearAll As Boolean, ByVal ClearContents As Boolean, ByVal ClearNumberFormats As Boolean, _
        ByVal ClearColors As Boolean, ByVal ClearComments As Boolean, ByVal ClearHyperlinks As Boolean)

    If ClearAll Or ClearContents Then Call TargetRange.ClearContents
    If ClearAll Or ClearNumberFormats Then TargetRange.NumberFormat = "General"
    If ClearAll Or ClearColors Then
        TargetRange.Font.ColorIndex = 0
        TargetRange.Interior.ColorIndex = xlColorIndexNone
    End If
    If ClearAll Or ClearComments Then Call TargetRange.ClearComments
    If ClearAll Or ClearHyperlinks Then Call TargetRange.ClearHyperlinks
End Sub

'
' Removes duplicate rows within the range.
'
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    Call target_range.RemoveDuplicates(Columns:=pBuildDuplicateColumnsArgument(RangeBounds, DuplicateColumns), Header:=xlNo)
End Sub
Private Sub IWorksheetService_RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call RemoveDuplicates(RangeBounds, DuplicateColumns)
End Sub

Private Function pBuildDuplicateColumnsArgument(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant) As Variant
    If IsArray(DuplicateColumns) Then
        If IsEmptyArray(DuplicateColumns) Then
            Err.Raise vbObjectError + 1, "Class WorksheetService.pBuildDuplicateColumnsArgument", "DuplicateColumns must contain one or more column numbers."
        End If

        Dim result_columns() As Variant
        Dim column_count As Long
        Dim column_item As Variant
        For Each column_item In DuplicateColumns
            If column_count = 0 Then
                ReDim result_columns(0 To 0)
            Else
                ReDim Preserve result_columns(0 To column_count)
            End If
            result_columns(column_count) = pGetValidatedRelativeColumnIndex(RangeBounds, column_item, "DuplicateColumns")
            column_count = column_count + 1
        Next column_item

        If column_count = 0 Then
            Err.Raise vbObjectError + 1, "Class WorksheetService.pBuildDuplicateColumnsArgument", "DuplicateColumns must contain one or more column numbers."
        End If

        pBuildDuplicateColumnsArgument = result_columns
    Else
        pBuildDuplicateColumnsArgument = pGetValidatedRelativeColumnIndex(RangeBounds, DuplicateColumns, "DuplicateColumns")
    End If
End Function

'
' Inserts rows.
'
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call pRaiseIfNotEntireRowRange(RangeBounds, "InsertRows")

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Call target_sheet.Range(CStr(RangeBounds.Row) & ":" & CStr(RangeBounds.FinishRow)).Insert
End Sub
Private Sub IWorksheetService_InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call InsertRows(RangeBounds)
End Sub

'
' Deletes rows.
'
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call pRaiseIfNotEntireRowRange(RangeBounds, "DeleteRows")

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Call target_sheet.Range(CStr(RangeBounds.Row) & ":" & CStr(RangeBounds.FinishRow)).Delete
End Sub
Private Sub IWorksheetService_DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call DeleteRows(RangeBounds)
End Sub

'
' Shows all data.
' Calls Sheet.ShowAllData, Rows.Hidden = False, and Columns.Hidden = False.
'
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    If target_sheet.FilterMode Then
        Call target_sheet.ShowAllData
    End If
    target_sheet.Rows.Hidden = False
    target_sheet.Columns.Hidden = False
End Sub
Private Sub IWorksheetService_SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call SetAllDataVisible(RangeBounds)
End Sub

'
' Sets the outline levels to display for a configured outline.
'
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    If RowLevels = 0 And ColumnLevels = 0 Then Exit Sub

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    If RowLevels > 0 And ColumnLevels > 0 Then
        Call target_sheet.Outline.ShowLevels(RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
    ElseIf RowLevels > 0 Then
        Call target_sheet.Outline.ShowLevels(RowLevels:=RowLevels)
    Else
        Call target_sheet.Outline.ShowLevels(ColumnLevels:=ColumnLevels)
    End If
End Sub
Private Sub IWorksheetService_SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call SetSheetOutlineLevel(RangeBounds, RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
End Sub

'
' Sets the sheet tab color. If both index and value are specified, the value takes precedence.
'
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    If TabColor >= 0 Then
        target_sheet.Tab.Color = TabColor
    Else
        target_sheet.Tab.ColorIndex = TabColorIndex
    End If
End Sub
Private Sub IWorksheetService_SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)

    Call SetSheetTabColor(RangeBounds, TabColorIndex:=TabColorIndex, TabColor:=TabColor)
End Sub

'
' Sets the range color. If both index and value are specified, the value takes precedence.
'
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    If FontColor >= 0 Then
        target_range.Font.Color = FontColor
    Else
        target_range.Font.ColorIndex = FontColorIndex
    End If

    If InteriorColor >= 0 Then
        target_range.Interior.Color = InteriorColor
    Else
        target_range.Interior.ColorIndex = InteriorColorIndex
    End If
End Sub
Private Sub IWorksheetService_SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)

    Call SetRangeColor(RangeBounds, FontColorIndex:=FontColorIndex, FontColor:=FontColor, InteriorColorIndex:=InteriorColorIndex, InteriorColor:=InteriorColor)
End Sub

'
' Sets whether to wrap text.
'
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    target_range.WrapText = WrapText
End Sub
Private Sub IWorksheetService_SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call SetWrapText(RangeBounds, WrapText)
End Sub

'
' Sets whether to shrink text to fit.
'
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    target_range.ShrinkToFit = ShrinkToFit
End Sub
Private Sub IWorksheetService_SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call SetShrinkToFit(RangeBounds, ShrinkToFit)
End Sub

'
' Sets alignment within cells.
'
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)
    If RangeBounds.IsEmpty Then Exit Sub

    If HorizontalAlignment = G_ALIGNMENT_NO_CHANGE _
            And VerticalAlignment = G_ALIGNMENT_NO_CHANGE _
            And Orientation = G_ALIGNMENT_NO_CHANGE _
            And IndentLevel = G_ALIGNMENT_NO_CHANGE Then
        Exit Sub
    End If

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    If Orientation <> G_ALIGNMENT_NO_CHANGE Then
        target_range.Orientation = Orientation ' Text orientation.
    End If

    If IndentLevel <> G_ALIGNMENT_NO_CHANGE Then
        target_range.IndentLevel = IndentLevel ' Indent.
    End If

    If HorizontalAlignment <> G_ALIGNMENT_NO_CHANGE Then
        target_range.HorizontalAlignment = HorizontalAlignment
    End If

    If VerticalAlignment <> G_ALIGNMENT_NO_CHANGE Then
        target_range.VerticalAlignment = VerticalAlignment
    End If

'    Select Case HorizontalAlignment
'     Case xlLeft ' Left-aligned.
'        target_range.HorizontalAlignment = xlLeft
'     Case xlCenter ' Centered.
'        target_range.HorizontalAlignment = xlCenter
'     Case xlRight ' Right-aligned.
'        target_range.HorizontalAlignment = xlRight
'     Case xlFill ' Repeat.
'        target_range.HorizontalAlignment = xlFill
'     Case xlJustify ' Justified.
'        target_range.HorizontalAlignment = xlJustify
'     Case xlCenterAcrossSelection ' Center across selection.
'        target_range.HorizontalAlignment = xlCenterAcrossSelection
'     Case xlDistributed ' Distributed.
'        target_range.HorizontalAlignment = xlDistributed
'     Case Else ' Standard.
'        target_range.HorizontalAlignment = xlGeneral
'    End Select
'    Select Case VerticalAlignment
'     Case xlTop ' Top-aligned.
'        target_range.VerticalAlignment = xlTop
'     Case xlCenter ' Centered.
'        target_range.VerticalAlignment = xlCenter
'     Case xlJustify ' Justified.
'        target_range.VerticalAlignment = xlJustify
'     Case xlDistributed ' Distributed.
'        target_range.VerticalAlignment = xlDistributed
'     Case Else ' Bottom-aligned.
'        target_range.VerticalAlignment = xlBottom
'    End Select
End Sub
Private Sub IWorksheetService_SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal VerticalAlignment As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal Orientation As Long = G_ALIGNMENT_NO_CHANGE, _
        Optional ByVal IndentLevel As Long = G_ALIGNMENT_NO_CHANGE)

    Call SetAlignment( _
            RangeBounds, _
            HorizontalAlignment:=HorizontalAlignment, _
            VerticalAlignment:=VerticalAlignment, _
            Orientation:=Orientation, _
            IndentLevel:=IndentLevel)
End Sub

Private Sub pResetFindState(ByVal TargetSheet As Worksheet)
    ' Excel Find conditions are shared with the UI and subsequent Range.Find calls, so
    ' after this service call, reset them to the standard state for partial-match value searches.
    ' State after reset: LookIn=xlValues, LookAt=xlPart, SearchOrder=xlByRows,
    ' SearchDirection=xlNext, MatchCase=True, MatchByte=True, SearchFormat=FalseÅB
    Call TargetSheet.Range("A1:A2").Find( _
            What:="", _
            After:=TargetSheet.Range("A1"), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=True, _
            MatchByte:=True, _
            SearchFormat:=False)
End Sub

Private Sub pRaiseIfSortKeyAndOrderInvalid(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    Dim item_idx As Long
    For item_idx = LBound(SortKeyAndOrder) To UBound(SortKeyAndOrder) - 1 Step 2
        Call pGetValidatedRelativeColumnIndex(RangeBounds, SortKeyAndOrder(item_idx), "SortKeyAndOrder")
        If Not IsEmpty(SortKeyAndOrder(item_idx + 1)) Then
            Dim sort_order As Long
            sort_order = pGetValidatedLong(SortKeyAndOrder(item_idx + 1), "SortKeyAndOrder")
            If sort_order <> xlAscending And sort_order <> xlDescending Then
                Err.Raise vbObjectError + 1, "Class WorksheetService.pRaiseIfSortKeyAndOrderInvalid", "SortKeyAndOrder sort order must be xlAscending or xlDescending."
            End If
        End If
    Next item_idx
End Sub

Private Function pGetWorksheet(ByVal RangeBounds As WorksheetRangeBounds) As Worksheet
    Set pGetWorksheet = pGetWorkbook(RangeBounds).Worksheets(RangeBounds.WorksheetName)
End Function

Private Function pGetRange(ByVal RangeBounds As WorksheetRangeBounds) As Range
    With pGetWorksheet(RangeBounds)
        Set pGetRange = .Range(.Cells(RangeBounds.Row, RangeBounds.Column), .Cells(RangeBounds.FinishRow, RangeBounds.FinishColumn))
    End With
End Function

Private Function pGetCell(ByVal RangeBounds As WorksheetRangeBounds) As Range
    Set pGetCell = pGetWorksheet(RangeBounds).Cells(RangeBounds.Row, RangeBounds.Column)
End Function

Private Sub pRaiseIfDestinationShapeIsInvalid( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long, _
        ByVal ArgumentName As String)

    If DestinationRangeBounds.IsEmpty Then Exit Sub
    If DestinationRangeBounds.IsCell Then Exit Sub
    If DestinationRangeBounds.RowCount = SourceRowCount And DestinationRangeBounds.ColumnCount = SourceColumnCount Then Exit Sub

    Err.Raise vbObjectError + 1, _
            "Class WorksheetService.pRaiseIfDestinationShapeIsInvalid", _
            ArgumentName & " must be an empty range, one cell, or a range with the same row and column counts as the source." & _
            "(destination: " & DestinationRangeBounds.RowCount & "x" & DestinationRangeBounds.ColumnCount & _
            ", source: " & SourceRowCount & "x" & SourceColumnCount & ")"
End Sub

Private Function pGetFormulaNumber(ByVal NumberValue As Double) As String
    Dim result_value As String
    result_value = LTrim$(Str$(NumberValue))
    If Left$(result_value, 1) = "." Then
        result_value = "0" & result_value
    ElseIf Left$(result_value, 2) = "-." Then
        result_value = "-0" & Mid$(result_value, 2)
    End If
    pGetFormulaNumber = result_value
End Function

Private Function pIsHashOnlyString(ByVal TargetString As String) As Boolean
    pIsHashOnlyString = TargetString <> "" And Replace(TargetString, "#", "") = ""
End Function

Private Function pIsDisplayedLiteralString(ByVal TargetCell As Range, ByVal CellText As String) As Boolean
    If VarType(TargetCell.Value) <> vbString Then
        pIsDisplayedLiteralString = False
        Exit Function
    End If

    pIsDisplayedLiteralString = CellText = CStr(TargetCell.Value)
End Function

Private Sub pRaiseReadCellTextColumnWidthError( _
        ByVal TargetCell As Range, _
        ByVal CellText As String, _
        ByVal OriginalErrorNumber As Long, _
        ByVal OriginalErrorSource As String, _
        ByVal OriginalErrorDescription As String)

    Dim reason_text As String
    If TargetCell.Worksheet.ProtectContents Then
        reason_text = "because of sheet protection"
    Else
        reason_text = "because Excel returned an error while changing the column width"
    End If

    Err.Raise _
            Number:=vbObjectError + 1, _
            Source:="Class WorksheetService.pRaiseReadCellTextColumnWidthError", _
            Description:="When reading the cell, the column is too narrow and displays #, but " & reason_text & _
                    ", the column width could not be temporarily expanded." & _
                    "(Sheet=" & TargetCell.Worksheet.Name & _
                    ", Address=" & TargetCell.Address(False, False) & _
                    ", Text=" & CellText & _
                    ", OriginalError=[0x" & Hex(OriginalErrorNumber) & "]@" & OriginalErrorSource & _
                    " " & OriginalErrorDescription & ")"
End Sub

Private Function pContainsItemInStringArray(ByRef TargetArray() As String, ByVal CheckItem As String) As Boolean
    Dim arr_item As Variant
    For Each arr_item In TargetArray
        If arr_item = CheckItem Then
            pContainsItemInStringArray = True
            Exit Function
        End If
    Next arr_item
    pContainsItemInStringArray = False
End Function

Private Sub pCopyCellCore( _
        ByVal SourceRow As Long, ByVal SourceColumn As Long, _
        ByVal DestinationRow As Long, ByVal DestinationColumn As Long, _
        ByVal SourceSheet As Worksheet, ByVal DestinationSheet As Worksheet, _
        ByVal AsValue As Boolean, _
        ByVal CopyNumberFormat As Boolean, _
        Optional ByVal RemainRowCount As Long = 2147483647#, _
        Optional ByVal RemainColumnCount As Long = 2147483647#, _
        Optional ByVal ForcePasteIfArrayFormula As Boolean = False)

    If DestinationSheet Is Nothing Then Set DestinationSheet = SourceSheet

    Dim src_cell As Range
    Set src_cell = SourceSheet.Cells(SourceRow, SourceColumn)

    Dim dst_cell As Range
    Set dst_cell = DestinationSheet.Cells(DestinationRow, DestinationColumn)

        ' Value copy.
    If AsValue Then
        If IsEmpty(src_cell.Value) Then
            dst_cell.ClearContents
        Else
            dst_cell.Value = src_cell.Value
        End If

        If CopyNumberFormat Then dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
        Exit Sub
    End If

        ' Formula copy.
    If src_cell.HasFormula Then
        ' This does not handle relative references. (They remain as in the source.)
        'dst_cell.Formula = src_cell.Formula
        If src_cell.HasArray Then
            If src_cell.Address = src_cell.CurrentArray.Cells(1, 1).Address Or ForcePasteIfArrayFormula Then
                Dim formula_str As String
                ' Temporarily put it into the source cell as a normal formula. (It does not have to calculate correctly.)
                formula_str = src_cell.FormulaArray
                src_cell = formula_str
                ' Put it into the destination cell. (For some reason, extracting from FormulaR1C1 correctly handles relative references.)
                If Application.ReferenceStyle = xlA1 Then
                    dst_cell = Application.ConvertFormula(src_cell.FormulaR1C1, xlR1C1, ToReferenceStyle:=xlA1)
                Else
                    dst_cell = Application.ConvertFormula(src_cell.FormulaR1C1, xlR1C1, ToReferenceStyle:=xlR1C1)
                End If
                ' Restore the source cell to the array formula.
                src_cell.FormulaArray = formula_str

                ' Convert the destination cell's normal formula into an array formula.
                formula_str = dst_cell.Formula
                Dim resized_dst As Range
                Set resized_dst = dst_cell.Resize( _
                        MinLng(src_cell.CurrentArray.Rows.Count, RemainRowCount), _
                        MinLng(src_cell.CurrentArray.Columns.Count, RemainColumnCount))
                resized_dst.FormulaArray = formula_str
                If CopyNumberFormat Then resized_dst.NumberFormatLocal = src_cell.NumberFormatLocal
            Else
                ' Do nothing for cells other than the anchor cell of a multi-cell array.
            End If
        Else
            Dim dst_was_protected As Boolean
            dst_was_protected = dst_cell.Worksheet.ProtectContents

            Dim dst_had_formula As Boolean
            dst_had_formula = dst_cell.HasFormula

            Dim dst_formula_r1c1 As String
            If dst_had_formula Then dst_formula_r1c1 = dst_cell.FormulaR1C1

            Dim dst_value As Variant
            dst_value = dst_cell.Value2

            Dim dst_number_format As String
            dst_number_format = dst_cell.NumberFormatLocal

            Dim dst_wrap_text As Boolean
            dst_wrap_text = dst_cell.WrapText

            On Error GoTo RESTORE_FORMULA_COPY_DESTINATION_ON_ERROR
            dst_cell.NumberFormat = "General"
            dst_cell.FormulaR1C1 = src_cell.FormulaR1C1
            If CopyNumberFormat Then
                dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
            Else
                dst_cell.NumberFormatLocal = dst_number_format
            End If
            On Error GoTo 0
        End If
    Else
        ' Value copy.
        If IsEmpty(src_cell.Value) Then
            dst_cell.ClearContents
        Else
            dst_cell.Value = src_cell.Value
        End If
        If CopyNumberFormat Then dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
    End If
    Exit Sub

RESTORE_FORMULA_COPY_DESTINATION_ON_ERROR:
    Dim err_num As Long
    err_num = Err.Number

    Dim err_source As String
    err_source = Err.Source

    Dim err_desc As String
    err_desc = Err.Description

    Call pRestoreCellContentsAndNumberFormat( _
            dst_cell, _
            dst_was_protected, _
            dst_had_formula, _
            dst_formula_r1c1, _
            dst_value, _
            dst_number_format, _
            dst_wrap_text)

    Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc
End Sub

Private Function pGetValidatedRelativeColumnIndex(ByVal RangeBounds As WorksheetRangeBounds, ByVal ColumnIndex As Variant, ByVal ArgumentName As String) As Long
    Dim column_num As Long
    column_num = pGetValidatedLong(ColumnIndex, ArgumentName)
    If column_num < 1 Or RangeBounds.ColumnCount < column_num Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pGetValidatedRelativeColumnIndex", ArgumentName & " must be a relative column number within the target range."
    End If

    pGetValidatedRelativeColumnIndex = column_num
End Function

Private Function pGetValidatedLong(ByVal TargetValue As Variant, ByVal ArgumentName As String) As Long
    On Error GoTo INVALID_VALUE
    If IsObject(TargetValue) Then GoTo INVALID_VALUE
    If IsError(TargetValue) Then GoTo INVALID_VALUE
    If IsEmpty(TargetValue) Then GoTo INVALID_VALUE
    If IsNull(TargetValue) Then GoTo INVALID_VALUE
    If Not IsNumeric(TargetValue) Then GoTo INVALID_VALUE

    Dim result_value As Long
    result_value = CLng(TargetValue)
    If CDbl(TargetValue) <> CDbl(result_value) Then GoTo INVALID_VALUE

    pGetValidatedLong = result_value
    On Error GoTo 0
    Exit Function

INVALID_VALUE:
    On Error GoTo 0
    Err.Raise vbObjectError + 1, "Class WorksheetService.pGetValidatedLong", ArgumentName & " must be an integer."
End Function

Private Sub pRaiseIfNotEntireRowRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal MethodName As String)
    If RangeBounds.IsEmpty Or Not RangeBounds.IsEntireRow Then
        Err.Raise vbObjectError + 1, "Class WorksheetService.pRaiseIfNotEntireRowRange", MethodName & " requires an entire-row range."
    End If
End Sub
