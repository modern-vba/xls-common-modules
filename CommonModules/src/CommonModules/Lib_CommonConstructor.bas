Attribute VB_Name = "Lib_CommonConstructor"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Standard module that groups functions equivalent to constructors for common class modules.
'!
' #############################################################################

'* Creates a new WorksheetRangeBounds instance.
'*
'* @param RowIndex [Optional] Row number.
'* @param ColumnIndex [Optional] Column number.
'* @param FinishRowIndex [Optional] Last row number.
'* @param FinishColumnIndex [Optional] Last column number.
'* @param SheetName [Optional] Worksheet name.
'* @param BookName [Optional] Workbook name.
'* @return New WorksheetRangeBounds instance.
'*
'* @details
'* Creates a new WorksheetRangeBounds instance.
'*
'* Usage example:
'* @code
'* ' One cell (B5).
'* Dim cell_bounds As WorksheetRangeBounds
'* Set cell_bounds = New_RangeBounds(Row:=5, Column:=2)
'*
'* ' Entire row (5:5).
'* Dim row_bounds As WorksheetRangeBounds
'* Set row_bounds = New_RangeBounds(Row:=5)
'*
'* ' Two entire columns (B:C).
'* Dim two_columns_bounds As WorksheetRangeBounds
'* Set two_columns_bounds = New_RangeBounds(Column:=2, FinishColumn:=3)
'*
'* ' Entire table (B2:E20).
'* Dim table_bounds As WorksheetRangeBounds
'* Set table_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=20, FinishColumn:=5)
'* @endcode
Public Function New_RangeBounds( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Sheet As String = "Sheet1", _
        Optional ByVal Book As String = "") As WorksheetRangeBounds

    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds

    Call result_value.Initialize( _
            Row:=Row, _
            Column:=Column, _
            FinishRow:=FinishRow, _
            FinishColumn:=FinishColumn, _
            Sheet:=Sheet, _
            Book:=Book)

    Set New_RangeBounds = result_value
End Function

'* Creates a new ObjectList instance and sets its element type contract.
'*
'* @param ElementTypeName Element type contract name.
'* @param RequireComparable Whether elements must implement IComparable.
'* @param ObjectKeyMode Object identity / duplicate check mode.
'* @return Initialized ObjectList instance.
'*
'* @details
'* New_ processing remains a thin factory and delegates the actual processing to ObjectList.Initialize.
Public Function New_ObjectList( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectList

    Dim result_value As ObjectList
    Set result_value = New ObjectList
    Call result_value.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)

    Set New_ObjectList = result_value
End Function

'* Creates a new ObjectSet instance and sets its element type contract.
'*
'* @param ElementTypeName Element type contract name.
'* @param RequireComparable Whether elements must implement IComparable.
'* @param ObjectKeyMode Object identity / duplicate check mode.
'* @return Initialized ObjectSet instance.
'*
'* @details
'* New_ processing remains a thin factory and delegates the actual processing to ObjectSet.Initialize.
Public Function New_ObjectSet( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectSet

    Dim result_value As ObjectSet
    Set result_value = New ObjectSet
    Call result_value.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)

    Set New_ObjectSet = result_value
End Function

'* Creates a new ObjectDictionary instance and sets its element type contract as needed.
'*
'* @param ElementTypeName [Optional] Element type contract name.
'* @param RequireComparable Whether elements must implement IComparable.
'* @param ObjectKeyMode Object identity / duplicate check mode.
'* @return New ObjectDictionary instance.
'*
'* @details
'* When ElementTypeName is omitted, the same type inference as ObjectList is performed from the first added element.
Public Function New_ObjectDictionary( _
        Optional ByVal ElementTypeName As String = "", _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE) As ObjectDictionary

    Dim result_value As ObjectDictionary
    Set result_value = New ObjectDictionary
    If ElementTypeName <> "" Then
        Call result_value.Initialize( _
                ElementTypeName:=ElementTypeName, _
                RequireComparable:=RequireComparable, _
                ObjectKeyMode:=ObjectKeyMode)
    End If

    Set New_ObjectDictionary = result_value
End Function

'* Creates a new WorksheetRangeBounds instance from an Excel address string.
'*
'* @param AddressString Excel address string to use for initialization.
'* @return New WorksheetRangeBounds instance.
'*
'* @details
'* New_ processing remains a thin factory and delegates the actual processing to WorksheetRangeBounds.InitializeFromAddress.
Public Function New_RangeBoundsFromAddress(ByVal AddressString As String) As WorksheetRangeBounds
    Dim result_value As WorksheetRangeBounds
    Set result_value = New WorksheetRangeBounds
    Call result_value.InitializeFromAddress(AddressString)

    Set New_RangeBoundsFromAddress = result_value
End Function

'* Creates a new WorksheetVirtualTable instance.
'*
'* @param ColumnRangeList List of WorksheetRangeBounds objects to treat as virtual table columns.
'* @param HeaderNames [Optional] String array of header names.
'* @param TreatFirstRowAsHeader When True, reads the first row of each range as headers.
'* @param HeaderCompareMode Comparison mode for header duplicate checks and row Dictionary objects.
'* @param HeaderJoinDelimiter Delimiter used when joining header cells from multi-column ranges.
'* @return Initialized WorksheetVirtualTable instance.
'*
'* @details
'* New_ processing remains a thin factory and delegates the actual processing to WorksheetVirtualTable.Initialize.
Public Function New_WorksheetVirtualTable( _
        ByVal ColumnRangeList As ObjectList, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "") As WorksheetVirtualTable

    Dim result_value As WorksheetVirtualTable
    Set result_value = New WorksheetVirtualTable
    Call result_value.Initialize( _
            ColumnRangeList:=ColumnRangeList, _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)

    Set New_WorksheetVirtualTable = result_value
End Function

'* Creates a new WorksheetVirtualTable instance from WorksheetRangeBounds.
'*
'* @param TableRange WorksheetRangeBounds to split into virtual table columns.
'* @param HeaderNames [Optional] String array of header names.
'* @param TreatFirstRowAsHeader When True, reads the first row of each range as headers.
'* @param HeaderCompareMode Comparison mode for header duplicate checks and row Dictionary objects.
'* @param HeaderJoinDelimiter Delimiter used when joining header cells from multi-column ranges.
'* @return Initialized WorksheetVirtualTable instance.
'*
'* @details
'* New_ processing remains a thin factory and delegates the actual processing to WorksheetVirtualTable.InitializeFromRangeBounds.
Public Function New_WorksheetVirtualTableFromRangeBounds( _
        ByVal TableRange As WorksheetRangeBounds, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "") As WorksheetVirtualTable

    Dim result_value As WorksheetVirtualTable
    Set result_value = New WorksheetVirtualTable
    Call result_value.InitializeFromRangeBounds( _
            TableRange:=TableRange, _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)

    Set New_WorksheetVirtualTableFromRangeBounds = result_value
End Function
