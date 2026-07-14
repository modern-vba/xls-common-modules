VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetVirtualTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Virtual table that associates multiple WorksheetRangeBounds objects by header and relative row.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pHeaders() As String
Private pColumnRangeList As ObjectList
Private pCount As Long
Private pHeaderCompareMode As VbCompareMethod
Private pHeaderRowCount As Long
Private pNewEnumSnapshot As Collection

'* Row count of the virtual table.
Public Property Get Count() As Long
    Call pCheckInit
    Count = pCount
End Property

'* Row count of the virtual table.
Public Property Get RowCount() As Long
    RowCount = Count
End Property

'* Array copy of resolved header names.
Public Property Get Headers() As String()
    Call pCheckInit
    Headers = pCopyHeaders()
End Property

'* Enumerator for standard For Each.
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Call pCheckInit
    If pNewEnumSnapshot Is Nothing Then Set pNewEnumSnapshot = pCreateNewEnumSnapshot()

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* Indexer for this collection. Read-only. Returns the row at the specified 0-based index (an `ObjectDictionary` with headers as keys and one-row `WorksheetRangeBounds` values). Missing rows in shorter fields are returned as empty ranges with RowCount = 0 while preserving the original field width.
'*
'* @param ItemIndex 0-based index of the row to get.
Public Property Get Item(ByVal ItemIndex As Long) As ObjectDictionary
Attribute Item.VB_UserMemId = 0
    Call pCheckInit
    Call pCheckItemIndex(ItemIndex)

    Dim result_value As ObjectDictionary
    Set result_value = New_ObjectDictionary("WorksheetRangeBounds")
    result_value.CompareMode = pHeaderCompareMode

    Dim col_idx As Long
    For col_idx = 0 To pColumnRangeList.Count - 1
        Dim row_bounds As WorksheetRangeBounds
        Set row_bounds = pGetFieldRowBounds(pColumnRangeList.Item(col_idx), ItemIndex)
        Call result_value.Add(pHeaders(col_idx), row_bounds)
    Next col_idx

    Set Item = result_value
End Property

Private Function pGetFieldRowBounds(ByVal FieldRange As WorksheetRangeBounds, ByVal ItemIndex As Long) As WorksheetRangeBounds
    Dim field_row_index As Long
    field_row_index = ItemIndex + pHeaderRowCount + 1

    If field_row_index <= FieldRange.RowCount Then
        Set pGetFieldRowBounds = FieldRange.GetRow(field_row_index)
        Exit Function
    End If

    Set pGetFieldRowBounds = FieldRange.TransformAbsolute(Row:=FieldRange.Row + field_row_index - 1, FinishRow:=FieldRange.Row + field_row_index - 2)
End Function

Private Sub pCheckItemIndex(ByVal ItemIndex As Long)
    If ItemIndex < 0 Or pCount <= ItemIndex Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pCheckItemIndex", "Row index is out of range. (" & CStr(ItemIndex) & ")"
    End If
End Sub

'* Initializes the virtual table.
'*
'* @param ColumnRangeList `ObjectList` of `WorksheetRangeBounds` objects to treat as virtual table fields.
'* @param HeaderNames [Optional] One-dimensional string array of header names.
'* @param TreatFirstRowAsHeader [Optional] When True, reads the first row of each field range as the header.
'* @param HeaderCompareMode [Optional] Comparison mode for header duplicate checks and row `ObjectDictionary` objects.
'* @param HeaderJoinDelimiter [Optional] Delimiter used when joining header cells from multi-column fields.
'*
'* @details
'* `ColumnRangeList` must have the type contract `ObjectList("WorksheetRangeBounds")`.
'* When `TreatFirstRowAsHeader=True`, `HeaderNames` is ignored.
'* Header count mismatches, duplicate headers, or type mismatches in `HeaderNames` are initialization errors.
Public Sub Initialize( _
        ByVal ColumnRangeList As ObjectList, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.Initialize", "The object is already initialized."
    Call pValidateColumnRangeList(ColumnRangeList)
    If TreatFirstRowAsHeader Then
        pHeaderRowCount = 1
        pHeaders = pReadHeaderNames(ColumnRangeList, HeaderJoinDelimiter)
    Else
        pHeaderRowCount = 0
        If IsMissing(HeaderNames) Then
            Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.Initialize", "HeaderNames is not specified."
        End If
        pHeaders = pGetHeaderNamesArray(HeaderNames)
    End If

    If pHeaderCount() <> ColumnRangeList.Count Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.Initialize", "The number of HeaderNames and ColumnRangeList items does not match."
    End If

    pHeaderCompareMode = HeaderCompareMode
    Call pValidateDuplicateHeaders

    Set pColumnRangeList = ColumnRangeList.CopyList()
    pCount = pCalculateRowCount()
    pIsInitialized = True
End Sub

'* Initializes the virtual table from WorksheetRangeBounds.
'*
'* @param TableRange WorksheetRangeBounds to split into virtual table columns.
'* @param HeaderNames [Optional] One-dimensional string array of header names.
'* @param TreatFirstRowAsHeader [Optional] When True, reads the first row of each field range as the header.
'* @param HeaderCompareMode [Optional] Comparison mode for header duplicate checks and row ObjectDictionary objects.
'* @param HeaderJoinDelimiter [Optional] Delimiter used when joining header cells from multi-column fields.
'*
'* @details
'* Gets a list of column-wise WorksheetRangeBounds objects with TableRange.GetColumns() and delegates to the existing Initialize.
Public Sub InitializeFromRangeBounds( _
        ByVal TableRange As WorksheetRangeBounds, _
        Optional ByVal HeaderNames As Variant, _
        Optional ByVal TreatFirstRowAsHeader As Boolean = False, _
        Optional ByVal HeaderCompareMode As VbCompareMethod = vbBinaryCompare, _
        Optional ByVal HeaderJoinDelimiter As String = "")

    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.InitializeFromRangeBounds", "The object is already initialized."
    If TableRange Is Nothing Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.InitializeFromRangeBounds", "TableRange is not specified."

    Call Initialize( _
            ColumnRangeList:=TableRange.GetColumns(), _
            HeaderNames:=HeaderNames, _
            TreatFirstRowAsHeader:=TreatFirstRowAsHeader, _
            HeaderCompareMode:=HeaderCompareMode, _
            HeaderJoinDelimiter:=HeaderJoinDelimiter)
End Sub

Private Sub pValidateColumnRangeList(ByVal ColumnRangeList As ObjectList)
    If ColumnRangeList Is Nothing Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateColumnRangeList", "ColumnRangeList is not specified."
    End If
    If Not ColumnRangeList.HasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateColumnRangeList", "ColumnRangeList requires a WorksheetRangeBounds type contract."
    End If
    If StrComp(ColumnRangeList.ElementTypeName, "WorksheetRangeBounds", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateColumnRangeList", "The ColumnRangeList element type contract must specify WorksheetRangeBounds."
    End If
End Sub

Private Function pGetHeaderNamesArray(ByVal HeaderNames As Variant) As String()
    Call pValidateHeaderNamesType(HeaderNames)

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, HeaderNames)
    If LBound(lbound_arr) <> UBound(lbound_arr) Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pGetHeaderNamesArray", "HeaderNames must be a one-dimensional string array."
    End If

    If IsEmptyArray(HeaderNames) Then
        pGetHeaderNamesArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To ubound_arr(LBound(ubound_arr)) - lbound_arr(LBound(lbound_arr)))

    Dim src_idx As Long
    Dim dst_idx As Long
    dst_idx = 0
    For src_idx = lbound_arr(LBound(lbound_arr)) To ubound_arr(LBound(ubound_arr))
        result_value(dst_idx) = CStr(HeaderNames(src_idx))
        dst_idx = dst_idx + 1
    Next src_idx

    pGetHeaderNamesArray = result_value
End Function

Private Function pReadHeaderNames(ByVal ColumnRangeList As ObjectList, ByVal HeaderJoinDelimiter As String) As String()
    If ColumnRangeList.Count = 0 Then
        pReadHeaderNames = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(0 To ColumnRangeList.Count - 1)

    Dim col_idx As Long
    For col_idx = 0 To ColumnRangeList.Count - 1
        result_value(col_idx) = pBuildHeaderName(ColumnRangeList.Item(col_idx), HeaderJoinDelimiter)
    Next col_idx

    pReadHeaderNames = result_value
End Function

Private Sub pValidateDuplicateHeaders()
    Dim header_dict As Object
    Set header_dict = CreateObject("Scripting.Dictionary")
    header_dict.CompareMode = pHeaderCompareMode

    Dim header_idx As Long
    For header_idx = LBound(pHeaders) To UBound(pHeaders)
        If header_dict.Exists(pHeaders(header_idx)) Then
            Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateDuplicateHeaders", "Header is duplicated. (" & pHeaders(header_idx) & ")"
        End If
        Call header_dict.Add(pHeaders(header_idx), True)
    Next header_idx
End Sub

Private Function pCalculateRowCount() As Long
    Dim result_value As Long
    result_value = 0

    Dim col_idx As Long
    For col_idx = 0 To pColumnRangeList.Count - 1
        Dim field_range As WorksheetRangeBounds
        Set field_range = pColumnRangeList.Item(col_idx)

        Dim field_row_count As Long
        field_row_count = field_range.RowCount - pHeaderRowCount
        If field_row_count < 0 Then field_row_count = 0
        If result_value < field_row_count Then result_value = field_row_count
    Next col_idx

    pCalculateRowCount = result_value
End Function

Private Function pHeaderCount() As Long
    If IsEmptyArray(pHeaders) Then
        pHeaderCount = 0
    Else
        pHeaderCount = UBound(pHeaders) - LBound(pHeaders) + 1
    End If
End Function

'* Returns the row at the specified row index.
'*
'* @param RowIndex Row index. Positive numbers are 1-origin from the top; negative numbers are 1-origin from the bottom. 0 is an out-of-range error.
'* @return `ObjectDictionary` with headers as keys and one-row `WorksheetRangeBounds` values.
'*
'* @details
'* Gets a row using the same row index rules as `WorksheetRangeBounds.GetRow`.
Public Function GetRow(ByVal RowIndex As Long) As ObjectDictionary
    Call pCheckInit
    If RowIndex = 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.GetRow", "Row index is out of range. (0)"
    End If

    Dim item_idx As Long
    If 0 < RowIndex Then
        item_idx = RowIndex - 1
    Else
        item_idx = pCount + RowIndex
    End If

    Set GetRow = Item(item_idx)
End Function

'* Returns the enumerator for this virtual table.
'*
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @return Enumerator that returns row `ObjectDictionary` objects.
'*
'* @details
'* The returned enumerator is read-only. `IEnumerator.Index` is 0-based.
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False) As IEnumerator
    Call pCheckInit

    Dim result_value As WorksheetVirtualTableEnumerator
    Set result_value = New WorksheetVirtualTableEnumerator
    Call result_value.Initialize(Me, Descending:=Descending)

    Set GetEnumerator = result_value
End Function

Private Sub pValidateHeaderNamesType(ByVal HeaderNames As Variant)
    If Not IsArray(HeaderNames) Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateHeaderNamesType", "HeaderNames must be a one-dimensional string array."
    End If

    Dim header_var_type As Long
    header_var_type = VarType(HeaderNames)
    If (header_var_type And vbArray) = 0 Or header_var_type - vbArray <> vbString Then
        Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pValidateHeaderNamesType", "HeaderNames must be a one-dimensional string array."
    End If
End Sub

Private Function pBuildHeaderName(ByVal FieldRange As WorksheetRangeBounds, ByVal HeaderJoinDelimiter As String) As String
    If FieldRange.RowCount = 0 Then
        pBuildHeaderName = ""
        Exit Function
    End If

    Dim header_row As WorksheetRangeBounds
    Set header_row = FieldRange.GetRow(1)

    Dim header_parts() As String
    ReDim header_parts(0 To header_row.ColumnCount - 1)

    Dim col_idx As Long
    For col_idx = 1 To header_row.ColumnCount
        header_parts(col_idx - 1) = WsSrv.ReadCell(header_row.GetCell(1, col_idx))
    Next col_idx

    pBuildHeaderName = Join(header_parts, HeaderJoinDelimiter)
End Function

Private Function pCopyHeaders() As String()
    If pHeaderCount() = 0 Then
        pCopyHeaders = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    ReDim result_value(LBound(pHeaders) To UBound(pHeaders))

    Dim item_idx As Long
    For item_idx = LBound(pHeaders) To UBound(pHeaders)
        result_value(item_idx) = pHeaders(item_idx)
    Next item_idx

    pCopyHeaders = result_value
End Function

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator()
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class WorksheetVirtualTable.pCheckInit", "The object is not initialized."
End Sub
