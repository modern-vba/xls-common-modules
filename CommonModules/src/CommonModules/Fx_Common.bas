Attribute VB_Name = "Fx_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Standard module that groups UDFs exposed by CommonModules as Excel worksheet functions.
'!
' #############################################################################

'* Compares two strings and gets the difference string.
'*
'* @param Expression1 Original string.
'* @param Expression2 Modified string.
'* @param ExtractType Extraction type. Specify 0 to get the string deleted from the original string. Specify 1 to get the string added to the original string.
'* @return Difference string.
'*
'* @details
'* Compares two strings and returns a string made by joining the differing characters in order from the front.
'* For example, if Expression1 is "aaa [bbb ccc] ddd", Expression2 is "aaa bbb ddd eee", and ExtractType is 0, this returns "[ccc] ".
Public Function DIFFSTR( _
    ByVal Expression1 As String, _
    ByVal Expression2 As String, _
    Optional ByVal ExtractType As Integer = 0) As Variant  ' String

    Call InitializeUdfCommonService

    '=== 1) Simple handling for cases that include empty strings ==================
    If Expression1 = "" Or Expression2 = "" Then
        If Expression1 = "" Then
            If ExtractType = 0 Then
                DIFFSTR = ""
            Else
                DIFFSTR = Expression2
            End If
        Else
            ' Expression2 is an empty string
            If ExtractType = 0 Then
                DIFFSTR = Expression1
            Else
                DIFFSTR = ""
            End If
        End If
        Exit Function
    End If

    '=== 2) Convert strings to character arrays ==============================
    Dim expr_1_arr() As String
    Dim expr_2_arr() As String
    expr_1_arr = ConvertStringToCharArray(Expression1)
    expr_2_arr = ConvertStringToCharArray(Expression2)

    '=== 3) Use DiffStringArray to calculate the differences ===============
    Dim diff_type_arr() As String
    Call DiffStringArray(expr_1_arr, expr_2_arr, diff_type_arr, EnableReplaceType:=False)

    '=== 4) Rebuild the difference string according to the extraction type ==============
    Dim result_str As String
    result_str = ""

    Dim i_idx As Long
    Dim last_idx As Long
    last_idx = UBound(expr_1_arr)  ' = UBound(expr_2_arr) = UBound(diff_type_arr)

    For i_idx = LBound(expr_1_arr) To last_idx
        Dim diff_tag As String
        diff_tag = diff_type_arr(i_idx)

        Select Case ExtractType
            Case 0
                ' Collect characters removed from the old string by deletions
                If (diff_tag = "DEL") Then
                    result_str = result_str & expr_1_arr(i_idx)
                End If
            Case 1
                ' Collect characters added to the new string by additions
                If (diff_tag = "ADD") Then
                    result_str = result_str & expr_2_arr(i_idx)
                End If
        End Select
    Next i_idx

    DIFFSTR = result_str
End Function

''* Custom implementation of the TEXTSPLIT worksheet function available in newer Excel versions.
''*
''* @param Expression Text to split.
''* @param ColumnDelimiter Column delimiter. (In a typical CSV, this is a comma.)
''* @param RowDelimiter Row delimiter. (In a typical CSV, this is a line break.) (Optional)
''* @param IgnoreEmpty Whether to ignore consecutive delimiters. Specify True to ignore them. The default is False, and values are filled with PadWith.
''* @param MatchMode (Unimplemented feature.)
''* @param PadWith Value to use for consecutive delimiters. The default is #N/A.
''* @return Array containing the split result.
''*
''* @details
''* Implementation for Excel versions that do not have TEXTSPLIT. For details, see Microsoft's TEXTSPLIT documentation.
''*
''* @see https://support.microsoft.com/ja-jp/office/textsplit-%E9%96%A2%E6%95%B0-b1ca414e-4c21-4ca0-b1b7-bdecace8a6e7
'Function TEXTSPLIT( _
'        ByVal Expression As String, ByVal ColumnDelimiter As String, Optional ByVal RowDelimiter As String, _
'        Optional ByVal IgnoreEmpty As Boolean = False, Optional ByVal MatchMode As Integer = 0, Optional ByVal PadWith As Variant = Nothing) As Variant()
'
'    If MatchMode <> 0 Then
'        ' MatchMode 1 (IgnoreCase) is not implemented
'        TEXTSPLIT = CVErr(xlErrValue)
'        Exit Function
'    End If
'
'    If IsObject(PadWith) Then
'        If PadWith Is Nothing Then
'            PadWith = CVErr(xlErrNA)
'        End If
'    End If
'
'    Dim rows_arr() As String
'    rows_arr = Split(Expression, RowDelimiter)
'
'    Dim max_col As Long
'    Dim rows_list As ObjectList
'    Set rows_list = New ObjectList
'    Dim row_item As Variant 'String
'    For Each row_item In rows_arr
'        If row_item <> "" Or Not IgnoreEmpty Then
'            Dim cols_arr() As String
'            cols_arr = Split(row_item, ColumnDelimiter)
'
'            Dim cols_list As ObjectList
'            Set cols_list = New ObjectList
'
'            Dim col_item As Variant 'String
'            For Each col_item In cols_arr
'                If col_item <> "" Or Not IgnoreEmpty Then
'                    Call cols_list.Add(col_item)
'                End If
'            Next col_item
'
'            If 0 < cols_list.Count Or Not IgnoreEmpty Then
'                Call rows_list.Add(cols_list)
'                max_col = MaxLng(max_col, cols_list.Count)
'            End If
'        End If
'    Next row_item
'
'    Dim result_value() As Variant
'    ReDim result_value(0 To rows_list.Count - 1, 0 To max_col - 1)
'
'    Dim row_idx As Long
'    Dim col_idx As Long
'    For row_idx = 0 To UBound(result_value, 1) Step 1
'        For col_idx = 0 To UBound(result_value, 2) Step 1
'            Set cols_list = rows_list.Item(row_idx)
'            If col_idx < cols_list.Count Then
'                result_value(row_idx, col_idx) = cols_list.Item(col_idx)
'            Else
'                result_value(row_idx, col_idx) = PadWith
'            End If
'        Next col_idx
'    Next row_idx
'
'    TEXTSPLIT = result_value
'End Function

''* Custom implementation of the TEXTJOIN worksheet function available in newer Excel versions.
''*
''* @param Delimiter Delimiter.
''* @param IgnoreEmpty Whether to ignore empty strings. When True, empty strings are ignored. When False, they are not ignored and delimiters are consecutive.
''* @param Expression1 First string.
''* @param Expressions Second and later strings. (Variable-length argument.)
''*
''* @return Joined string.
''*
''* @details
''* Implementation for Excel versions that do not have TEXTJOIN. For details, see Microsoft's TEXTJOIN documentation.
''*
''* @see https://support.microsoft.com/ja-jp/office/textjoin-%E9%96%A2%E6%95%B0-357b449a-ec91-49d0-80c3-0e8fc845691c
'Function TEXTJOIN(ByVal Delimiter As String, ByVal IgnoreEmpty As Boolean, ByVal Expression1 As Variant, ParamArray Expressions() As Variant) As Variant
'    Dim result_value As String
'    Dim is_first As Boolean
'
'    is_first = True
'
'    On Error Resume Next
'    Call pTextJoinCore(result_value, is_first, Delimiter, Expression1, IgnoreEmpty)
'    If Err.Number <> 0 Then
'        TEXTJOIN = CVErr(xlErrNA)
'        Exit Function
'    End If
'
'    Dim param_item As Variant
'    For Each param_item In Expressions
'        Call pTextJoinCore(result_value, is_first, Delimiter, param_item, IgnoreEmpty)
'        If Err.Number <> 0 Then
'            TEXTJOIN = CVErr(xlErrNA)
'            Exit Function
'        End If
'    Next param_item
'
'    TEXTJOIN = result_value
'End Function

'Sub pTextJoinCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If TypeOf Expression Is Range Or IsArray(Expression) Then
'        ' When Expression is a Range or array
'        Dim cell_item As Range
'        For Each cell_item In Expression
'            Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, cell_item.Value, IgnoreEmpty)
'        Next cell_item
''    ElseIf IsArray(Expression) Then
''        ' When Expression is an array
''        If LBound(Expression) <= UBound(Expression) Then
''            Dim item_idx As Long
''            For item_idx = LBound(Expression) To UBound(Expression)
''                Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression(item_idx), IgnoreEmpty)
''            Next item_idx
''        ElseIf Not IgnoreEmpty Then
''            ResultString = ResultString & Delimiter
''        End If
'    Else
'        ' When Expression is neither a Range nor an array
'        Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression, IgnoreEmpty)
'    End If
'End Sub

'Private Sub pTextJoinItemCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If IsError(Expression) Then Err.Raise xlErrNA, "TEXTJOIN", "The values to join include an error value."
'    If Expression <> "" Or Not IgnoreEmpty Then
'        If IsFirst Then
'            ResultString = Expression
'            IsFirst = False
'        Else
'            ResultString = ResultString & Delimiter & Expression
'        End If
'    End If
'End Sub
