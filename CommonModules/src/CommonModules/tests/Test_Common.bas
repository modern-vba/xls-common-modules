Attribute VB_Name = "Test_Lib_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for the Lib_Common module.
'!
' #############################################################################

' -----------------------------------------------------------------------------
' DIFFSTR
' -----------------------------------------------------------------------------

Public Sub Test_DIFFSTR_RemoveAndAdd_ReturnsExpectedCharacters(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "b", DIFFSTR("abc", "ac", 0)
    Assert.Equals "b", DIFFSTR("ac", "abc", 1)
    Assert.Equals "", DIFFSTR("abc", "abc", 0)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_DIFFSTR_NotInitialized_InitializesUdfCommonServicesOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing
    Set WsSrv = Nothing
    Set FsSrv = Nothing
    Set TfSrv = Nothing

    ' Act
    Dim actual_value As Variant
    actual_value = DIFFSTR("", "abc", 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "abc", CStr(actual_value)
    Assert.IsTypeOf "WorkbookService", WbSrv
    Assert.IsTypeOf "WorksheetService", WsSrv
    Assert.IsNothing FsSrv
    Assert.IsNothing TfSrv
End Sub

' -----------------------------------------------------------------------------
' InitializeCommonService
' -----------------------------------------------------------------------------

Public Sub Test_InitializeCommonService_NotInitialized_InitializesCommonServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing
    Set WsSrv = Nothing
    Set FsSrv = Nothing
    Set TfSrv = Nothing

    ' Act
    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "WorkbookService", WbSrv
    Assert.IsTypeOf "WorksheetService", WsSrv
    Assert.IsTypeOf "FileSystemService", FsSrv
    Assert.IsTypeOf "TextFileService", TfSrv
End Sub

' -----------------------------------------------------------------------------
' New_RangeBounds
' -----------------------------------------------------------------------------

Public Sub Test_NewRangeBounds_Parameters_ReturnsInitializedRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As WorksheetRangeBounds
    Set actual_value = New_RangeBounds( _
            Row:=2, _
            Column:=3, _
            FinishRow:=4, _
            FinishColumn:=5, _
            Sheet:="Data", _
            Book:="Book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_value.Row
    Assert.EqualsNumeric 3, actual_value.Column
    Assert.EqualsNumeric 4, actual_value.FinishRow
    Assert.EqualsNumeric 5, actual_value.FinishColumn
    Assert.Equals "Data", actual_value.WorksheetName
    Assert.Equals "Book.xlsm", actual_value.WorkbookName
End Sub

' -----------------------------------------------------------------------------
' New_RangeBoundsFromAddress
' -----------------------------------------------------------------------------

Public Sub Test_NewRangeBoundsFromAddress_Address_ReturnsInitializedRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As WorksheetRangeBounds
    Set actual_value = New_RangeBoundsFromAddress("'[Book.xlsm]Data'!$B$2:$C$4")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_value.Row
    Assert.EqualsNumeric 2, actual_value.Column
    Assert.EqualsNumeric 4, actual_value.FinishRow
    Assert.EqualsNumeric 3, actual_value.FinishColumn
    Assert.Equals "Data", actual_value.WorksheetName
    Assert.Equals "Book.xlsm", actual_value.WorkbookName
End Sub


' -----------------------------------------------------------------------------
' ExpandRangeBoundsToMax
' -----------------------------------------------------------------------------

Public Sub Test_ExpandRangeBoundsToMax_ExpandRows_ExtendsShortRangesToMaxRowCount(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim short_bounds As WorksheetRangeBounds
    Set short_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=3, FinishColumn:=3, Sheet:="Data", Book:="BookA.xlsm")

    Dim long_bounds As WorksheetRangeBounds
    Set long_bounds = New_RangeBounds(Row:=10, Column:=5, FinishRow:=14, FinishColumn:=5, Sheet:="Data", Book:="BookA.xlsm")

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, False, short_bounds, long_bounds)

    Dim actual_short As WorksheetRangeBounds
    Dim actual_long As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_short = actual_list.Item(0)
        Set actual_long = actual_list.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.EqualsNumeric 2, actual_short.Row
    Assert.EqualsNumeric 3, actual_short.Column
    Assert.EqualsNumeric 6, actual_short.FinishRow
    Assert.EqualsNumeric 3, actual_short.FinishColumn
    Assert.Equals "Data", actual_short.WorksheetName
    Assert.Equals "BookA.xlsm", actual_short.WorkbookName
    Assert.EqualsNumeric 10, actual_long.Row
    Assert.EqualsNumeric 5, actual_long.Column
    Assert.EqualsNumeric 14, actual_long.FinishRow
    Assert.EqualsNumeric 5, actual_long.FinishColumn
End Sub

Public Sub Test_ExpandRangeBoundsToMax_ExpandColumns_ExtendsShortRangesToMaxColumnCount(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim narrow_bounds As WorksheetRangeBounds
    Set narrow_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=2, FinishColumn:=4, Sheet:="Data", Book:="BookA.xlsm")

    Dim wide_bounds As WorksheetRangeBounds
    Set wide_bounds = New_RangeBounds(Row:=8, Column:=10, FinishRow:=8, FinishColumn:=14, Sheet:="Data", Book:="BookA.xlsm")

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(False, True, narrow_bounds, wide_bounds)

    Dim actual_narrow As WorksheetRangeBounds
    Dim actual_wide As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_narrow = actual_list.Item(0)
        Set actual_wide = actual_list.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.EqualsNumeric 2, actual_narrow.Row
    Assert.EqualsNumeric 3, actual_narrow.Column
    Assert.EqualsNumeric 2, actual_narrow.FinishRow
    Assert.EqualsNumeric 7, actual_narrow.FinishColumn
    Assert.EqualsNumeric 8, actual_wide.Row
    Assert.EqualsNumeric 10, actual_wide.Column
    Assert.EqualsNumeric 8, actual_wide.FinishRow
    Assert.EqualsNumeric 14, actual_wide.FinishColumn
End Sub

Public Sub Test_ExpandRangeBoundsToMax_ExpandRowsColumns_PreservesEachSheetAndBook(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim small_bounds As WorksheetRangeBounds
    Set small_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=3, FinishColumn:=4, Sheet:="DataA", Book:="BookA.xlsm")

    Dim large_bounds As WorksheetRangeBounds
    Set large_bounds = New_RangeBounds(Row:=20, Column:=8, FinishRow:=24, FinishColumn:=12, Sheet:="DataB", Book:="BookB.xlsm")

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, True, small_bounds, large_bounds)

    Dim actual_small As WorksheetRangeBounds
    Dim actual_large As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_small = actual_list.Item(0)
        Set actual_large = actual_list.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.EqualsNumeric 2, actual_small.Row
    Assert.EqualsNumeric 3, actual_small.Column
    Assert.EqualsNumeric 6, actual_small.FinishRow
    Assert.EqualsNumeric 7, actual_small.FinishColumn
    Assert.Equals "DataA", actual_small.WorksheetName
    Assert.Equals "BookA.xlsm", actual_small.WorkbookName
    Assert.EqualsNumeric 20, actual_large.Row
    Assert.EqualsNumeric 8, actual_large.Column
    Assert.EqualsNumeric 24, actual_large.FinishRow
    Assert.EqualsNumeric 12, actual_large.FinishColumn
    Assert.Equals "DataB", actual_large.WorksheetName
    Assert.Equals "BookB.xlsm", actual_large.WorkbookName
End Sub

Public Sub Test_ExpandRangeBoundsToMax_NoExpand_PreservesOriginalShape(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim first_bounds As WorksheetRangeBounds
    Set first_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="DataA", Book:="BookA.xlsm")

    Dim second_bounds As WorksheetRangeBounds
    Set second_bounds = New_RangeBounds(Row:=10, Column:=7, FinishRow:=15, FinishColumn:=10, Sheet:="DataB", Book:="BookB.xlsm")

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(False, False, first_bounds, second_bounds)

    Dim actual_first As WorksheetRangeBounds
    Dim actual_second As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_first = actual_list.Item(0)
        Set actual_second = actual_list.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.IsTrue first_bounds.Equals(actual_first)
    Assert.IsTrue second_bounds.Equals(actual_second)
End Sub

Public Sub Test_ExpandRangeBoundsToMax_NoInput_ReturnsEmptyTypedList(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_list.Count
    Assert.IsTrue actual_list.HasItemTypeContract
    Assert.Equals "WorksheetRangeBounds", actual_list.ElementTypeName
End Sub

Public Sub Test_ExpandRangeBoundsToMax_EmptyRange_ReformsUsingPreservedShape(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim empty_bounds As WorksheetRangeBounds
    Set empty_bounds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=4, FinishColumn:=6, Sheet:="DataA", Book:="BookA.xlsm")

    Dim full_bounds As WorksheetRangeBounds
    Set full_bounds = New_RangeBounds(Row:=10, Column:=10, FinishRow:=12, FinishColumn:=11, Sheet:="DataB", Book:="BookB.xlsm")

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, True, empty_bounds, full_bounds)

    Dim actual_empty As WorksheetRangeBounds
    Dim actual_full As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_empty = actual_list.Item(0)
        Set actual_full = actual_list.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.IsFalse actual_empty.IsEmpty
    Assert.EqualsNumeric 5, actual_empty.Row
    Assert.EqualsNumeric 2, actual_empty.Column
    Assert.EqualsNumeric 7, actual_empty.FinishRow
    Assert.EqualsNumeric 6, actual_empty.FinishColumn
    Assert.EqualsNumeric 3, actual_empty.RowCount
    Assert.EqualsNumeric 5, actual_empty.ColumnCount
    Assert.EqualsNumeric 10, actual_full.Row
    Assert.EqualsNumeric 10, actual_full.Column
    Assert.EqualsNumeric 12, actual_full.FinishRow
    Assert.EqualsNumeric 14, actual_full.FinishColumn
End Sub

Public Sub Test_ExpandRangeBoundsToMax_InvalidArgument_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, True, "invalid")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExpandRangeBoundsToMax", Err.Source
    Assert.IsTrue 0 < InStr(1, Err.Description, "WorksheetRangeBounds", vbBinaryCompare)
End Sub

Public Sub Test_ExpandRangeBoundsToMax_NothingArgument_RaisesArgumentError(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim nothing_bounds As WorksheetRangeBounds

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, True, nothing_bounds)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExpandRangeBoundsToMax", Err.Source
    Assert.IsTrue 0 < InStr(1, Err.Description, "Nothing", vbBinaryCompare)
End Sub

Public Sub Test_ExpandRangeBoundsToMax_OneDirectionEmptyRange_MatchesTransform(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    Call InitializeCommonService

    ' Arrange
    Dim empty_bounds As WorksheetRangeBounds
    Set empty_bounds = New_RangeBounds(Row:=5, Column:=2, FinishRow:=0, FinishColumn:=0, Sheet:="DataA", Book:="BookA.xlsm")

    Dim tall_bounds As WorksheetRangeBounds
    Set tall_bounds = New_RangeBounds(Row:=10, Column:=4, FinishRow:=12, FinishColumn:=4, Sheet:="DataB", Book:="BookB.xlsm")

    Dim expected_empty As WorksheetRangeBounds
    Set expected_empty = empty_bounds.Transform(AddRow:=tall_bounds.RowCount - empty_bounds.RowCount, AddColumn:=0)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ExpandRangeBoundsToMax(True, False, empty_bounds, tall_bounds)

    Dim actual_empty As WorksheetRangeBounds
    If Err.Number = 0 Then
        Set actual_empty = actual_list.Item(0)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue expected_empty.Equals(actual_empty)
    Assert.IsTrue actual_empty.IsEmpty
End Sub
' -----------------------------------------------------------------------------
' HandleError
' -----------------------------------------------------------------------------

Public Sub Test_HandleError_RaisedError_CapturesAndClearsError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim actual_number As Long
    Dim actual_source As String
    Dim actual_description As String

    ' Act
    Err.Raise 1001, "TestSource", "TestDescription"
    Call HandleError(actual_number, actual_source, actual_description, "supplement")

    ' Assert
    Assert.EqualsNumeric 1001, actual_number
    Assert.Equals "TestSource", actual_source
    Assert.Equals "TestDescription [ supplement ]", actual_description
    Assert.EqualsNumeric 0, Err.Number
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' ExcelErrorToString
' -----------------------------------------------------------------------------

Public Sub Test_ExcelErrorToString_AllExcelErrorValues_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim error_values As Variant
    error_values = Array( _
            CVErr(xlErrDiv0), _
            CVErr(xlErrNA), _
            CVErr(xlErrName), _
            CVErr(xlErrNull), _
            CVErr(xlErrNum), _
            CVErr(xlErrRef), _
            CVErr(xlErrValue), _
            CVErr(xlErrGettingData), _
            CVErr(xlErrSpill), _
            CVErr(xlErrConnect), _
            CVErr(xlErrBlocked), _
            CVErr(xlErrUnknown), _
            CVErr(xlErrField), _
            CVErr(xlErrCalc))

    Dim error_texts As Variant
    error_texts = Array( _
            "#DIV/0!", _
            "#N/A", _
            "#NAME?", _
            "#NULL!", _
            "#NUM!", _
            "#REF!", _
            "#VALUE!", _
            "#GETTING_DATA", _
            "#SPILL!", _
            "#CONNECT!", _
            "#BLOCKED!", _
            "#UNKNOWN!", _
            "#FIELD!", _
            "#CALC!")

    ' Act / Assert
    Dim error_idx As Long
    For error_idx = LBound(error_values) To UBound(error_values)
        Assert.Equals CStr(error_texts(error_idx)), ExcelErrorToString(error_values(error_idx))
    Next error_idx
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ExcelErrorToString_NonErrorValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelErrorToString("#N/A")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelErrorToString", Err.Source
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' TryConvertExcelErrorStringToCVErr
' -----------------------------------------------------------------------------

Public Sub Test_TryConvertExcelErrorStringToCVErr_AllExcelErrorTexts_ReturnsErrorValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim error_texts As Variant
    error_texts = Array( _
            "#DIV/0!", _
            "#N/A", _
            "#NAME?", _
            "#NULL!", _
            "#NUM!", _
            "#REF!", _
            "#VALUE!", _
            "#GETTING_DATA", _
            "#SPILL!", _
            "#CONNECT!", _
            "#BLOCKED!", _
            "#UNKNOWN!", _
            "#FIELD!", _
            "#CALC!")

    Dim error_values As Variant
    error_values = Array( _
            CVErr(xlErrDiv0), _
            CVErr(xlErrNA), _
            CVErr(xlErrName), _
            CVErr(xlErrNull), _
            CVErr(xlErrNum), _
            CVErr(xlErrRef), _
            CVErr(xlErrValue), _
            CVErr(xlErrGettingData), _
            CVErr(xlErrSpill), _
            CVErr(xlErrConnect), _
            CVErr(xlErrBlocked), _
            CVErr(xlErrUnknown), _
            CVErr(xlErrField), _
            CVErr(xlErrCalc))

    ' Act / Assert
    Dim error_idx As Long
    For error_idx = LBound(error_texts) To UBound(error_texts)
        Dim actual_error As Variant
        Dim actual_result As Boolean
        actual_result = TryConvertExcelErrorStringToCVErr(CStr(error_texts(error_idx)), actual_error)

        Assert.IsTrue actual_result
        Assert.Equals error_values(error_idx), actual_error
    Next error_idx
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_TryConvertExcelErrorStringToCVErr_LowerCase_ReturnsErrorValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_error As Variant
    Dim actual_result As Boolean
    actual_result = TryConvertExcelErrorStringToCVErr("#div/0!", actual_error)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_result
    Assert.Equals CVErr(xlErrDiv0), actual_error
End Sub

Public Sub Test_TryConvertExcelErrorStringToCVErr_UnknownText_ReturnsFalseAndEmpty(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim actual_error As Variant
    actual_error = CVErr(xlErrNA)

    ' Act
    Dim actual_result As Boolean
    actual_result = TryConvertExcelErrorStringToCVErr("#ERRNO_2045!", actual_error)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_result
    Assert.IsTrue IsEmpty(actual_error)
End Sub

' -----------------------------------------------------------------------------
' ReplaceSpecialCharacterOnFileSystemPath
' -----------------------------------------------------------------------------

Public Sub Test_ReplaceSpecialCharacterOnFileSystemPath_ReservedCharacters_ReturnsFullWidth(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ReplaceSpecialCharacterOnFileSystemPath("\/" & ":*?" & """" & "<>|")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Å_Å^ÅFÅñÅHÅçÅÉÅÑÅb", actual_value
End Sub

' -----------------------------------------------------------------------------
' JoinPath
' -----------------------------------------------------------------------------

Public Sub Test_JoinPath_WindowsAndUrlPaths_JoinsWithExpectedSeparators(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "C:\root\child\leaf", JoinPath("C:\root\", "\child", "leaf")
    Assert.Equals "https://example.com/root/child/file.txt?x=1", JoinPath("https://example.com/root/", "\child", "file.txt?x=1")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_JoinPath_RootPaths_JoinsWithoutDuplicatingSeparators(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "C:\Work", JoinPath("C:\", "\Work")
    Assert.Equals "\\server\share\dir", JoinPath("\\server\share\", "dir")
    Assert.Equals "https://example.com/sites/book.xlsm", JoinPath("https://example.com/", "/sites", "book.xlsm")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' GetAbsolutePathFromParent
' -----------------------------------------------------------------------------

Public Sub Test_GetAbsolutePathFromParent_RelativeParentSegment_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "..\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_RootRelative_UsesParentDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_AbsolutePath_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "C:\Other\..\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_DriveRelativePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetAbsolutePathFromParent("C:\Base\Child", "C:File.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetAbsolutePathFromParent", Err.Source
End Sub

Public Sub Test_GetAbsolutePathFromParent_NonAbsoluteParentPath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetAbsolutePathFromParent("Base\Child", "File.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetAbsolutePathFromParent", Err.Source
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlParentRelativePath_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "..\book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team/book.xlsm", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlRootRelativePath_UsesUrlRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "\shared\book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/shared/book.xlsm", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlQueryFragment_PreservesSuffix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "sub\..\book.xlsm?q=C:\Temp\A/B#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team/docs/book.xlsm?q=C:\Temp\A/B#frag/c", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_WindowsForwardSlash_ReturnsBackslash(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:/Base/Child", "../File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\File.txt", actual_value
End Sub

' -----------------------------------------------------------------------------
' GetPathRoot
' -----------------------------------------------------------------------------

Public Sub Test_GetPathRoot_DrivePath_ReturnsDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("C:\Base\Child")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\", actual_value
End Sub

Public Sub Test_GetPathRoot_LowercaseDrivePath_ReturnsDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("z:\Base\Child")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Z:\", actual_value
End Sub

Public Sub Test_GetPathRoot_UncPath_ReturnsShareRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("\\server\share\folder\file.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "\\server\share\", actual_value
End Sub

Public Sub Test_GetPathRoot_UncSharePathWithoutTrailingSeparator_ReturnsShareRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("\\server\share")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "\\server\share\", actual_value
End Sub

Public Sub Test_GetPathRoot_IncompleteUncPath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetPathRoot("\\server\")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetPathRoot", Err.Source
End Sub

Public Sub Test_GetPathRoot_RootRelativePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetPathRoot("\folder\file.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetPathRoot", Err.Source
End Sub

Public Sub Test_GetPathRoot_NonAlphabetDrivePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetPathRoot("1:\folder\file.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetPathRoot", Err.Source
End Sub

Public Sub Test_GetPathRoot_Url_ReturnsAuthorityRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("https://example.com/sites/team/book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/", actual_value
End Sub

' -----------------------------------------------------------------------------
' IsDrivePath
' -----------------------------------------------------------------------------

Public Sub Test_IsDrivePath_DriveSpecifier_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDrivePath("C:\Base\File.txt")
    Assert.IsTrue IsDrivePath("C:/Base/File.txt")
    Assert.IsTrue IsDrivePath("C:Base\File.txt")
    Assert.IsTrue IsDrivePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDrivePath_NonDriveSpecifier_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDrivePath("\\Server\Share\File.txt")
    Assert.IsFalse IsDrivePath("Base\File.txt")
    Assert.IsFalse IsDrivePath("1:\Base\File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsDriveAbsolutePath
' -----------------------------------------------------------------------------

Public Sub Test_IsDriveAbsolutePath_DriveAbsolutePath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDriveAbsolutePath("C:\Base\File.txt")
    Assert.IsTrue IsDriveAbsolutePath("C:/Base/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveAbsolutePath_DriveRelativePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDriveAbsolutePath("C:Base\File.txt")
    Assert.IsFalse IsDriveAbsolutePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsDriveRelativePath
' -----------------------------------------------------------------------------

Public Sub Test_IsDriveRelativePath_DriveRelativePath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDriveRelativePath("C:Base\File.txt")
    Assert.IsTrue IsDriveRelativePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveRelativePath_DriveAbsolutePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDriveRelativePath("C:\Base\File.txt")
    Assert.IsFalse IsDriveRelativePath("C:/Base/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsUncPath
' -----------------------------------------------------------------------------

Public Sub Test_IsUncPath_ValidUncPath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsUncPath("\\Server\Share")
    Assert.IsTrue IsUncPath("\\Server\Share\")
    Assert.IsTrue IsUncPath("\\Server\Share\File.txt")
    Assert.IsTrue IsUncPath("//Server/Share/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsUncPath_IncompleteUncPath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsUncPath("\\Server")
    Assert.IsFalse IsUncPath("\\Server\")
    Assert.IsFalse IsUncPath("//Server/")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsUrlPath
' -----------------------------------------------------------------------------

Public Sub Test_IsUrlPath_SchemeUrl_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsUrlPath("https://example.com/sites/book.xlsm")
    Assert.IsTrue IsUrlPath("custom+scheme://host/path")
    Assert.IsFalse IsUrlPath("C:\Base\Child")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsAbsolutePath
' -----------------------------------------------------------------------------

Public Sub Test_IsAbsolutePath_DriveRelativePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsAbsolutePath("C:folder\file.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsAbsolutePath_AbsolutePathKinds_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsAbsolutePath("C:\Base\File.txt")
    Assert.IsTrue IsAbsolutePath("\\Server\Share")
    Assert.IsTrue IsAbsolutePath("https://example.com/sites/book.xlsm")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsAbsolutePath_StaleErr_ReturnsTrueForAbsolutePath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Err.Clear
    Err.Raise 5, "Test_Lib_Common", "stale error"
    Dim actual_value As Boolean
    actual_value = IsAbsolutePath("C:\Base\File.txt")

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsAbsolutePath_IncompleteUncPath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsAbsolutePath("\\Server\")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' SplitPath
' -----------------------------------------------------------------------------

Public Sub Test_SplitPath_WindowsPath_ReturnsParentAndLeaf(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\Base\Child\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_WindowsForwardSlash_ReturnsBackslashParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:/Base/Child/File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_DriveRootChild_ReturnsDriveRootParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\Work")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\", parent_path
    Assert.Equals "Work", leaf_path
End Sub

Public Sub Test_SplitPath_UncShareRootChild_ReturnsShareRootParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "\\server\share\Work")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "\\server\share\", parent_path
    Assert.Equals "Work", leaf_path
End Sub

Public Sub Test_SplitPath_DriveRoot_ReturnsRootParentForTrailingSeparator(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\", parent_path
    Assert.Equals "", leaf_path
End Sub

Public Sub Test_SplitPath_NoSeparator_ReturnsEmptyParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_EndSeparator_RespectsIgnoreEndSep(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\Base\Child\")

    Dim ignored_parent_path As String
    Dim ignored_leaf_path As String
    Call SplitPath(ignored_parent_path, ignored_leaf_path, "C:\Base\Child\", IgnoreEndSep:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "", leaf_path
    Assert.Equals "C:\Base", ignored_parent_path
    Assert.Equals "Child", ignored_leaf_path
End Sub

Public Sub Test_SplitPath_Url_KeepsPathSuffixInLeaf(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team", parent_path
    Assert.Equals "book.xlsm?q=a/b#frag/c", leaf_path
End Sub

' -----------------------------------------------------------------------------
' ParseLeafPath
' -----------------------------------------------------------------------------

Public Sub Test_ParseLeafPath_FileName_ReturnsBaseNameAndExtension(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, "book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book", base_name
    Assert.Equals ".xlsm", file_ext
    Assert.Equals "", path_suffix
End Sub

Public Sub Test_ParseLeafPath_HashFileName_KeepsHashInBaseName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, "book#1.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book#1", base_name
    Assert.Equals ".xlsm", file_ext
    Assert.Equals "", path_suffix
End Sub

Public Sub Test_ParseLeafPath_PathSuffix_ReturnsSuffixSeparately(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, "book.xlsm?q=a/b#frag/c", AsUrl:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book", base_name
    Assert.Equals ".xlsm", file_ext
    Assert.Equals "?q=a/b#frag/c", path_suffix
End Sub

Public Sub Test_ParseLeafPath_DotFile_ReturnsNoExtension(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, ".gitignore#frag", AsUrl:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals ".gitignore", base_name
    Assert.Equals "", file_ext
    Assert.Equals "#frag", path_suffix
End Sub

' -----------------------------------------------------------------------------
' GetParentPath
' -----------------------------------------------------------------------------

Public Sub Test_GetParentPath_Url_IgnoresQueryFragment(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetParentPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team", actual_value
End Sub

Public Sub Test_GetParentPath_NoSeparator_ReturnsEmptyParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetParentPath("File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", actual_value
End Sub

Public Sub Test_GetParentPath_DriveRootChild_ReturnsDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "C:\", GetParentPath("C:\Work")
    Assert.Equals "C:\", GetParentPath("C:\Work\", IgnoreEndSep:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetParentPath_UncShareRootChild_ReturnsShareRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "\\server\share\", GetParentPath("\\server\share\Work")
    Assert.Equals "\\server\share\", GetParentPath("\\server\share\Work\", IgnoreEndSep:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetParentPath_UrlRootChild_ReturnsAuthorityRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetParentPath("https://example.com/book.xlsm?q=a/b")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/", actual_value
End Sub

' -----------------------------------------------------------------------------
' GetLeafFromPath
' -----------------------------------------------------------------------------

Public Sub Test_GetLeafFromPath_Url_IgnoresQueryFragment(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetLeafFromPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book.xlsm", actual_value
End Sub

Public Sub Test_GetLeafFromPath_UrlExtension_IgnoresPathSuffix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetLeafFromPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c", BaseName:=False, Extension:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals ".xlsm", actual_value
End Sub

Public Sub Test_GetLeafFromPath_LocalHashFileName_ReturnsLeafParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "book#1.xlsm", GetLeafFromPath("C:\tmp\book#1.xlsm")
    Assert.Equals ".xlsm", GetLeafFromPath("C:\tmp\book#1.xlsm", BaseName:=False, Extension:=True)
    Assert.Equals "book#1", GetLeafFromPath("C:\tmp\book#1.xlsm", Extension:=False)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetLeafFromPath_LocalHashFileNameWithDots_ReturnsLeafParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "book.name#v2.xlsm", GetLeafFromPath("C:\tmp\book.name#v2.xlsm")
    Assert.Equals ".xlsm", GetLeafFromPath("C:\tmp\book.name#v2.xlsm", BaseName:=False, Extension:=True)
    Assert.Equals "book.name#v2", GetLeafFromPath("C:\tmp\book.name#v2.xlsm", Extension:=False)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetLeafFromPath_EndSeparatorWithIgnoreEndSep_ReturnsLeaf(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetLeafFromPath("C:\Base\Child\", IgnoreEndSep:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Child", actual_value
End Sub

Public Sub Test_GetLeafFromPath_DotFile_ReturnsWholeNameWithoutExtension(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals ".gitignore", GetLeafFromPath(".gitignore")
    Assert.Equals "", GetLeafFromPath(".gitignore", BaseName:=False, Extension:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsNullDate
' -----------------------------------------------------------------------------

Public Sub Test_IsNullDate_NullAndValidDate_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsNullDate(G_DATE_NULL)
    Assert.IsFalse IsNullDate(DateSerial(2026, 1, 1))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ConvertStringToCharArray
' -----------------------------------------------------------------------------

Public Sub Test_ConvertStringToCharArray_Text_ReturnsCharacters(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertStringToCharArray("abc")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "a", actual_arr(0)
    Assert.Equals "b", actual_arr(1)
    Assert.Equals "c", actual_arr(2)
End Sub

Public Sub Test_ConvertStringToCharArray_EmptyString_ReturnsInitializedZeroLengthArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertStringToCharArray("")

    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, actual_arr)

    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, lbound_arr(0)
    Assert.EqualsNumeric -1, ubound_arr(0)
    Assert.IsFalse enum_obj.MoveNext()
End Sub

Public Sub Test_ConvertStringToCharArray_SingleCharacter_ReturnsOneCharacter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertStringToCharArray("a")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 0, UBound(actual_arr)
    Assert.Equals "a", actual_arr(0)
End Sub

Public Sub Test_ConvertStringToCharArray_UnicodeSequence_ReturnsVbaStringUnits(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_text As String
    src_text = ChrW$(&HD83D) & ChrW$(&HDE00) & "e" & ChrW$(&H301)

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertStringToCharArray(src_text)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 3, UBound(actual_arr)
    Assert.EqualsNumeric AscW(Mid$(src_text, 1, 1)), AscW(actual_arr(0))
    Assert.EqualsNumeric AscW(Mid$(src_text, 2, 1)), AscW(actual_arr(1))
    Assert.Equals "e", actual_arr(2)
    Assert.EqualsNumeric AscW(Mid$(src_text, 4, 1)), AscW(actual_arr(3))
End Sub

' -----------------------------------------------------------------------------
' ConvertArrayStringToVariant
' -----------------------------------------------------------------------------

Public Sub Test_ConvertArrayStringToVariant_StringArray_ReturnsVariantArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(1 To 2) As String
    src_arr(1) = "alpha"
    src_arr(2) = "beta"

    ' Act
    Dim actual_arr() As Variant
    actual_arr = ConvertArrayStringToVariant(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(1)
    Assert.Equals "beta", actual_arr(2)
End Sub

Public Sub Test_ConvertArrayStringToVariant_EmptyStringArray_ReturnsEmptyVariantArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr() As String
    src_arr = EmptyStringArray()

    ' Act
    Dim actual_arr() As Variant
    actual_arr = ConvertArrayStringToVariant(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
End Sub

' -----------------------------------------------------------------------------
' ConvertArrayVariantToString
' -----------------------------------------------------------------------------

Public Sub Test_ConvertArrayVariantToString_VariantArray_ReturnsStringArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(1 To 2) As Variant
    src_arr(1) = "alpha"
    src_arr(2) = "beta"

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertArrayVariantToString(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(1)
    Assert.Equals "beta", actual_arr(2)
End Sub

Public Sub Test_ConvertArrayVariantToString_EmptyVariantArray_ReturnsEmptyStringArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr() As Variant
    src_arr = EmptyVariantArray()

    ' Act
    Dim actual_arr() As String
    actual_arr = ConvertArrayVariantToString(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
End Sub

' -----------------------------------------------------------------------------
' ConvertBooleanToString
' -----------------------------------------------------------------------------

Public Sub Test_ConvertBooleanToString_CustomLabels_ReturnsLabel(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "ON", ConvertBooleanToString(True, "ON", "OFF")
    Assert.Equals "OFF", ConvertBooleanToString(False, "ON", "OFF")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ConvertStringToBoolean
' -----------------------------------------------------------------------------

Public Sub Test_ConvertStringToBoolean_CustomLabels_ReturnsBoolean(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue ConvertStringToBoolean("ON", "ON", "OFF")
    Assert.IsFalse ConvertStringToBoolean("OFF", "ON", "OFF")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ConvertArray2dTo1d
' -----------------------------------------------------------------------------

Public Sub Test_ConvertArray2dTo1d_RowAndColumnDirection_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(1 To 2, 1 To 3) As Variant
    src_arr(1, 1) = "A"
    src_arr(1, 2) = "B"
    src_arr(1, 3) = "C"
    src_arr(2, 1) = "D"
    src_arr(2, 2) = "E"
    src_arr(2, 3) = "F"

    ' Act
    Dim row_arr() As Variant
    row_arr = ConvertArray2dTo1d(src_arr)

    Dim col_arr() As Variant
    col_arr = ConvertArray2dTo1d(src_arr, ColumnDirection:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsArray Array("A", "B", "C", "D", "E", "F"), row_arr
    Assert.EqualsArray Array("A", "D", "B", "E", "C", "F"), col_arr
End Sub

Public Sub Test_ConvertArray2dTo1d_ObjectValues_PreservesReferences(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub

    Dim second_obj As Test_ObjectSetEquatableStub
    Set second_obj = New Test_ObjectSetEquatableStub

    Dim src_arr(1 To 1, 1 To 2) As Variant
    Set src_arr(1, 1) = first_obj
    Set src_arr(1, 2) = second_obj

    ' Act
    Dim actual_arr() As Variant
    actual_arr = ConvertArray2dTo1d(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim actual_first As Object
    Set actual_first = actual_arr(0)
    Dim actual_second As Object
    Set actual_second = actual_arr(1)

    Dim first_is_same As Boolean
    first_is_same = (actual_first Is first_obj)
    Assert.IsTrue first_is_same

    Dim second_is_same As Boolean
    second_is_same = (actual_second Is second_obj)
    Assert.IsTrue second_is_same
End Sub

' -----------------------------------------------------------------------------
' ConvertArray1dTo2d
' -----------------------------------------------------------------------------

Public Sub Test_ConvertArray1dTo2d_RowAndColumnDirection_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array("A", "B", "C", "D", "E")

    ' Act
    Dim row_arr() As Variant
    row_arr = ConvertArray1dTo2d(src_arr, RowCount:=2)

    Dim col_arr() As Variant
    col_arr = ConvertArray1dTo2d(src_arr, ColCount:=2, ColumnDirection:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, LBound(row_arr, 1)
    Assert.EqualsNumeric 2, UBound(row_arr, 1)
    Assert.EqualsNumeric 1, LBound(row_arr, 2)
    Assert.EqualsNumeric 3, UBound(row_arr, 2)
    Assert.Equals "A", row_arr(1, 1)
    Assert.Equals "B", row_arr(1, 2)
    Assert.Equals "C", row_arr(1, 3)
    Assert.Equals "D", row_arr(2, 1)
    Assert.Equals "E", row_arr(2, 2)
    Assert.IsEmpty row_arr(2, 3)

    Assert.EqualsNumeric 1, LBound(col_arr, 1)
    Assert.EqualsNumeric 3, UBound(col_arr, 1)
    Assert.EqualsNumeric 1, LBound(col_arr, 2)
    Assert.EqualsNumeric 2, UBound(col_arr, 2)
    Assert.Equals "A", col_arr(1, 1)
    Assert.Equals "B", col_arr(2, 1)
    Assert.Equals "C", col_arr(3, 1)
    Assert.Equals "D", col_arr(1, 2)
    Assert.Equals "E", col_arr(2, 2)
    Assert.IsEmpty col_arr(3, 2)
End Sub

Public Sub Test_EmptyVariantArray_ReturnsInitializedZeroLengthArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr As Variant
    actual_arr = EmptyVariantArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsArray(actual_arr)
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
End Sub

Public Sub Test_EmptyStringArray_ReturnsInitializedZeroLengthArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = EmptyStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(actual_arr)
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric -1, UBound(actual_arr)
End Sub

' -----------------------------------------------------------------------------
' GetArrayBounds
' -----------------------------------------------------------------------------

Public Sub Test_GetArrayBounds_MultidimensionalArray_ReturnsBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(2 To 3, 4 To 6) As String

    ' Act
    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(lbound_arr)
    Assert.EqualsNumeric 1, UBound(lbound_arr)
    Assert.EqualsNumeric 2, lbound_arr(0)
    Assert.EqualsNumeric 4, lbound_arr(1)
    Assert.EqualsNumeric 3, ubound_arr(0)
    Assert.EqualsNumeric 6, ubound_arr(1)
End Sub

Public Sub Test_GetArrayBounds_StaleErr_ReturnsBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(2 To 3, 4 To 6) As String

    ' Act
    Err.Clear
    Err.Raise 5, "Test_Lib_Common", "stale error"
    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, src_arr)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.EqualsNumeric 2, lbound_arr(0)
    Assert.EqualsNumeric 4, lbound_arr(1)
    Assert.EqualsNumeric 3, ubound_arr(0)
    Assert.EqualsNumeric 6, ubound_arr(1)
End Sub

Public Sub Test_GetArrayBounds_EmptyArray_ReturnsOneDimensionalEmptyBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array()

    ' Act
    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(lbound_arr)
    Assert.EqualsNumeric 0, UBound(lbound_arr)
    Assert.EqualsNumeric 0, lbound_arr(0)
    Assert.EqualsNumeric -1, ubound_arr(0)
End Sub

Public Sub Test_GetArrayBounds_UninitializedArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr() As Variant

    ' Act
    Dim lbound_arr() As Long
    Dim ubound_arr() As Long
    Call GetArrayBounds(lbound_arr, ubound_arr, src_arr)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsEmptyArray
' -----------------------------------------------------------------------------

Public Sub Test_IsEmptyArray_DynamicAndInitialized_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim empty_arr() As String
    Dim filled_arr(0 To 0) As String

    ' Act / Assert
    Assert.IsTrue IsEmptyArray(empty_arr)
    Assert.IsFalse IsEmptyArray(filled_arr)
    Assert.IsFalse IsEmptyArray("not array")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' GetArrayEnumerator
' -----------------------------------------------------------------------------

Public Sub Test_GetArrayEnumerator_Descending_ReturnsReversedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array("A", "B", "C")

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(src_arr, Descending:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue enum_obj.MoveNext()
    Assert.Equals "C", enum_obj.Current
    Assert.IsTrue enum_obj.MoveNext()
    Assert.Equals "B", enum_obj.Current
    Assert.IsTrue enum_obj.MoveNext()
    Assert.Equals "A", enum_obj.Current
    Assert.IsFalse enum_obj.MoveNext()
End Sub

' -----------------------------------------------------------------------------
' ContainsInArray
' -----------------------------------------------------------------------------

Public Sub Test_ContainsInArray_MatchingItem_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array("alpha", "beta", "gamma")

    ' Act / Assert
    Assert.IsTrue ContainsInArray("beta", src_arr)
    Assert.IsFalse ContainsInArray("delta", src_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ContainsInArray_ObjectValues_UsesSameInstanceOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub
    first_obj.IdentityKey = "A"

    Dim same_key_obj As Test_ObjectSetEquatableStub
    Set same_key_obj = New Test_ObjectSetEquatableStub
    same_key_obj.IdentityKey = "A"

    Dim src_arr(0 To 0) As Variant
    Set src_arr(0) = first_obj

    ' Act / Assert
    Assert.IsTrue ContainsInArray(first_obj, src_arr)
    Assert.IsFalse ContainsInArray(same_key_obj, src_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ContainsInArray_NullValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array("alpha", Null, "gamma")

    Dim non_null_arr As Variant
    non_null_arr = Array("alpha", "beta")

    ' Act / Assert
    Assert.IsTrue ContainsInArray(Null, src_arr)
    Assert.IsFalse ContainsInArray("delta", src_arr)
    Assert.IsFalse ContainsInArray(Null, non_null_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ContainsInArray_ExcelErrorValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr As Variant
    src_arr = Array("alpha", CVErr(xlErrNA), CVErr(xlErrDiv0))

    Dim normal_arr As Variant
    normal_arr = Array("alpha", "beta")

    ' Act / Assert
    Assert.IsTrue ContainsInArray(CVErr(xlErrNA), src_arr)
    Assert.IsFalse ContainsInArray(CVErr(xlErrValue), src_arr)
    Assert.IsFalse ContainsInArray(CVErr(xlErrNA), normal_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ContainsInArray_EmptyAndArrayValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim empty_item As Variant

    Dim empty_arr As Variant
    empty_arr = Array(empty_item, "alpha")

    Dim nested_arr As Variant
    nested_arr = Array(1, 2)

    Dim array_item_arr As Variant
    array_item_arr = Array(nested_arr, "alpha")

    ' Act / Assert
    Assert.IsTrue ContainsInArray(empty_item, empty_arr)
    Assert.IsFalse ContainsInArray(empty_item, Array("alpha"))
    Assert.IsTrue ContainsInArray("alpha", array_item_arr)
    Assert.IsFalse ContainsInArray(nested_arr, array_item_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' SortArray
' -----------------------------------------------------------------------------

Public Sub Test_SortArray_AscendingAndDescending_SortsValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim asc_arr As Variant
    asc_arr = Array(3, 1, 2)

    Dim desc_arr As Variant
    desc_arr = Array(3, 1, 2)

    ' Act
    Call SortArray(asc_arr)
    Call SortArray(desc_arr, Descending:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsArray Array(1, 2, 3), asc_arr
    Assert.EqualsArray Array(3, 2, 1), desc_arr
End Sub

Public Sub Test_SortArray_EmptyAndSingleElement_NoError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim empty_arr As Variant
    empty_arr = Array()

    Dim single_arr As Variant
    single_arr = Array(2)

    ' Act
    Call SortArray(empty_arr)
    Call SortArray(single_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(empty_arr)
    Assert.EqualsArray Array(2), single_arr
End Sub

Public Sub Test_SortArray_ObjectValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub

    Dim src_arr(0 To 0) As Variant
    Set src_arr(0) = first_obj

    ' Act
    Call SortArray(src_arr)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SortArray", Err.Source
End Sub

' -----------------------------------------------------------------------------
' ConcatArray
' -----------------------------------------------------------------------------

Public Sub Test_ConcatArray_ArraysAndScalar_ReturnsCombined(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_arr As Variant
    first_arr = Array("A", "B")

    Dim second_arr As Variant
    second_arr = Array("C")

    ' Act
    Dim actual_arr() As Variant
    actual_arr = ConcatArray(first_arr, second_arr, "D", Array("E", "F"))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsArray Array("A", "B", "C", "D", "E", "F"), actual_arr
End Sub

Public Sub Test_ConcatArray_ObjectValues_PreservesReferences(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub

    Dim second_obj As Test_ObjectSetEquatableStub
    Set second_obj = New Test_ObjectSetEquatableStub

    Dim third_obj As Test_ObjectSetEquatableStub
    Set third_obj = New Test_ObjectSetEquatableStub

    Dim first_arr(0 To 0) As Variant
    Set first_arr(0) = first_obj

    Dim third_arr(0 To 0) As Variant
    Set third_arr(0) = third_obj

    ' Act
    Dim actual_arr() As Variant
    actual_arr = ConcatArray(first_arr, second_obj, third_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Dim actual_first As Object
    Set actual_first = actual_arr(0)
    Dim actual_second As Object
    Set actual_second = actual_arr(1)
    Dim actual_third As Object
    Set actual_third = actual_arr(2)

    Dim first_is_same As Boolean
    first_is_same = (actual_first Is first_obj)
    Assert.IsTrue first_is_same

    Dim second_is_same As Boolean
    second_is_same = (actual_second Is second_obj)
    Assert.IsTrue second_is_same

    Dim third_is_same As Boolean
    third_is_same = (actual_third Is third_obj)
    Assert.IsTrue third_is_same
End Sub

' -----------------------------------------------------------------------------
' New_ObjectList / New_ObjectSet
' -----------------------------------------------------------------------------

Public Sub Test_New_ObjectList_WithContract_ReturnsInitializedList(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As ObjectList
    Set actual_value = New_ObjectList( _
            ElementTypeName:="ILeafCondition", _
            RequireComparable:=True, _
            ObjectKeyMode:=G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_value.Count
    Assert.IsTrue actual_value.HasItemTypeContract
    Assert.Equals "ILeafCondition", actual_value.ElementTypeName
    Assert.Equals "Object@ILeafCondition", actual_value.ItemTypeName
    Assert.IsTrue actual_value.RequireComparable
    Assert.EqualsNumeric G_OBJECT_KEY_MODE_I_EQUATABLE, actual_value.ObjectKeyMode
End Sub

Public Sub Test_New_ObjectSet_WithContract_ReturnsInitializedSet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As ObjectSet
    Set actual_value = New_ObjectSet( _
            ElementTypeName:="ILeafCondition", _
            ObjectKeyMode:=G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, actual_value.Count
    Assert.IsTrue actual_value.HasItemTypeContract
    Assert.Equals "ILeafCondition", actual_value.ElementTypeName
    Assert.Equals "Object@ILeafCondition", actual_value.ItemTypeName
    Assert.IsFalse actual_value.RequireComparable
    Assert.EqualsNumeric G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE, actual_value.ObjectKeyMode
End Sub
' -----------------------------------------------------------------------------
' FormatIDName
' -----------------------------------------------------------------------------

Public Sub Test_FormatIDName_CustomFormat_ReturnsFormattedName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = FormatIDName(7, "Sample", NumFormat:="000", Separator:="-")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "007-Sample", actual_value
End Sub

' -----------------------------------------------------------------------------
' IsValidElementTypeKey
' -----------------------------------------------------------------------------

Public Sub Test_IsValidElementTypeKey_ValidAndInvalidValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsValidElementTypeKey("ILeafCondition")
    Assert.IsTrue IsValidElementTypeKey("WorksheetRangeBounds")
    Assert.IsTrue IsValidElementTypeKey("Test_ObjectSetEquatableStub")
    Assert.IsFalse IsValidElementTypeKey("")
    Assert.IsFalse IsValidElementTypeKey("1Invalid")
    Assert.IsFalse IsValidElementTypeKey("Invalid-Type")
    Assert.IsFalse IsValidElementTypeKey("Invalid Key!")
    Assert.IsFalse IsValidElementTypeKey(String$(32, "A"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub
' -----------------------------------------------------------------------------
' GetTypeString
' -----------------------------------------------------------------------------

Public Sub Test_GetTypeString_PrimitiveObjectAndArray_ReturnsTypePrefix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim long_arr(0 To 1) As Long

    Dim matrix(1 To 2, 1 To 3) As String

    Dim variant_arr As Variant
    variant_arr = Array("A", "B")

    Dim item_obj As Test_ObjectSetEquatableStub
    Set item_obj = New Test_ObjectSetEquatableStub

    ' Act / Assert
    Assert.Equals "String", GetTypeString("A")
    Assert.Equals "Long", GetTypeString(CLng(1))
    Assert.Equals "Long[]", GetTypeString(long_arr)
    Assert.Equals "Long[]", GetTypeString(long_arr, IncludeArrayRank:=True)
    Assert.Equals "Long[0:1]", GetTypeString(long_arr, IncludeArrayBounds:=True)
    Assert.Equals "String[,]", GetTypeString(matrix, IncludeArrayRank:=True)
    Assert.Equals "String[1:2,1:3]", GetTypeString(matrix, IncludeArrayBounds:=True)
    Assert.Equals "Variant[]", GetTypeString(variant_arr)
    Assert.Equals "Variant[0:1]", GetTypeString(variant_arr, IncludeArrayBounds:=True)
    Assert.Equals "Object@Test_ObjectSetEquatableStub", GetTypeString(item_obj)
    Assert.Equals "IEquatable@Test_ObjectSetEquatableStub", GetTypeString(item_obj, G_OBJECT_KEY_MODE_I_EQUATABLE)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypeString_ElementTypeProviderWithOption_UsesElementTypeKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ElementTypeProviderStub
    Set item_obj = New Test_ElementTypeProviderStub
    item_obj.ElementTypeKey = "ILeafCondition"

    ' Act / Assert
    Assert.Equals "Object@Test_ElementTypeProviderStub", GetTypeString(item_obj)
    Assert.Equals "Object@ILeafCondition", GetTypeString(item_obj, UseElementTypeKey:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' GetTypedValueKey
' -----------------------------------------------------------------------------

Public Sub Test_GetTypedValueKey_PrimitiveValues_ReturnsTypedKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "String(A\,B\(C\)\\D)", GetTypedValueKey("A,B(C)\D")
    Assert.Equals "String(A\tB)", GetTypedValueKey("A" & vbTab & "B")
    Assert.Equals "Long(1)", GetTypedValueKey(CLng(1))
    Assert.Equals "String(1)", GetTypedValueKey(CStr(1))
    Assert.NotEquals GetTypedValueKey(CLng(1)), GetTypedValueKey(CStr(1))
    Assert.Equals "Boolean(True)", GetTypedValueKey(CBool(True))
    Assert.Equals "Currency(1)", GetTypedValueKey(CCur(1))
    Assert.Equals "Null()", GetTypedValueKey(Null)
    Assert.Equals "Empty()", GetTypedValueKey(Empty)
    Assert.Equals "Error(2042)", GetTypedValueKey(CVErr(2042))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueKey_PrimitiveArrays_ReturnsArrayTypeAndBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim string_arr(1 To 2) As String
    string_arr(1) = "A"
    string_arr(2) = "B"

    Dim long_arr(0 To 1) As Long
    long_arr(0) = 1
    long_arr(1) = 2

    Dim variant_arr As Variant
    variant_arr = Array("1", CLng(1), Empty, Null, CVErr(2042))

    ' Act / Assert
    Assert.Equals "String[1:2](A,B)", GetTypedValueKey(string_arr)
    Assert.Equals "Long[0:1](1,2)", GetTypedValueKey(long_arr)
    Assert.Equals "Variant[0:4](String(1),Long(1),Empty(),Null(),Error(2042))", GetTypedValueKey(variant_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueKey_StaleErr_ReturnsArrayBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim string_arr(0 To 1) As String
    string_arr(0) = "A"
    string_arr(1) = "B"

    ' Act
    Err.Clear
    Err.Raise 5, "Test_Lib_Common", "stale error"
    Dim actual_value As String
    actual_value = GetTypedValueKey(string_arr)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "String[0:1](A,B)", actual_value
End Sub

Public Sub Test_GetTypedValueKey_MultidimensionalArrays_ReturnsNestedItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim matrix(1 To 2, 1 To 3) As String
    matrix(1, 1) = "A"
    matrix(1, 2) = "B"
    matrix(1, 3) = "C"
    matrix(2, 1) = "D"
    matrix(2, 2) = "E"
    matrix(2, 3) = "F"

    Dim cube(1 To 2, 1 To 2, 1 To 2) As Long
    cube(1, 1, 1) = 111
    cube(1, 1, 2) = 112
    cube(1, 2, 1) = 121
    cube(1, 2, 2) = 122
    cube(2, 1, 1) = 211
    cube(2, 1, 2) = 212
    cube(2, 2, 1) = 221
    cube(2, 2, 2) = 222

    ' Act / Assert
    Assert.Equals "String[1:2,1:3]((A,B,C),(D,E,F))", GetTypedValueKey(matrix)
    Assert.Equals "Long[1:2,1:2,1:2](((111,112),(121,122)),((211,212),(221,222)))", GetTypedValueKey(cube)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueKey_JaggedArray_ReturnsRecursiveArrayItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim jagged_arr(0 To 1) As Variant
    jagged_arr(0) = Array("A", "B")
    jagged_arr(1) = Array("C", "D")

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(jagged_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:1](Variant[0:1](String(A),String(B)),Variant[0:1](String(C),String(D)))", actual_value
End Sub

Public Sub Test_GetTypedValueKey_EmptyArrays_ReturnsArrayTypeOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim long_arr() As Long

    Dim variant_arr As Variant
    variant_arr = Array()

    ' Act / Assert
    Assert.Equals "Long[]()", GetTypedValueKey(long_arr)
    Assert.Equals "Variant[0:-1]()", GetTypedValueKey(variant_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueKey_IEquatableObject_ReturnsIdentityString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ObjectSetEquatableStub
    Set item_obj = New Test_ObjectSetEquatableStub
    item_obj.IdentityKey = "A" & vbTab & "B@(C)"

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(item_obj, G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IEquatable@Test_ObjectSetEquatableStub(A\tB\@\(C\))", actual_value
End Sub

Public Sub Test_GetTypedValueKey_IEquatableObjectWithElementTypeOption_UsesElementTypeKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ElementTypeEquatableStub
    Set item_obj = New Test_ElementTypeEquatableStub
    item_obj.ElementTypeKey = "ILeafCondition"
    item_obj.IdentityKey = "A" & vbTab & "B@(C)"

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(item_obj, G_OBJECT_KEY_MODE_I_EQUATABLE, UseElementTypeKey:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IEquatable@ILeafCondition(A\tB\@\(C\))", actual_value
End Sub

Public Sub Test_GetTypedMultiKeyByModeFromArray_ElementTypeProviderWithOption_UsesElementTypeKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ElementTypeEquatableStub
    Set item_obj = New Test_ElementTypeEquatableStub
    item_obj.ElementTypeKey = "ILeafCondition"
    item_obj.IdentityKey = "same-id"

    Dim key_arr(0 To 1) As Variant
    key_arr(0) = "prefix"
    Set key_arr(1) = item_obj

    ' Act
    Dim actual_value As String
    actual_value = GetTypedMultiKeyByModeFromArray( _
            G_OBJECT_KEY_MODE_I_EQUATABLE, _
            key_arr, _
            UseElementTypeKey:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "String(prefix)" & vbTab & "IEquatable@ILeafCondition(same-id)", actual_value
End Sub

Public Sub Test_GetTypedValueKey_IDuplicateCheckableObject_ReturnsDuplicateKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ObjectSetDupCheckStub
    Set item_obj = New Test_ObjectSetDupCheckStub
    item_obj.DuplicateKey = "Key=1"

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(item_obj, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IDuplicateCheckable@Test_ObjectSetDupCheckStub(Key\=1)", actual_value
End Sub

Public Sub Test_GetTypedValueKey_IEquatableArray_ReturnsPolicyAndDeclaredArrayType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_item As Test_ObjectSetEquatableStub
    Set first_item = New Test_ObjectSetEquatableStub
    first_item.IdentityKey = "A"

    Dim second_item As Test_ObjectSetEquatableStub
    Set second_item = New Test_ObjectSetEquatableStub
    second_item.IdentityKey = "B"

    Dim item_arr(0 To 1) As IEquatable
    Set item_arr(0) = first_item
    Set item_arr(1) = second_item

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(item_arr, G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IEquatable@IEquatable[0:1](IEquatable@Test_ObjectSetEquatableStub(A),IEquatable@Test_ObjectSetEquatableStub(B))", actual_value
End Sub

Public Sub Test_GetTypedValueKey_IEquatableModeWithUnsupportedObject_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Object
    Set item_obj = CreateObject("Scripting.Dictionary")

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueKey(item_obj, G_OBJECT_KEY_MODE_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetTypedValueKey", Err.Source
End Sub

' -----------------------------------------------------------------------------
' GetValueKey
' -----------------------------------------------------------------------------

Public Sub Test_GetValueKey_PrimitiveValues_ReturnsPrimitiveKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim date_value As Date
    date_value = DateSerial(2026, 1, 2) + TimeSerial(3, 4, 5)

    ' Act / Assert
    Assert.Equals "Primitive(1)", GetValueKey(CByte(1))
    Assert.Equals "Primitive(1)", GetValueKey(CInt(1))
    Assert.Equals "Primitive(1)", GetValueKey(CLng(1))
    Assert.Equals "Primitive(1)", GetValueKey(CStr(1))
    Assert.Equals "Primitive(1.5)", GetValueKey(CSng(1.5))
    Assert.Equals "Primitive(1.5)", GetValueKey(CDbl(1.5))
    Assert.Equals "Primitive(True)", GetValueKey(CBool(True))
    Assert.Equals "Primitive(True)", GetValueKey(CStr("True"))
    Assert.Equals "Primitive(2026-01-02T03\:04\:05)", GetValueKey(date_value)
    Assert.Equals "Currency(1)", GetValueKey(CCur(1))
    Assert.Equals "Null()", GetValueKey(Null)
    Assert.Equals "Empty()", GetValueKey(Empty)
    Assert.Equals "Error(2042)", GetValueKey(CVErr(2042))
    Assert.NotEquals GetValueKey(CCur(1)), GetValueKey(CLng(1))
    Assert.Equals GetValueKey(date_value), GetValueKey(CStr("2026-01-02T03:04:05"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetValueKey_PrimitiveArrays_ReturnsPrimitiveArrayType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim string_arr(1 To 2) As String
    string_arr(1) = "A"
    string_arr(2) = "B"

    Dim byte_arr(0 To 1) As Byte
    byte_arr(0) = 1
    byte_arr(1) = 2

    Dim long_arr(0 To 1) As Long
    long_arr(0) = 1
    long_arr(1) = 2

    Dim variant_arr As Variant
    variant_arr = Array(CByte(1), CLng(1), CStr(1), CCur(1), Empty, Null, CVErr(2042))

    ' Act / Assert
    Assert.Equals "Primitive[1:2](A,B)", GetValueKey(string_arr)
    Assert.Equals "Primitive[0:1](1,2)", GetValueKey(byte_arr)
    Assert.Equals "Primitive[0:1](1,2)", GetValueKey(long_arr)
    Assert.Equals "Variant[0:6](Primitive(1),Primitive(1),Primitive(1),Currency(1),Empty(),Null(),Error(2042))", GetValueKey(variant_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' GetTypedMultiKey
' -----------------------------------------------------------------------------

Public Sub Test_GetTypedMultiKey_DifferentPrimitiveTypes_ReturnsDifferentKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.NotEquals GetTypedMultiKey(CLng(1)), GetTypedMultiKey(CStr(1))
    Assert.NotEquals GetTypedMultiKey(CBool(True)), GetTypedMultiKey(CStr("True"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedMultiKey_EmptyArguments_ReturnsEmptyKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim actual_value As String
    actual_value = GetTypedMultiKey()

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", actual_value
End Sub

Public Sub Test_GetTypedMultiKeyByMode_IEquatableObjects_UsesIdentityKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim first_obj As Test_ObjectSetEquatableStub
    Set first_obj = New Test_ObjectSetEquatableStub
    first_obj.IdentityKey = "same-key"

    Dim second_obj As Test_ObjectSetEquatableStub
    Set second_obj = New Test_ObjectSetEquatableStub
    second_obj.IdentityKey = "same-key"

    Dim first_key As String
    first_key = GetTypedMultiKeyByMode(G_OBJECT_KEY_MODE_I_EQUATABLE, first_obj)

    Dim second_key As String
    second_key = GetTypedMultiKeyByMode(G_OBJECT_KEY_MODE_I_EQUATABLE, second_obj)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals first_key, second_key
End Sub

Public Sub Test_GetTypedMultiKeyByMode_UnsupportedObject_FallsBackToReferenceKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim first_obj As Collection
    Set first_obj = New Collection

    Dim second_obj As Collection
    Set second_obj = New Collection

    Dim first_key As String
    first_key = GetTypedMultiKeyByMode(G_OBJECT_KEY_MODE_I_EQUATABLE, first_obj)

    Dim second_key As String
    second_key = GetTypedMultiKeyByMode(G_OBJECT_KEY_MODE_I_EQUATABLE, second_obj)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.NotEquals first_key, second_key
End Sub

Public Sub Test_GetTypedMultiKeyByMode_InvalidModeWithEmptyArguments_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim actual_value As String
    actual_value = GetTypedMultiKeyByMode(9999)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetTypedMultiKeyByModeFromArray", Err.Source
End Sub

Public Sub Test_GetTypedMultiKeyByModeFromArray_NonArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim actual_value As String
    actual_value = GetTypedMultiKeyByModeFromArray(G_OBJECT_KEY_MODE_REFERENCE, CLng(1))

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetTypedMultiKeyByModeFromArray", Err.Source
End Sub

' -----------------------------------------------------------------------------
' GetMultiKey
' -----------------------------------------------------------------------------
Public Sub Test_GetMultiKey_PrimitiveTypes_ReturnsPrimitiveKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(CByte(1), CLng(1), CStr(1), CBool(True), CStr("True"), CCur(1))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Primitive(1)" & vbTab & "Primitive(1)" & vbTab & "Primitive(1)" & vbTab & "Primitive(True)" & vbTab & "Primitive(True)" & vbTab & "Currency(1)", actual_value
End Sub

Public Sub Test_GetMultiKey_Null_ReturnsStrictKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Null()", actual_value
End Sub

Public Sub Test_GetMultiKey_ErrorValue_ReturnsStrictKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(CVErr(2042))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Error(2042)", actual_value
End Sub

Public Sub Test_GetMultiKey_TabInValue_UsesValueKeyEscaping(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey("A" & vbTab & "B", CLng(1))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Primitive(A\tB)" & vbTab & "Primitive(1)", actual_value
End Sub

Public Sub Test_GetMultiKey_ArrayValue_ReturnsValueArrayKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim variant_arr As Variant
    variant_arr = Array("A", "B")

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(variant_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:1](Primitive(A),Primitive(B))", actual_value
End Sub

Public Sub Test_GetMultiKey_EmptyArrayValue_ReturnsValueArrayKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim variant_arr As Variant
    variant_arr = Array()

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(variant_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:-1]()", actual_value
End Sub

' -----------------------------------------------------------------------------
' IsEmptyStringArray
' -----------------------------------------------------------------------------

Public Sub Test_IsEmptyStringArray_BasicArrays_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim uninitialized_arr() As String

    Dim zero_length_arr() As String
    zero_length_arr = EmptyStringArray()

    Dim single_empty_string_arr(0 To 0) As String

    Dim one_value_arr(0 To 0) As String
    one_value_arr(0) = "alpha"

    Dim values_arr(0 To 1) As String
    values_arr(0) = "alpha"
    values_arr(1) = "beta"

    ' Act / Assert
    Assert.IsTrue IsEmptyStringArray(uninitialized_arr)
    Assert.IsTrue IsEmptyStringArray(zero_length_arr)
    Assert.IsTrue IsEmptyStringArray(single_empty_string_arr)
    Assert.IsFalse IsEmptyStringArray(one_value_arr)
    Assert.IsFalse IsEmptyStringArray(values_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsEmptyStringArray_BlankAsEmptyFalse_DistinguishesEmptyStringItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim zero_length_arr() As String
    zero_length_arr = EmptyStringArray()

    Dim single_empty_string_arr(0 To 0) As String

    Dim two_empty_string_arr(0 To 1) As String

    ' Act / Assert
    Assert.IsTrue IsEmptyStringArray(zero_length_arr, BlankAsEmpty:=False)
    Assert.IsFalse IsEmptyStringArray(single_empty_string_arr, BlankAsEmpty:=False)
    Assert.IsFalse IsEmptyStringArray(two_empty_string_arr, BlankAsEmpty:=False)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ReplaceMulti
' -----------------------------------------------------------------------------

Public Sub Test_ReplaceMulti_ArrayReplacement_ReturnsAllCombinations(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ReplaceMulti("a-b", "a", Array("x", "y"), "b", Array("1", "2"))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 3, UBound(actual_arr)
    Assert.Equals "x-1", actual_arr(0)
    Assert.Equals "x-2", actual_arr(1)
    Assert.Equals "y-1", actual_arr(2)
    Assert.Equals "y-2", actual_arr(3)
End Sub

Public Sub Test_ReplaceMulti_EmptyReplacementArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ReplaceMulti("abc", "a", EmptyStringArray())

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ReplaceMulti", Err.Source
    Err.Clear
    Assert.IsTrue IsEmptyArray(actual_arr)
End Sub

Public Sub Test_ReplaceMulti_EmptyFindString_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = ReplaceMulti("abc", "", "x")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue IsEmptyArray(actual_arr)
End Sub

' -----------------------------------------------------------------------------
' EscapeLineSeparator
' -----------------------------------------------------------------------------

Public Sub Test_EscapeLineSeparator_CRLFAndBackslash_ReturnsEscapedText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_text As String
    src_text = "A" & vbCrLf & "B" & vbCr & "C" & vbLf & "D\"

    ' Act
    Dim actual_text As String
    actual_text = EscapeLineSeparator(src_text)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A\r\nB\rC\nD\\", actual_text
End Sub

Public Sub Test_EscapeLineSeparator_EmptyEscapeChar_UsesDefaultBackslash(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_text As String
    src_text = "A" & vbLf & "B\"

    ' Act
    Dim actual_text As String
    actual_text = EscapeLineSeparator(src_text, "")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A\nB\\", actual_text
End Sub

Public Sub Test_EscapeLineSeparator_OneCharEscape_ReturnsEscapedText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_text As String
    src_text = "A" & vbLf & "B|" & vbCr & "C"

    ' Act
    Dim actual_text As String
    actual_text = EscapeLineSeparator(src_text, "|")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A|nB|||rC", actual_text
End Sub

Public Sub Test_EscapeLineSeparator_MultiCharEscape_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = EscapeLineSeparator("A" & vbLf & "B", "%%")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.Equals "", actual_text
End Sub

' -----------------------------------------------------------------------------
' UnescapeLineSeparator
' -----------------------------------------------------------------------------

Public Sub Test_UnescapeLineSeparator_EscapedText_ReturnsOriginalText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_text As String
    expected_text = "A" & vbCrLf & "B" & vbCr & "C" & vbLf & "D\"

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("A\r\nB\rC\nD\\")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_text, actual_text
End Sub

Public Sub Test_UnescapeLineSeparator_EmptyEscapeChar_UsesDefaultBackslash(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("A\nB", "")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A" & vbLf & "B", actual_text
End Sub

Public Sub Test_UnescapeLineSeparator_OneCharEscape_ReturnsOriginalText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim expected_text As String
    expected_text = "A" & vbLf & "B|" & vbCr & "C"

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("A|nB|||rC", "|")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_text, actual_text
End Sub

Public Sub Test_UnescapeLineSeparator_MultiCharEscape_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("A%%nB", "%%")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.Equals "", actual_text
End Sub

Public Sub Test_UnescapeLineSeparator_UnknownEscape_DropsEscapeChar(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("A\tB")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "AtB", actual_text
End Sub

Public Sub Test_UnescapeLineSeparator_TrailingEscape_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = UnescapeLineSeparator("abc\")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 512, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.Equals "", actual_text
End Sub

' -----------------------------------------------------------------------------
' SplitByLineSeparator
' -----------------------------------------------------------------------------

Public Sub Test_SplitByLineSeparator_MixedSeparators_ReturnsLines(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr() As String
    actual_arr = SplitByLineSeparator("A" & vbCrLf & "B" & vbCr & "C" & vbLf & "D")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 3, UBound(actual_arr)
    Assert.Equals "A", actual_arr(0)
    Assert.Equals "B", actual_arr(1)
    Assert.Equals "C", actual_arr(2)
    Assert.Equals "D", actual_arr(3)
End Sub

' -----------------------------------------------------------------------------
' UnifyLineSeparator
' -----------------------------------------------------------------------------

Public Sub Test_UnifyLineSeparator_MixedSeparators_ReturnsSpecifiedSeparator(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_text As String
    actual_text = UnifyLineSeparator("A" & vbCrLf & "B" & vbCr & "C" & vbLf & "D", vbCrLf)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D", actual_text
End Sub

' -----------------------------------------------------------------------------
' JoinStringList
' -----------------------------------------------------------------------------

Public Sub Test_JoinStringList_PassedStringList_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_lst As ObjectList
    Set src_lst = New ObjectList
    Call src_lst.Add("alpha")
    Call src_lst.Add("beta")
    Call src_lst.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringList(src_lst, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

' -----------------------------------------------------------------------------
' JoinStringSet
' -----------------------------------------------------------------------------

Public Sub Test_JoinStringSet_PassedStringSet_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_set As ObjectSet
    Set src_set = New ObjectSet
    Call src_set.Add("alpha")
    Call src_set.Add("beta")
    Call src_set.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringSet(src_set, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

' -----------------------------------------------------------------------------
' SplitMessage
' -----------------------------------------------------------------------------

Public Sub Test_SplitMessage_LongLine_SplitsByByteSize(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr As Variant
    actual_arr = SplitMessage("abcdef", PageSize:=4)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "abcd", actual_arr(0)
    Assert.Equals "ef", actual_arr(1)
End Sub

Public Sub Test_SplitMessage_PageSizeLessThanFour_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Call SplitMessage("abcd", PageSize:=0)
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function SplitMessage", Err.Source
    Err.Clear

    Call SplitMessage("abcd", PageSize:=-1)
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function SplitMessage", Err.Source
    Err.Clear

    Call SplitMessage("abcd", PageSize:=1)
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function SplitMessage", Err.Source
    Err.Clear

    Call SplitMessage("abcd", PageSize:=3)
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function SplitMessage", Err.Source
End Sub

Public Sub Test_SplitMessage_MultibyteLine_SplitsByByteSize(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr As Variant
    actual_arr = SplitMessage(ChrW$(&H3042) & ChrW$(&H3044) & ChrW$(&H3046), PageSize:=4)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals ChrW$(&H3042) & ChrW$(&H3044), actual_arr(0)
    Assert.Equals ChrW$(&H3046), actual_arr(1)
End Sub

Public Sub Test_SplitMessage_EmptyMessage_ReturnsSingleEmptyPage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_arr As Variant
    actual_arr = SplitMessage("")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 0, UBound(actual_arr)
    Assert.Equals "", actual_arr(0)
End Sub

Public Sub Test_SplitMessage_BlankLinesOnly_ReturnsSinglePage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim crlf_arr As Variant
    crlf_arr = SplitMessage(vbCrLf)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(crlf_arr)
    Assert.EqualsNumeric 0, UBound(crlf_arr)
    Assert.Equals vbCrLf, crlf_arr(0)

    ' Act
    Err.Clear
    Dim lf_arr As Variant
    lf_arr = SplitMessage(vbLf & vbLf)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(lf_arr)
    Assert.EqualsNumeric 0, UBound(lf_arr)
    Assert.Equals vbCrLf & vbCrLf, lf_arr(0)
End Sub

Public Sub Test_SplitMessage_BlankLinesAroundContent_PreservesLineBreaks(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim trailing_arr As Variant
    trailing_arr = SplitMessage("A" & vbCrLf)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(trailing_arr)
    Assert.EqualsNumeric 0, UBound(trailing_arr)
    Assert.Equals "A" & vbCrLf, trailing_arr(0)

    ' Act
    Err.Clear
    Dim leading_arr As Variant
    leading_arr = SplitMessage(vbCrLf & "A")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(leading_arr)
    Assert.EqualsNumeric 0, UBound(leading_arr)
    Assert.Equals vbCrLf & "A", leading_arr(0)
End Sub

' -----------------------------------------------------------------------------
' Strip
' -----------------------------------------------------------------------------

Public Sub Test_Strip_WhitespaceOptions_ReturnsTrimmedText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "abc", Strip(vbTab & " abc " & vbCrLf)
    Assert.Equals "Å@abc", Strip("Å@abc ", RemoveFullWidthSpace:=False)
    Assert.Equals "abc", Strip("Å@abc ", RemoveFullWidthSpace:=True)
    Assert.Equals " abc", Strip(" abc ", IgnoreHead:=True)
    Assert.Equals "abc ", Strip(" abc ", IgnoreTail:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' StartsWith
' -----------------------------------------------------------------------------

Public Sub Test_StartsWith_Prefix_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue StartsWith("abcdef", "abc")
    Assert.IsFalse StartsWith("abcdef", "bcd")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_StartsWith_EmptySearchString_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue StartsWith("abcdef", "")
    Assert.IsTrue StartsWith("", "")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' EndsWith
' -----------------------------------------------------------------------------

Public Sub Test_EndsWith_Suffix_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue EndsWith("abcdef", "def")
    Assert.IsFalse EndsWith("abcdef", "de")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_EndsWith_EmptySearchString_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue EndsWith("abcdef", "")
    Assert.IsTrue EndsWith("", "")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsQuotedWith
' -----------------------------------------------------------------------------

Public Sub Test_IsQuotedWith_Brackets_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsQuotedWith("[abc]", "[", "]")
    Assert.IsFalse IsQuotedWith("[abc)", "[", "]")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsQuotedWith_EmptyQuoteString_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsQuotedWith("abc", "")
    Assert.IsFalse IsQuotedWith("abc]", "", "]")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsQuotedWith_EmptyEndString_TreatsAsOmitted(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsQuotedWith("'abc'", "'", "")
    Assert.IsFalse IsQuotedWith("'abc]", "'", "")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' MaxLng
' -----------------------------------------------------------------------------

Public Sub Test_MaxLng_MultipleValues_ReturnsMaximum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 5, MaxLng(1, 5, 3, 2)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' MaxDbl
' -----------------------------------------------------------------------------

Public Sub Test_MaxDbl_MultipleValues_ReturnsMaximum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 5.5, MaxDbl(1.5, 5.5, 3.5, 2.5)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' MinLng
' -----------------------------------------------------------------------------

Public Sub Test_MinLng_MultipleValues_ReturnsMinimum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric -4, MinLng(1, -4, 3, 2)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' MinDbl
' -----------------------------------------------------------------------------

Public Sub Test_MinDbl_MultipleValues_ReturnsMinimum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric -4.5, MinDbl(1.5, -4.5, 3.5, 2.5)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' LongToBin
' -----------------------------------------------------------------------------

Public Sub Test_LongToBin_PositiveAndNegative_ReturnsBinaryText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "00000000000000000000000000000101", LongToBin(5)
    Assert.Equals "11111111111111111111111111111111", LongToBin(-1)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' BitLeft
' -----------------------------------------------------------------------------

Public Sub Test_BitLeft_PositiveValue_ReturnsShiftedValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 8, BitLeft(1, 3)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_BitLeft_BoundaryShiftCounts_ReturnsExpectedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim min_shift As Long
    min_shift = &H80000000

    ' Act / Assert
    Assert.EqualsNumeric 0, BitLeft(1, 32)
    Assert.EqualsNumeric 0, BitLeft(1, -32)
    Assert.EqualsNumeric 0, BitLeft(1, min_shift)
    Assert.EqualsNumeric 1, BitLeft(&H80000000, -31)
    Assert.Equals "10000000000000000000000000000000", LongToBin(BitLeft(1, 31))
    Assert.Equals "10000000000000000000000000000000", LongToBin(BitLeft(&HFFFFFFFF, 31))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' BitRight
' -----------------------------------------------------------------------------

Public Sub Test_BitRight_LogicalAndArithmetic_ReturnsShiftedValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 1, BitRight(8, 3)
    Assert.EqualsNumeric -4, BitRight(-8, 1, Arithmetic:=True)
    Assert.EqualsNumeric 2147483644#, BitRight(-8, 1)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_BitRight_BoundaryShiftCounts_ReturnsExpectedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim min_shift As Long
    min_shift = &H80000000

    ' Act / Assert
    Assert.EqualsNumeric 1, BitRight(&H80000000, 31)
    Assert.EqualsNumeric 0, BitRight(&H80000000, 32)
    Assert.EqualsNumeric 0, BitRight(&HFFFFFFFF, 32)
    Assert.EqualsNumeric 0, BitRight(1, -32)
    Assert.EqualsNumeric 0, BitRight(1, min_shift)
    Assert.Equals "10000000000000000000000000000000", LongToBin(BitRight(1, -31))
    Assert.Equals "11111111111111111111111111111111", LongToBin(BitRight(&H80000000, 31, Arithmetic:=True))
    Assert.Equals "11111111111111111111111111111111", LongToBin(BitRight(&H80000000, 32, Arithmetic:=True))
    Assert.EqualsNumeric 0, BitRight(&HFFFFFFFF, -32, Arithmetic:=True)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' CompareAsUnsignedLong
' -----------------------------------------------------------------------------

Public Sub Test_CompareAsUnsignedLong_BoundaryValues_ReturnsUnsignedOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 1, CompareAsUnsignedLong(-1, 1)
    Assert.EqualsNumeric -1, CompareAsUnsignedLong(1, -1)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsLessThanUnsignedLong
' -----------------------------------------------------------------------------

Public Sub Test_IsLessThanUnsignedLong_BoundaryValues_ReturnsUnsignedResult(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsLessThanUnsignedLong(1, -1)
    Assert.IsFalse IsLessThanUnsignedLong(-1, 1)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' AddUnsignedLong
' -----------------------------------------------------------------------------

Public Sub Test_AddUnsignedLong_BasicValues_ReturnsSum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 3, AddUnsignedLong(1, 2)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_AddUnsignedLong_BoundaryValues_ReturnsUnsignedSum(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "10000000000000000000000000000000", LongToBin(AddUnsignedLong(&H7FFFFFFF, 1))
    Assert.Equals "11111111111111111111111111111111", LongToBin(AddUnsignedLong(&H80000000, &H7FFFFFFF))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_AddUnsignedLong_Overflow_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As Long
    actual_value = AddUnsignedLong(&HFFFFFFFF, 1)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function AddUnsignedLong", Err.Source
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' SubtractUnsignedLong
' -----------------------------------------------------------------------------

Public Sub Test_SubtractUnsignedLong_BasicValues_ReturnsDifference(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 2, SubtractUnsignedLong(3, 1)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SubtractUnsignedLong_BoundaryValues_ReturnsUnsignedDifference(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.EqualsNumeric 2147483647#, SubtractUnsignedLong(&H80000000, 1)
    Assert.EqualsNumeric 1, SubtractUnsignedLong(&H80000000, &H7FFFFFFF)
    Assert.EqualsNumeric 2147483647#, SubtractUnsignedLong(&HFFFFFFFF, &H80000000)
    Assert.Equals "11111111111111111111111111111111", LongToBin(SubtractUnsignedLong(&HFFFFFFFF, 0))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SubtractUnsignedLong_NegativeUnsignedResult_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As Long
    actual_value = SubtractUnsignedLong(0, 1)

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function SubtractUnsignedLong", Err.Source
    Err.Clear
End Sub

' -----------------------------------------------------------------------------
' IsInteger
' -----------------------------------------------------------------------------

Public Sub Test_IsInteger_BasicValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsInteger(123)
    Assert.IsFalse IsInteger(123.5)
    Assert.IsFalse IsInteger("abc")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsInteger_StaleErr_ReturnsTrueForInteger(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Err.Clear
    Err.Raise 5, "Test_Lib_Common", "stale error"
    Dim actual_value As Boolean
    actual_value = IsInteger(123)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsInteger_BoundaryValues_ReturnsFalseWithoutOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsInteger("-32768")
    Assert.IsTrue IsInteger("32767")
    Assert.IsFalse IsInteger("-32769")
    Assert.IsFalse IsInteger("32768")
    Assert.IsFalse IsInteger("1E+20")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsLong
' -----------------------------------------------------------------------------

Public Sub Test_IsLong_BasicValues_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsLong(123456)
    Assert.IsFalse IsLong(123456.5)
    Assert.IsFalse IsLong("abc")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsLong_StaleErr_ReturnsTrueForLong(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Err.Clear
    Err.Raise 5, "Test_Lib_Common", "stale error"
    Dim actual_value As Boolean
    actual_value = IsLong(123456)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsLong_BoundaryValues_ReturnsFalseWithoutOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsLong("-2147483648")
    Assert.IsTrue IsLong("2147483647")
    Assert.IsFalse IsLong("-2147483649")
    Assert.IsFalse IsLong("2147483648")
    Assert.IsFalse IsLong("1E+20")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' GetExcelFileFormat
' -----------------------------------------------------------------------------

Public Sub Test_GetExcelFileFormat_UppercaseXlsm_ReturnsMacroEnabledWorkbook(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.XLSM")

    ' Assert
    Assert.EqualsNumeric xlOpenXMLWorkbookMacroEnabled, actual_value
End Sub

Public Sub Test_GetExcelFileFormat_MixedCaseCsv_ReturnsCsv(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.Csv")

    ' Assert
    Assert.EqualsNumeric xlCSV, actual_value
End Sub

Public Sub Test_GetExcelFileFormat_UnknownExtension_ReturnsDefaultWorkbook(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.unknown")

    ' Assert
    Assert.EqualsNumeric xlOpenXMLWorkbook, actual_value
End Sub

Public Sub Test_GetExcelFileFormat_DotFileName_ReturnsDefaultWorkbook(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat(".xlsm")

    ' Assert
    Assert.EqualsNumeric xlOpenXMLWorkbook, actual_value
End Sub

' -----------------------------------------------------------------------------
' IsMultiRange
' -----------------------------------------------------------------------------

Public Sub Test_IsMultiRange_MultiRange_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsMultiRange("A1")
    Assert.IsFalse IsMultiRange("A1:B2")
    Assert.IsTrue IsMultiRange("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsArea
' -----------------------------------------------------------------------------

Public Sub Test_IsArea_AreaAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsArea("A1")
    Assert.IsTrue IsArea("A1:B1")
    Assert.IsTrue IsArea("A1:A2")
    Assert.IsTrue IsArea("A1:B2")
    Assert.IsFalse IsArea("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsCell
' -----------------------------------------------------------------------------

Public Sub Test_IsCell_CellAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsCell("A1")
    Assert.IsFalse IsCell("A1:B1")
    Assert.IsFalse IsCell("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsEntireRow
' -----------------------------------------------------------------------------

Public Sub Test_IsEntireRow_RowAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsEntireRow("A1")
    Assert.IsTrue IsEntireRow("1:1")
    Assert.IsFalse IsEntireRow("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsEntireColumn
' -----------------------------------------------------------------------------

Public Sub Test_IsEntireColumn_ColumnAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsEntireColumn("A1")
    Assert.IsTrue IsEntireColumn("A:A")
    Assert.IsFalse IsEntireColumn("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsOneRow
' -----------------------------------------------------------------------------

Public Sub Test_IsOneRow_CellRowAreaAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsOneRow("A1")
    Assert.IsTrue IsOneRow("A1:B1")
    Assert.IsFalse IsOneRow("A1:A2")
    Assert.IsFalse IsOneRow("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsOneColumn
' -----------------------------------------------------------------------------

Public Sub Test_IsOneColumn_CellColumnAreaAndMultiRange_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsOneColumn("A1")
    Assert.IsFalse IsOneColumn("A1:B1")
    Assert.IsTrue IsOneColumn("A1:A2")
    Assert.IsFalse IsOneColumn("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsOneRowArea
' -----------------------------------------------------------------------------

Public Sub Test_IsOneRowArea_RowAndArea_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsOneRowArea("A1")
    Assert.IsTrue IsOneRowArea("A1:B1")
    Assert.IsFalse IsOneRowArea("A1:A2")
    Assert.IsFalse IsOneRowArea("A1:B2")
    Assert.IsTrue IsOneRowArea("1:1")
    Assert.IsFalse IsOneRowArea("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' IsOneColumnArea
' -----------------------------------------------------------------------------

Public Sub Test_IsOneColumnArea_ColumnAndArea_ReturnsExpected(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsOneColumnArea("A1")
    Assert.IsFalse IsOneColumnArea("A1:B1")
    Assert.IsTrue IsOneColumnArea("A1:A2")
    Assert.IsFalse IsOneColumnArea("A1:B2")
    Assert.IsTrue IsOneColumnArea("A:A")
    Assert.IsFalse IsOneColumnArea("A1,B2")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' RangeAddress
' -----------------------------------------------------------------------------

Public Sub Test_RangeAddress_ColumnRangeMixedAbsoluteFlags_UsesFinishColumnFlag(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartColumn:=2, _
            FinishColumn:=4, _
            IsAbsoluteStartColumn:=True, _
            IsAbsoluteFinishColumn:=False, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "$B:E", actual_value
End Sub

Public Sub Test_RangeAddress_A1AbsoluteCell_DoesNotRequireReferenceIndexes(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=1, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "$A$1", actual_value
End Sub

Public Sub Test_RangeAddress_ColumnRangeMixedAbsoluteFlags_CanMakeOnlyFinishColumnAbsolute(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartColumn:=2, _
            FinishColumn:=4, _
            IsAbsoluteStartColumn:=False, _
            IsAbsoluteFinishColumn:=True, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:$D", actual_value
End Sub

Public Sub Test_RangeAddress_A1MaxColumn_ReturnsXFD(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "XFD1", actual_value
End Sub

Public Sub Test_RangeAddress_A1ColumnOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX + 1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1RelativeColumnOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1MaxRow_ReturnsMaxRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX, _
            StartColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A1048576", actual_value
End Sub

Public Sub Test_RangeAddress_A1RowOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX + 1, _
            StartColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1RelativeRowOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX, _
            StartColumn:=1, _
            ReferenceRow:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteMaxRow_ReturnsMaxRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX, _
            StartColumn:=1, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "R1048576C1", actual_value
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteRowZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=0, _
            StartColumn:=1, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteRowOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX + 1, _
            StartColumn:=1, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteMaxColumn_ReturnsMaxColumn(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "R1C16384", actual_value
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteColumnZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=0, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteColumnOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX + 1, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteStartColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_R1C1RelativeOverSheetLimit_ReturnsRelativeAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX + 1, _
            StartColumn:=G_COL_MAX + 1, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "R[1048577]C[16385]", actual_value
End Sub

Public Sub Test_RangeAddress_R1C1AbsoluteWholeRowAndColumn_ReturnsOmittedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim row_address As String
    row_address = RangeAddress( _
            StartRow:=1, _
            FinishRow:=2, _
            IsAbsoluteStartRow:=True, _
            IsAbsoluteFinishRow:=True, _
            AddressType:="R1C1")

    Dim col_address As String
    col_address = RangeAddress( _
            StartColumn:=1, _
            FinishColumn:=2, _
            IsAbsoluteStartColumn:=True, _
            IsAbsoluteFinishColumn:=True, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "R1:R2", row_address
    Assert.Equals "C1:C2", col_address
End Sub

Public Sub Test_RangeAddress_R1C1RelativeWholeRowAndColumn_ReturnsOmittedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim row_address As String
    row_address = RangeAddress( _
            StartRow:=1, _
            FinishRow:=2, _
            AddressType:="R1C1")

    Dim col_address As String
    col_address = RangeAddress( _
            StartColumn:=1, _
            FinishColumn:=2, _
            AddressType:="R1C1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "R[1]:R[2]", row_address
    Assert.Equals "C[1]:C[2]", col_address
End Sub

Public Sub Test_RangeAddress_BookOnly_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=1, _
            BookName:="Book.xlsx")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
    Err.Clear
    Assert.Equals "", actual_value
End Sub

' -----------------------------------------------------------------------------
' ExcelBookAndSheetAddress
' -----------------------------------------------------------------------------

Public Sub Test_ExcelBookAndSheetAddress_BookAndSheet_ReturnsQualifiedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "[Book.xlsx]Sheet1!", ExcelBookAndSheetAddress("Book.xlsx", "Sheet1")
    Assert.Equals "'[Book.xlsx]Sheet 1'!", ExcelBookAndSheetAddress("Book.xlsx", "Sheet 1")
    Assert.Equals "Sheet1!", ExcelBookAndSheetAddress(SheetName:="Sheet1")
    Assert.Equals "", ExcelBookAndSheetAddress()
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ExcelBookAndSheetAddress_BookOnly_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelBookAndSheetAddress(BookName:="Book.xlsx")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelBookAndSheetAddress", Err.Source
    Err.Clear
    Assert.Equals "", actual_value
End Sub

Public Sub Test_ExcelBookAndSheetAddress_UnquotedSheetNames_ReturnsUnquotedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "Data.1!", ExcelBookAndSheetAddress(SheetName:="Data.1")
    Assert.Equals "InputCheck!", ExcelBookAndSheetAddress(SheetName:="InputCheck")
    Assert.Equals "Input_Check!", ExcelBookAndSheetAddress(SheetName:="Input_Check")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ExcelBookAndSheetAddress_SheetNamesRequiringQuote_ReturnsQuotedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "'Input-Check'!", ExcelBookAndSheetAddress(SheetName:="Input-Check")
    Assert.Equals "'2026.05'!", ExcelBookAndSheetAddress(SheetName:="2026.05")
    Assert.Equals "'A1'!", ExcelBookAndSheetAddress(SheetName:="A1")
    Assert.Equals "'R1C1'!", ExcelBookAndSheetAddress(SheetName:="R1C1")
    Assert.Equals "'O''Brien'!", ExcelBookAndSheetAddress(SheetName:="O'Brien")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ExcelBookAndSheetAddress_BookAndQuotedSheet_ReturnsQuotedAddress(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "'[Book.xlsx]Input-Check'!", ExcelBookAndSheetAddress("Book.xlsx", "Input-Check")
    Assert.Equals "'[Book.xlsx]2026.05'!", ExcelBookAndSheetAddress("Book.xlsx", "2026.05")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ExcelA1ColumnAddress
' -----------------------------------------------------------------------------

Public Sub Test_ExcelA1ColumnAddress_MaxColumn_ReturnsXFD(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(G_COL_MAX)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "XFD", actual_value
End Sub

Public Sub Test_ExcelA1ColumnAddress_BoundaryColumns_ReturnsA1ColumnName(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "A", ExcelA1ColumnAddress(1)
    Assert.Equals "Z", ExcelA1ColumnAddress(26)
    Assert.Equals "AA", ExcelA1ColumnAddress(27)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_ExcelA1ColumnAddress_Zero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelA1ColumnAddress", Err.Source
End Sub

Public Sub Test_ExcelA1ColumnAddress_OverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(G_COL_MAX + 1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelA1ColumnAddress", Err.Source
End Sub

' -----------------------------------------------------------------------------
' SplitExcelAddress
' -----------------------------------------------------------------------------

Public Sub Test_SplitExcelAddress_QuotedWorkbookSheet_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'[Book.xlsm]Data Sheet'!$A$1:$B$2")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "Book.xlsm", book_name
    Assert.Equals "Data Sheet", sheet_name
    Assert.Equals "$A$1:$B$2", cell_address
End Sub

Public Sub Test_SplitExcelAddress_QuotedSheetWithExclamation_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'Input!Check'!A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "", book_name
    Assert.Equals "Input!Check", sheet_name
    Assert.Equals "A1", cell_address
End Sub

Public Sub Test_SplitExcelAddress_QuotedSheetWithEscapedApostrophe_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'O''Brien'!$A$1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "", book_name
    Assert.Equals "O'Brien", sheet_name
    Assert.Equals "$A$1", cell_address
End Sub

Public Sub Test_SplitExcelAddress_InvalidAddress_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "Sheet1!!A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' SplitA1RangeAddress
' -----------------------------------------------------------------------------

Public Sub Test_SplitA1RangeAddress_Cell_ReturnsIndexes(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$B$2")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, start_row
    Assert.EqualsNumeric 2, start_col
    Assert.EqualsNumeric 2, finish_row
    Assert.EqualsNumeric 2, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_RowRange_ReturnsOmittedColumns(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "1:3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, start_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_col
    Assert.EqualsNumeric 3, finish_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_ColumnRange_ReturnsOmittedRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A:C")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_row
    Assert.EqualsNumeric 1, start_col
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_row
    Assert.EqualsNumeric 3, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_AbsoluteRowRange_ReturnsOmittedColumns(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$1:$3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, start_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_col
    Assert.EqualsNumeric 3, finish_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_AbsoluteColumnRange_ReturnsOmittedRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$A:$C")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_row
    Assert.EqualsNumeric 1, start_col
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_row
    Assert.EqualsNumeric 3, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_DollarInsideColumn_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A$B1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
End Sub

Public Sub Test_SplitA1RangeAddress_DoubleColumnDollar_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$$A$1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
End Sub

Public Sub Test_SplitA1RangeAddress_TrailingDollar_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A1$")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
End Sub

Public Sub Test_SplitA1RangeAddress_MultiRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A1,B2")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

' -----------------------------------------------------------------------------
' ConvertRangeToStringList
' -----------------------------------------------------------------------------

Public Sub Test_ConvertRangeToStringList_UsesTextAndPreservesInjectedService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim ws_stub As WorksheetServiceTestDouble
    Set ws_stub = New WorksheetServiceTestDouble
    Set WsSrv = ws_stub
    Dim target_range As WorksheetRangeBounds
    Set target_range = New_RangeBounds(Row:=1, Column:=1, FinishRow:=1, FinishColumn:=2, Sheet:="TestSheet", Book:=ThisWorkbook.Name)

    Dim first_cell As WorksheetRangeBounds
    Set first_cell = New_RangeBounds(Row:=1, Column:=1, Sheet:="TestSheet", Book:=ThisWorkbook.Name)
    Dim first_value As Variant
    first_value = Array("Text-1", "@")
    Call ws_stub.Store.SetReturn("ReadCell", first_value(0), first_cell, True)
    Call ws_stub.Store.SetOutput("ReadCell", "Expression", first_value(0), first_cell, True)
    Call ws_stub.Store.SetOutput("ReadCell", "NumberFormat", first_value(1), first_cell, True)

    Dim second_cell As WorksheetRangeBounds
    Set second_cell = New_RangeBounds(Row:=1, Column:=2, Sheet:="TestSheet", Book:=ThisWorkbook.Name)
    Dim second_value As Variant
    second_value = Array("Text-2", "@")
    Call ws_stub.Store.SetReturn("ReadCell", second_value(0), second_cell, True)
    Call ws_stub.Store.SetOutput("ReadCell", "Expression", second_value(0), second_cell, True)
    Call ws_stub.Store.SetOutput("ReadCell", "NumberFormat", second_value(1), second_cell, True)

    ' Act
    Dim actual_list As ObjectList
    Set actual_list = ConvertRangeToStringList(target_range, GetText:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "WorksheetServiceTestDouble", WsSrv
    Assert.EqualsNumeric 2, actual_list.Count
    Assert.Equals "Text-1", actual_list.Item(0)
    Assert.Equals "Text-2", actual_list.Item(1)

    Call InitializeCommonService(Force:=True)
End Sub
