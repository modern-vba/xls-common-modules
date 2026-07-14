Attribute VB_Name = "Lib_UnitTest"
Option Explicit
Option Private Module

' #############################################################################
'!
'! @brief
'! Entry point for unit tests.
'!
' #############################################################################

Private Const C_SHEET_NAME As String = "UNIT_TEST_SHEET"
Private Const C_NEW_BOOK As Boolean = False
Private Const C_COL_MOD As Long = 1
Private Const C_COL_SUB As Long = 2
Private Const C_COL_OKNG As Long = 3
Private Const C_COL_DESC As Long = 4
Private Const C_COL_END As Long = 4
Private Const C_COL_BTN As Long = 5
Private Const C_COL_MBTN As Long = 6
Private Const C_SUB_MAIN As String = "UnitTestMain"
Private Const C_RESULT_OK As String = "OK"
Private Const C_RESULT_NG As String = "NG"
Private Const C_RESULT_ERR As String = "ERR"
Private Const C_NO_ASSERTION_MESSAGE As String = "No assertions were executed."
Private Const C_RUNTIME_RUNNER_MODULE_PREFIX As String = "Tmp_UTRUN"
Private Const C_RUNTIME_RUNNER_MODULE_SUFFIX_LENGTH As Long = 22
Private Const C_RUNTIME_RUNNER_FUNCTION As String = "Run"
Private Const C_VBEXT_CT_STDMODULE As Long = 1
Private Const C_COLOR_RESET_BG As Long = &H404040
Private Const C_COLOR_RESET_FG As Long = &HC0C0C0

Private pRuntimeRunnerModuleName As String

'* Entry point for unit tests.
'*
'* @details
'* Entry point for unit tests.
'* Scans all modules in this project, finds sub procedures whose names start with Test_ and whose argument is UnitTestAssert, and runs them.
'*
'* Unit test example:
'* @code
'* Sub Test_SomeFunction(Assert As UnitTestAssert)
'*     ' Arrange
'*     Dim some_obj As SomeObject: t some_obj = New SomeObject
'*
'*     ' Act
'*     Dim actual_value As Long
'*     actual_value = some_obj.SomeFunction()
'*
'*     ' Assert
'*     Assert.Equals 0, actual_value)
'* End Sub
'* @endcode
Public Sub UnitTestMain( _
        Optional ByVal TestModuleName As String = "", _
        Optional ByVal TestSubName As String = "")
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    Call app_state.DisableUpdates(StopEvents:=False)
    Call pRemoveRuntimeRunnerModule
    Call pInitializeRuntimeRunnerModuleName
    Call pEnsureRuntimeRunnerModule

    Dim result_sheet As Worksheet

    ' Get the test number.
    Dim test_idx As Long: test_idx = 0

    If TestModuleName = "" And TestSubName <> "" Then
        Err.Raise vbObjectError + 2, "Module Lib_UnitTest", "TestModuleName is required when TestSubName is specified."
    End If

    If TestModuleName = "" Then
        Dim caller_str As String

        On Error Resume Next
        caller_str = CStr(Application.Caller)
        If Err.Number <> 0 Then
            Err.Clear
            caller_str = ""
        End If
        On Error GoTo ON_ERROR

        If IsNumeric(caller_str) Then
            test_idx = CLng(caller_str)
        Else
            test_idx = 0
        End If
    End If

    ' Prepare the output destination for test results.
    Set result_sheet = pPrepareResultSheet(test_idx)

    ' Run tests.
    If TestModuleName <> "" Then
        Call pRunSelectedTest(result_sheet, TestModuleName, TestSubName)
    ElseIf test_idx = 0 Then
        Call pRunAllTest(result_sheet)
    Else
        Call pRunTestCore(result_sheet, result_sheet.Cells(test_idx, C_COL_MOD), result_sheet.Cells(test_idx, C_COL_SUB), test_idx)
    End If

    Call pRemoveRuntimeRunnerModule
    If Not result_sheet Is Nothing Then Call result_sheet.Activate
    Call app_state.Restore
    On Error GoTo 0
    Exit Sub

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext


    On Error Resume Next
    Call pRemoveRuntimeRunnerModule
    If Not result_sheet Is Nothing Then Call result_sheet.Activate
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

Private Sub pRunAllTest(ByVal ResultSheet As Worksheet)
    ' Get the VBIDE project object.
    Dim vb_proj As Variant 'VBIDE.VBProject
    Set vb_proj = ThisWorkbook.VBProject

    ' Prepare the regular expression used to extract test Sub procedures.
    Dim sub_re As RegExp
    Set sub_re = New RegExp
    sub_re.Pattern = "^\s*(?:Public\s+)?Sub\s+(Test_[^\s(]+)\s*\(\s*(?:ByVal\s+|ByRef\s+)?[^\s,()]+\s+As\s+UnitTestAssert\s*(?:,\s*Optional\s+[^,)]+)*\).*$"

    ' Process all components in the project object.
    Dim row_idx As Long: row_idx = 2
    Dim vb_comp As Variant 'VBIDE.VBComponent
    For Each vb_comp In vb_proj.VBComponents
        If vb_comp.Type = C_VBEXT_CT_STDMODULE Then
            ' Get the module name.
            Dim mod_name As String
            On Error Resume Next
            mod_name = vb_comp.Name
            If Err.Number <> 0 Then
                Debug.Print "<" & row_idx & "> [&H" & Hex(Err.Number) & "] " & Err.Source & " | " & Err.Description
                Err.Clear
                On Error GoTo 0
            Else
                On Error GoTo 0

                'Debug.Print "Search " & mod_name

                ' Get the component's code module.
                Dim vb_comp_code As Variant 'VBIDE.CodeModule
                Set vb_comp_code = vb_comp.CodeModule

                ' Process all lines in the code module.
                Dim line_idx As Long
                line_idx = 1
                Do While line_idx <= vb_comp_code.CountOfLines
                    Dim code_line As String
                    code_line = pReadLogicalLine(vb_comp_code, line_idx)

                    ' Check whether the line matches the regular expression.
                    Dim match_result As MatchCollection
                    Set match_result = sub_re.Execute(code_line)

                    If 0 < match_result.Count Then
                        'Debug.Print "Found " & TestName

                        Dim sub_name As String
                        sub_name = match_result.Item(0).SubMatches(0)

                        ' Run the test.
                        Call pRunTestCore(ResultSheet, mod_name, sub_name, row_idx)

                        ' Add a button.
                        Call AddButton(ResultSheet, row_idx, C_COL_BTN, "Run Again", pBuildWorkbookMacroName(C_SUB_MAIN), row_idx)

                        ' Advance the row.
                        row_idx = row_idx + 1
                    End If
                    line_idx = line_idx + 1
                Loop
            End If
        End If
    Next vb_comp
End Sub

Private Sub pRunSelectedTest(ByVal ResultSheet As Worksheet, ByVal TestModuleName As String, ByVal TestSubName As String)
    Dim vb_comp As Object
    Set vb_comp = pGetSelectedTestModule(TestModuleName)

    Dim sub_re As RegExp
    Set sub_re = New RegExp
    sub_re.Pattern = "^\s*(?:Public\s+)?Sub\s+(Test_[^\s(]+)\s*\(\s*(?:ByVal\s+|ByRef\s+)?[^\s,()]+\s+As\s+UnitTestAssert\s*(?:,\s*Optional\s+[^,)]+)*\).*$"

    Dim found_module_test As Boolean: found_module_test = False
    Dim found_selected_test As Boolean: found_selected_test = False
    Dim row_idx As Long: row_idx = 2

    Dim vb_comp_code As Object
    Set vb_comp_code = vb_comp.CodeModule

    Dim line_idx As Long
    line_idx = 1
    Do While line_idx <= vb_comp_code.CountOfLines
        Dim code_line As String
        code_line = pReadLogicalLine(vb_comp_code, line_idx)

        Dim match_result As MatchCollection
        Set match_result = sub_re.Execute(code_line)

        If 0 < match_result.Count Then
            found_module_test = True

            Dim sub_name As String
            sub_name = match_result.Item(0).SubMatches(0)

            If TestSubName = "" Or StrComp(sub_name, TestSubName, vbTextCompare) = 0 Then
                Call pRunTestCore(ResultSheet, vb_comp.Name, sub_name, row_idx)
                Call AddButton(ResultSheet, row_idx, C_COL_BTN, "Run Again", pBuildWorkbookMacroName(C_SUB_MAIN), row_idx)

                found_selected_test = True
                row_idx = row_idx + 1

                If TestSubName <> "" Then Exit Do
            End If
        End If
        line_idx = line_idx + 1
    Loop

    If Not found_module_test Then
        Err.Raise vbObjectError + 2, "Module Lib_UnitTest", "Discoverable unit tests were not found in module: " & TestModuleName
    End If

    If TestSubName <> "" And Not found_selected_test Then
        Err.Raise vbObjectError + 2, "Module Lib_UnitTest", "Discoverable test procedure was not found: " & TestModuleName & "." & TestSubName
    End If
End Sub

Private Function pGetSelectedTestModule(ByVal TestModuleName As String) As Object
    Dim vb_proj As Object: Set vb_proj = ThisWorkbook.VBProject
    Dim vb_comp As Object

    On Error Resume Next
    Set vb_comp = vb_proj.VBComponents.Item(TestModuleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set vb_comp = Nothing
    End If
    On Error GoTo 0

    If vb_comp Is Nothing Then
        Err.Raise vbObjectError + 2, "Module Lib_UnitTest", "Test module was not found: " & TestModuleName
    End If

    If vb_comp.Type <> C_VBEXT_CT_STDMODULE Then
        Err.Raise vbObjectError + 2, "Module Lib_UnitTest", "Test module must be a standard module: " & TestModuleName
    End If

    Set pGetSelectedTestModule = vb_comp
End Function

Private Sub pRunTestCore(ByVal ResultSheet As Worksheet, ByVal TestModName As String, ByVal TestSubName As String, ByVal RowIndex As Long)
    Dim assert_obj As UnitTestAssert: Set assert_obj = New UnitTestAssert
    Dim run_result As Variant

    On Error GoTo RUNNER_ERROR
    Call pWriteRuntimeRunnerModule(TestModName, TestSubName)
    run_result = Application.Run(pBuildWorkbookMacroName(pRuntimeRunnerModuleName & "." & C_RUNTIME_RUNNER_FUNCTION), assert_obj)
    On Error GoTo 0

    If CBool(run_result(0)) Then
        Call pWriteRuntimeErrorResult(ResultSheet, RowIndex, TestModName, TestSubName, CLng(run_result(1)), CStr(run_result(2)), CStr(run_result(3)))
    Else
        Call pWriteResult(ResultSheet, RowIndex, TestModName, TestSubName, assert_obj)
    End If
    Exit Sub

RUNNER_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Err.Clear
    On Error GoTo 0

    Call pWriteRunnerErrorResult(ResultSheet, RowIndex, TestModName, TestSubName, err_num, err_source, err_desc)
End Sub

Private Sub pRemoveRuntimeRunnerModule()
    Dim vb_proj As Object: Set vb_proj = ThisWorkbook.VBProject
    Dim remove_names As Collection: Set remove_names = New Collection
    Dim vb_comp As Object

    For Each vb_comp In vb_proj.VBComponents
        If vb_comp.Type = C_VBEXT_CT_STDMODULE Then
            If vb_comp.Name = pRuntimeRunnerModuleName Or pIsRuntimeRunnerModuleName(vb_comp.Name) Then
                Call remove_names.Add(vb_comp.Name)
            End If
        End If
    Next vb_comp

    Dim remove_name As Variant
    For Each remove_name In remove_names
        Call vb_proj.VBComponents.Remove(vb_proj.VBComponents.Item(CStr(remove_name)))
    Next remove_name

    pRuntimeRunnerModuleName = ""
End Sub

Private Sub pInitializeRuntimeRunnerModuleName()
    pRuntimeRunnerModuleName = pCreateRuntimeRunnerModuleName()
End Sub

Private Function pEnsureRuntimeRunnerModule() As Object
    Dim vb_proj As Object: Set vb_proj = ThisWorkbook.VBProject

    If pRuntimeRunnerModuleName = "" Then
        Call pInitializeRuntimeRunnerModuleName
    End If

    On Error Resume Next
    Set pEnsureRuntimeRunnerModule = vb_proj.VBComponents.Item(pRuntimeRunnerModuleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set pEnsureRuntimeRunnerModule = Nothing
    End If
    On Error GoTo 0

    If pEnsureRuntimeRunnerModule Is Nothing Then
        Set pEnsureRuntimeRunnerModule = vb_proj.VBComponents.Add(C_VBEXT_CT_STDMODULE)
        pEnsureRuntimeRunnerModule.Name = pRuntimeRunnerModuleName
    End If
End Function

Private Function pPrepareResultSheet(ByVal TestIndex As Long) As Worksheet
    Dim result_sheet As Worksheet
    If C_NEW_BOOK Then
        ' Output to a new workbook.
        Dim result_book As Workbook
        Set result_book = Workbooks.Add

        Set result_sheet = result_book.Worksheets(1)
        Call AddButton(result_sheet, 1, C_COL_MBTN, "Run All", pBuildWorkbookMacroName(C_SUB_MAIN), "Button_UnitTestMain")
    Else
        ' Output to ThisWorkbook.
        If TestIndex = 0 Then
            ' For all-test execution.

            ' Try to create the sheet.
            On Error Resume Next
            Set result_sheet = ThisWorkbook.Worksheets(C_SHEET_NAME)
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            If result_sheet Is Nothing Then
                ' New sheet.
                Set result_sheet = ThisWorkbook.Worksheets.Add()
                result_sheet.Name = C_SHEET_NAME
            Else
                ' Existing sheet.
                result_sheet.Columns(C_COL_OKNG).Interior.Color = C_COLOR_RESET_BG
                result_sheet.Columns(C_COL_OKNG).Font.Color = C_COLOR_RESET_FG
                Call result_sheet.Cells.ClearContents
                Call ClearButton(result_sheet)
            End If
            Call AddButton(result_sheet, 1, C_COL_MBTN, "Run All", pBuildWorkbookMacroName(C_SUB_MAIN), "Button_UnitTestMain")
        Else
            Set result_sheet = ThisWorkbook.Worksheets(C_SHEET_NAME)
            result_sheet.Cells(TestIndex, C_COL_OKNG).Interior.Color = C_COLOR_RESET_BG
            result_sheet.Cells(TestIndex, C_COL_OKNG).Font.Color = C_COLOR_RESET_FG
            Call result_sheet.Cells(TestIndex, C_COL_OKNG).ClearContents
            Call result_sheet.Cells(TestIndex, C_COL_DESC).ClearContents
        End If
    End If

    result_sheet.Cells(1, C_COL_MOD).Value = "Category"
    result_sheet.Cells(1, C_COL_SUB).Value = "Test Item"
    result_sheet.Cells(1, C_COL_OKNG).Value = "Result"
    result_sheet.Cells(1, C_COL_DESC).Value = "Description"
    If TestIndex = 0 Then
        If result_sheet.AutoFilterMode Then
            result_sheet.AutoFilterMode = False
        End If

        Call result_sheet.Range(RangeAddress(StartColumn:=1, FinishColumn:=C_COL_END)).AutoFilter
    End If

    Set pPrepareResultSheet = result_sheet
End Function

Private Function pReadLogicalLine(ByVal CodeModule As Variant, ByRef LineIndex As Long) As String
    Dim logical_line As String

    Do While LineIndex <= CodeModule.CountOfLines
        Dim physical_line As String
        physical_line = CodeModule.Lines(LineIndex, 1)

        If pHasLineContinuation(physical_line) Then
            logical_line = logical_line & pRemoveLineContinuation(physical_line) & " "
            LineIndex = LineIndex + 1
        Else
            logical_line = logical_line & physical_line
            Exit Do
        End If
    Loop

    pReadLogicalLine = logical_line
End Function

Private Function pBuildWorkbookMacroName(ByVal MacroName As String) As String
    pBuildWorkbookMacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & MacroName
End Function

Private Sub pWriteRuntimeRunnerModule(ByVal TestModName As String, ByVal TestSubName As String)
    Dim runner_comp As Object: Set runner_comp = pEnsureRuntimeRunnerModule()
    Dim runner_code As Object: Set runner_code = runner_comp.CodeModule

    If runner_code.CountOfLines > 0 Then
        runner_code.DeleteLines 1, runner_code.CountOfLines
    End If

    runner_code.AddFromString pBuildRuntimeRunnerCode(TestModName, TestSubName)
End Sub

Private Sub pWriteRuntimeErrorResult( _
        ByVal ResultSheet As Worksheet, _
        ByVal RowIndex As Long, _
        ByVal ModuleName As String, _
        ByVal TestSubName As String, _
        ByVal ErrorNumber As Long, _
        ByVal ErrorSource As String, _
        ByVal ErrorDescription As String)

    ResultSheet.Cells(RowIndex, C_COL_MOD).Value = ModuleName
    ResultSheet.Cells(RowIndex, C_COL_SUB).Value = TestSubName
    ResultSheet.Cells(RowIndex, C_COL_DESC).Value = _
            "Runtime error [&H" & Hex(ErrorNumber) & "] @<" & ErrorSource & "> """ & ErrorDescription & """"
    Call pWriteResultStatus(ResultSheet, RowIndex, C_RESULT_ERR, RGB(255, 192, 0))
End Sub

Private Sub pWriteResult(ByVal ResultSheet As Worksheet, ByVal RowIndex As Long, ByVal ModuleName As String, ByVal TestSubName As String, ByVal AssertObject As UnitTestAssert)
    ResultSheet.Cells(RowIndex, C_COL_MOD).Value = ModuleName
    ResultSheet.Cells(RowIndex, C_COL_SUB).Value = TestSubName

    If AssertObject.AssertionCount <= 0 Then
        ResultSheet.Cells(RowIndex, C_COL_DESC).Value = C_NO_ASSERTION_MESSAGE
        Call pWriteResultStatus(ResultSheet, RowIndex, C_RESULT_ERR, RGB(255, 192, 0))
    ElseIf AssertObject.IsFailed Then
        ResultSheet.Cells(RowIndex, C_COL_DESC).Value = AssertObject.ResultMessage
        Call pWriteResultStatus(ResultSheet, RowIndex, C_RESULT_NG, RGB(255, 128, 128))
    Else
        ResultSheet.Cells(RowIndex, C_COL_DESC).Value = AssertObject.ResultMessage
        Call pWriteResultStatus(ResultSheet, RowIndex, C_RESULT_OK, RGB(128, 255, 128))
    End If
End Sub

Private Sub pWriteRunnerErrorResult( _
        ByVal ResultSheet As Worksheet, _
        ByVal RowIndex As Long, _
        ByVal ModuleName As String, _
        ByVal TestSubName As String, _
        ByVal ErrorNumber As Long, _
        ByVal ErrorSource As String, _
        ByVal ErrorDescription As String)

    ResultSheet.Cells(RowIndex, C_COL_MOD).Value = ModuleName
    ResultSheet.Cells(RowIndex, C_COL_SUB).Value = TestSubName
    ResultSheet.Cells(RowIndex, C_COL_DESC).Value = _
            "Runner error [&H" & Hex(ErrorNumber) & "] @<" & ErrorSource & "> """ & ErrorDescription & """"
    Call pWriteResultStatus(ResultSheet, RowIndex, C_RESULT_ERR, RGB(255, 192, 0))
End Sub

Private Function pHasLineContinuation(ByVal CodeLine As String) As Boolean
    Dim trimmed_line As String
    trimmed_line = RTrim$(CodeLine)
    If Len(trimmed_line) < 2 Then Exit Function
    If Right$(trimmed_line, 1) <> "_" Then Exit Function

    Dim previous_char As String
    previous_char = Mid$(trimmed_line, Len(trimmed_line) - 1, 1)
    pHasLineContinuation = (previous_char = " " Or previous_char = vbTab)
End Function

Private Function pRemoveLineContinuation(ByVal CodeLine As String) As String
    Dim trimmed_line As String
    trimmed_line = RTrim$(CodeLine)
    pRemoveLineContinuation = RTrim$(Left$(trimmed_line, Len(trimmed_line) - 1))
End Function

Private Function pCreateRuntimeRunnerModuleName() As String
    Const C_MAX_ATTEMPT As Long = 100

    Randomize

    Dim attempt_idx As Long
    For attempt_idx = 1 To C_MAX_ATTEMPT
        Dim candidate_name As String
        candidate_name = C_RUNTIME_RUNNER_MODULE_PREFIX & pRandomUppercaseString(C_RUNTIME_RUNNER_MODULE_SUFFIX_LENGTH)

        If Not pVBComponentExists(candidate_name) Then
            pCreateRuntimeRunnerModuleName = candidate_name
            Exit Function
        End If
    Next attempt_idx

    Err.Raise vbObjectError + 1, "Module Lib_UnitTest", "Failed to generate the temporary module name for execution."
End Function

Private Function pRandomUppercaseString(ByVal Length As Long) As String
    Dim result_value As String

    Dim char_idx As Long
    For char_idx = 1 To Length
        result_value = result_value & Chr$(Asc("A") + CLng(Int(26 * Rnd)))
    Next char_idx

    pRandomUppercaseString = result_value
End Function

Private Function pVBComponentExists(ByVal ModuleName As String) As Boolean
    Dim vb_proj As Object: Set vb_proj = ThisWorkbook.VBProject
    Dim target_comp As Object

    On Error Resume Next
    Set target_comp = vb_proj.VBComponents.Item(ModuleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set target_comp = Nothing
    End If
    On Error GoTo 0

    pVBComponentExists = Not target_comp Is Nothing
End Function

Private Function pIsRuntimeRunnerModuleName(ByVal ModuleName As String) As Boolean
    Static runtime_runner_name_re As RegExp

    If runtime_runner_name_re Is Nothing Then
        Set runtime_runner_name_re = New RegExp
        runtime_runner_name_re.Pattern = "^" & C_RUNTIME_RUNNER_MODULE_PREFIX & "[A-Z]{" & C_RUNTIME_RUNNER_MODULE_SUFFIX_LENGTH & "}$"
        runtime_runner_name_re.IgnoreCase = False
        runtime_runner_name_re.Global = False
    End If

    pIsRuntimeRunnerModuleName = runtime_runner_name_re.Test(ModuleName)
End Function

Private Function pBuildRuntimeRunnerCode(ByVal TestModName As String, ByVal TestSubName As String) As String
    pBuildRuntimeRunnerCode = _
            "Option Explicit" & vbCrLf & _
            "Option Base 0" & vbCrLf & _
            vbCrLf & _
            "Public Function " & C_RUNTIME_RUNNER_FUNCTION & "(ByVal AssertObject As UnitTestAssert) As Variant" & vbCrLf & _
            "    On Error GoTo ON_ERROR" & vbCrLf & _
            "    Call " & TestModName & "." & TestSubName & "(AssertObject)" & vbCrLf & _
            "    " & C_RUNTIME_RUNNER_FUNCTION & " = Array(False, 0, """", """")" & vbCrLf & _
            "    Exit Function" & vbCrLf & _
            "ON_ERROR:" & vbCrLf & _
            "    " & C_RUNTIME_RUNNER_FUNCTION & " = Array(True, Err.Number, Err.Source, Err.Description)" & vbCrLf & _
            "End Function" & vbCrLf
End Function

Private Sub pWriteResultStatus(ByVal ResultSheet As Worksheet, ByVal RowIndex As Long, ByVal ResultText As String, ByVal BackgroundColor As Long)
    ResultSheet.Cells(RowIndex, C_COL_OKNG).Value = ResultText
    ResultSheet.Cells(RowIndex, C_COL_OKNG).Interior.Color = BackgroundColor
    ResultSheet.Cells(RowIndex, C_COL_OKNG).Font.Color = RGB(64, 64, 64)
End Sub
