Attribute VB_Name = "Test_FileSystemService"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the FileSystemService class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_GetAbsolutePath_RelativePathWithUninitializedWbSrv_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath("relative.txt")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class FileSystemService.GetAbsolutePath", actual_err_source
    Assert.Equals "InitializeCommonService has not been called.", actual_err_desc
End Sub

Public Sub Test_GetAbsolutePath_AbsolutePathWithUninitializedWbSrv_ReturnsAbsolutePath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim expected_path As String
    expected_path = "C:\codex_abs_test.txt"

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath(expected_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

Public Sub Test_GetAbsolutePath_AbsolutePathWithEnvironmentVariableAndUninitializedWbSrv_ReturnsExpandedAbsolutePath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")

    Dim input_path As String
    input_path = "%TEMP%" & G_FS_PATH_SEP & "codex_env_path_test.txt"

    Dim expanded_path As String
    expanded_path = shell_obj.ExpandEnvironmentStrings(input_path)

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath(input_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals expanded_path, actual_path
End Sub

Public Sub Test_GetAbsolutePath_RelativePathAfterEnvironmentVariableExpansion_MakesAbsoluteFromThisWorkbook(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    Call wb_stub.Store.SetReturn("GetThisWorkbookDirectoryPath", "C:\InjectedBase")
    Set WbSrv = wb_stub

    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")

    Dim env_values As Object
    Set env_values = shell_obj.Environment("Process")
    env_values("CODEX_XLS_COMMON_TEST_REL_PATH") = "EnvChild"

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath("%CODEX_XLS_COMMON_TEST_REL_PATH%\File.txt")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call env_values.Remove("CODEX_XLS_COMMON_TEST_REL_PATH")
    Call InitializeCommonService(Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "C:\InjectedBase\EnvChild\File.txt", actual_path
End Sub

Public Sub Test_GetAbsolutePath_EnvironmentVariableSyntaxInUrl_NormalizesAsUrlWithoutExpansion(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim input_path As String
    input_path = "https://example.com/%TEMP%/sub/../file.txt"

    Dim expected_path As String
    expected_path = GetAbsolutePathFromParent(input_path, input_path)

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath(input_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals expected_path, actual_path
    Assert.IsTrue InStr(1, actual_path, "%TEMP%", vbBinaryCompare) > 0
End Sub

Public Sub Test_GetAbsolutePath_LocalAbsolutePathContainingPercent_FollowsOsEnvironmentVariableExpansion(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing

    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")

    Dim env_values As Object
    Set env_values = shell_obj.Environment("Process")
    Call env_values.Remove("CODEX_XLS_COMMON_TEST_UNDEFINED_PATH")
    Err.Clear

    Dim input_path As String
    input_path = "C:\100%\%CODEX_XLS_COMMON_TEST_UNDEFINED_PATH%\File.txt"

    Dim expected_path As String
    expected_path = shell_obj.ExpandEnvironmentStrings(input_path)

    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetAbsolutePath(input_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

' ----------------------------------------------------------------------------
' GetTemporaryDirectoryPath
' ----------------------------------------------------------------------------

Public Sub Test_GetTemporaryDirectoryPath_NormalCall_ReturnsTemporaryDirectoryPath(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    ' Act
    Dim actual_path As String
    actual_path = file_srv.GetTemporaryDirectoryPath()

    ' Assert
    Assert.IsTrue IsAbsolutePath(actual_path)
    Assert.IsTrue file_srv.IsDirectory(actual_path)
    Assert.IsFalse Right$(actual_path, 1) = G_FS_PATH_SEP
End Sub

' ----------------------------------------------------------------------------
' GetFileSize
' ----------------------------------------------------------------------------

Public Sub Test_GetFileSize_ExistingFile_ReturnsByteCount(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim expected_size As Double
    expected_size = CDbl(FileLen(ThisWorkbook.FullName))

    ' Act
    Dim actual_size As Double
    actual_size = file_srv.GetFileSize(ThisWorkbook.FullName)

    ' Assert
    Assert.EqualsNumeric expected_size, actual_size
End Sub

Public Sub Test_GetFileSize_DirectoryPath_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim target_path As String
    target_path = ThisWorkbook.Path

    ' Act
    Dim actual_size As Double
    actual_size = file_srv.GetFileSize(target_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class FileSystemService.GetFileSize", actual_err_source
    Assert.Equals "The target for getting the file size is not a file. (" & target_path & ")", actual_err_desc
End Sub

Public Sub Test_GetFileSize_MissingPath_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim file_srv As FileSystemService
    Set file_srv = New FileSystemService

    Dim target_path As String
    target_path = JoinPath(ThisWorkbook.Path, "__get_file_size_missing_file__.tmp")
    If file_srv.PathExists(target_path) Then
        Assert.IsFalse True, "missing file test path exists"
        Exit Sub
    End If

    ' Act
    Dim actual_size As Double
    actual_size = file_srv.GetFileSize(target_path)

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class FileSystemService.GetFileSize", actual_err_source
    Assert.Equals "The file to get the file size from does not exist. (" & target_path & ")", actual_err_desc
End Sub

Public Sub Test_GetFileSize_ThroughInterface_ReturnsByteCount(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_srv As IFileSystemService
    Set fs_srv = New FileSystemService

    Dim expected_size As Double
    expected_size = CDbl(FileLen(ThisWorkbook.FullName))

    ' Act
    Dim actual_size As Double
    actual_size = fs_srv.GetFileSize(ThisWorkbook.FullName)

    ' Assert
    Assert.EqualsNumeric expected_size, actual_size
End Sub
