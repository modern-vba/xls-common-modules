Attribute VB_Name = "Test_FileSystemServiceTestDoubl"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the FileSystemServiceTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

' ----------------------------------------------------------------------------
' GetAbsolutePath
' ----------------------------------------------------------------------------

Public Sub Test_GetAbsolutePath_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_path As Variant
    expected_path = "C:\Base\Sub\File.txt"
    Call fs_stub.Store.SetReturn("GetAbsolutePath", expected_path, "Sub\File.txt")

    Dim fs_srv As IFileSystemService
    Set fs_srv = fs_stub

    ' Act
    Dim actual_path As String
    actual_path = fs_srv.GetAbsolutePath("Sub\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals CStr(expected_path), actual_path
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetAbsolutePath", "Sub\File.txt")
End Sub

Public Sub Test_GetAbsolutePath_NoValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual_path As String
    actual_path = fs_stub.GetAbsolutePath("Sub\File.txt")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetAbsolutePath", "Sub\File.txt")
End Sub

' ----------------------------------------------------------------------------
' GetTemporaryDirectoryPath
' ----------------------------------------------------------------------------

Public Sub Test_GetTemporaryDirectoryPath_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_path As String
    expected_path = "C:\Temp"
    Call fs_stub.Store.SetReturn("GetTemporaryDirectoryPath", expected_path)

    Dim fs_srv As IFileSystemService
    Set fs_srv = fs_stub

    ' Act
    Dim actual_path As String
    actual_path = fs_srv.GetTemporaryDirectoryPath()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetTemporaryDirectoryPath")
End Sub

Public Sub Test_GetTemporaryDirectoryPath_NoValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual_path As String
    actual_path = fs_stub.GetTemporaryDirectoryPath()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetTemporaryDirectoryPath")
End Sub

' ----------------------------------------------------------------------------
' CreateTemporaryDirectory
' ----------------------------------------------------------------------------

Public Sub Test_CreateTemporaryDirectory_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_path As String
    expected_path = "C:\Temp\codex_tmp123.tmp"
    Call fs_stub.Store.SetReturn("CreateTemporaryDirectory", expected_path, "codex_")

    Dim fs_srv As IFileSystemService
    Set fs_srv = fs_stub

    ' Act
    Dim actual_path As String
    actual_path = fs_srv.CreateTemporaryDirectory("codex_")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("CreateTemporaryDirectory", "codex_")
End Sub

Public Sub Test_CreateTemporaryDirectory_NoValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual_path As String
    actual_path = fs_stub.CreateTemporaryDirectory()

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateTemporaryDirectory", "")
End Sub

' ----------------------------------------------------------------------------
' GetFileList
' ----------------------------------------------------------------------------

Public Sub Test_GetFileList_WithSetValue_ReturnsStubArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_arr(0 To 1) As String
    expected_arr(0) = "C:\Temp\a.txt"
    expected_arr(1) = "C:\Temp\b.txt"
    Call fs_stub.Store.SetReturn("GetFileList", expected_arr, "C:\Temp", "^.*\.txt$", "", True)

    ' Act
    Dim actual_arr() As String
    actual_arr = fs_stub.GetFileList("C:\Temp", "^.*\.txt$")

    ' Assert
    Assert.EqualsArray expected_arr, actual_arr
End Sub

Public Sub Test_GetFileList_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_arr() As String
    actual_arr = fs_stub.GetFileList("C:\Temp", "^.*\.txt$")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetFileList", "C:\Temp", "^.*\.txt$", "", True)
End Sub

' ----------------------------------------------------------------------------
' GetDirectoryList
' ----------------------------------------------------------------------------

Public Sub Test_GetDirectoryList_WithArgs_ReturnsStubArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_dirs(0 To 1) As String
    expected_dirs(0) = "C:\Temp\Dir1"
    expected_dirs(1) = "C:\Temp\Dir2"
    Call fs_stub.Store.SetReturn("GetDirectoryList", expected_dirs, "C:\Temp", "", "", True)

    ' Act
    Dim actual_arr() As String
    actual_arr = fs_stub.GetDirectoryList("C:\Temp")

    ' Assert
    Assert.EqualsArray expected_dirs, actual_arr
End Sub

Public Sub Test_GetDirectoryList_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_arr() As String
    actual_arr = fs_stub.GetDirectoryList("C:\Temp")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetDirectoryList", "C:\Temp", "", "", True)
End Sub

' ----------------------------------------------------------------------------
' PathExists
' ----------------------------------------------------------------------------

Public Sub Test_PathExists_SetValueTrue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetReturn("PathExists", True, "C:\ExistsPath")

    ' Act
    Dim actual As Boolean
    actual = fs_stub.PathExists("C:\ExistsPath")

    ' Assert
    Assert.IsTrue actual
End Sub

Public Sub Test_PathExists_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual As Boolean
    actual = fs_stub.PathExists("C:\NoSuchPath")

    ' Assert
    Assert.IsFalse actual
End Sub

' ----------------------------------------------------------------------------
' IsFile
' ----------------------------------------------------------------------------

Public Sub Test_IsFile_WithDictionary_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetReturn("IsFile", True, "C:\Temp\SomeFile.txt")

    ' Act
    Dim actual As Boolean
    actual = fs_stub.IsFile("C:\Temp\SomeFile.txt")

    ' Assert
    Assert.IsTrue actual
End Sub

' ----------------------------------------------------------------------------
' IsDirectory
' ----------------------------------------------------------------------------

Public Sub Test_IsDirectory_WithDictionary_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetReturn("IsDirectory", True, "C:\Temp\KnownDir")

    ' Act
    Dim actual As Boolean
    actual = fs_stub.IsDirectory("C:\Temp\KnownDir")

    ' Assert
    Assert.IsTrue actual
End Sub

' ----------------------------------------------------------------------------
' GetLastModified
' ----------------------------------------------------------------------------

Public Sub Test_GetLastModified_WithSetDate_ReturnsDate(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim test_date As Date
    test_date = #11/1/2023 2:05:00 PM#

    Call fs_stub.Store.SetReturn("GetLastModified", test_date, "C:\fileA.txt")

    ' Act
    Dim actual_date As Date
    actual_date = fs_stub.GetLastModified("C:\fileA.txt")

    ' Assert
    Assert.Equals test_date, actual_date
End Sub

' ----------------------------------------------------------------------------
' GetFileSize
' ----------------------------------------------------------------------------

Public Sub Test_GetFileSize_WithStubValue_ReturnsDouble(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetReturn("GetFileSize", 1234#, "C:\file.bin")

    Dim fs_srv As IFileSystemService
    Set fs_srv = fs_stub

    ' Act
    Dim actual_size As Double
    actual_size = fs_srv.GetFileSize("C:\file.bin")

    ' Assert
    Assert.EqualsNumeric 1234#, actual_size
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileSize", "C:\file.bin")
End Sub

Public Sub Test_GetFileSize_NoValue_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual_size As Double
    actual_size = fs_stub.GetFileSize("C:\missing.bin")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetFileSize", "C:\missing.bin")
End Sub

Public Sub Test_GetFileSize_SetError_RaisesInjectedError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetError("GetFileSize", vbObjectError + 2049, "Class Injected", "Injected error.", "C:\bad.bin")

    ' Act
    Dim actual_size As Double
    actual_size = fs_stub.GetFileSize("C:\bad.bin")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 2049, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class Injected", Err.Source
    Assert.Equals "Injected error.", Err.Description
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetFileSize", "C:\bad.bin")
End Sub

' ----------------------------------------------------------------------------
' CreateDirectory
' ----------------------------------------------------------------------------

Public Sub Test_CreateDirectory_RecordsArgsAndReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Register the CreateDirectory return value in Store.
    Call fs_stub.Store.SetReturn("CreateDirectory", False, "C:\MakeDir", True, True)

    ' Act
    Dim actual_result As Boolean
    actual_result = fs_stub.CreateDirectory("C:\MakeDir", True, True)

    ' Assert
    ' 1) The return value is False (from the dictionary).
    Assert.Equals False, actual_result

    ' 2) The latest return value is recorded in Store.
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("CreateDirectory", "C:\MakeDir", True, True)
    Assert.EqualsNumeric 3, stored_val.ArgumentCount
End Sub

Public Sub Test_CreateDirectory_NoEntryInValues_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    Dim actual_result As Boolean
    actual_result = fs_stub.CreateDirectory("D:\NewDir")

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TestDoubleBehaviorStore.")
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateDirectory", "D:\NewDir", False, False)
End Sub

' ----------------------------------------------------------------------------
' MoveDirectory
' ----------------------------------------------------------------------------

Public Sub Test_MoveDirectory_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    fs_stub.MoveDirectory "C:\SrcFolder", "D:\DestFolder", True

    ' Assert
    ' (SourcePath, DestPath, Force) -> True
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("MoveDirectory", "C:\SrcFolder", "D:\DestFolder", True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' CopyDirectory
' ----------------------------------------------------------------------------

Public Sub Test_CopyDirectory_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    fs_stub.CopyDirectory "C:\FolderA", "C:\FolderB", False

    ' Assert
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("CopyDirectory", "C:\FolderA", "C:\FolderB", False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' RemoveDirectory
' ----------------------------------------------------------------------------

Public Sub Test_RemoveDirectory_WithForceTrue_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Set dictionary => True
    Call fs_stub.Store.SetReturn("RemoveDirectory", True, "C:\TargetDir", True)

    ' Act
    Dim actual As Boolean
    actual = fs_stub.RemoveDirectory("C:\TargetDir", True)

    ' Assert
    Assert.Equals True, actual

    ' spied call
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("RemoveDirectory", "C:\TargetDir", True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_RemoveDirectory_NoRegInValues_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual As Boolean
    actual = fs_stub.RemoveDirectory("D:\Unknown", False)

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("RemoveDirectory", "D:\Unknown", False)
End Sub

' ----------------------------------------------------------------------------
' MoveFile
' ----------------------------------------------------------------------------

Public Sub Test_MoveFile_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    fs_stub.MoveFile "C:\Temp\fileA.txt", "C:\Temp\dest\FileA_moved.txt", True

    ' Assert
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("MoveFile", "C:\Temp\fileA.txt", "C:\Temp\dest\FileA_moved.txt", True)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' CopyFile
' ----------------------------------------------------------------------------

Public Sub Test_CopyFile_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Act
    fs_stub.CopyFile "D:\Source.docx", "E:\Backup\Source.docx", False

    ' Assert
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("CopyFile", "D:\Source.docx", "E:\Backup\Source.docx", False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

' ----------------------------------------------------------------------------
' RemoveFile
' ----------------------------------------------------------------------------

Public Sub Test_RemoveFile_WithDictionaryReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Call fs_stub.Store.SetReturn("RemoveFile", True, "C:\Data\fileX.txt", False)

    ' Act
    Dim actual_res As Boolean
    actual_res = fs_stub.RemoveFile("C:\Data\fileX.txt", False)

    ' Assert
    Assert.IsTrue actual_res

    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("RemoveFile", "C:\Data\fileX.txt", False)
    Assert.IsTrue 0 < stored_val.ArgumentCount
End Sub

Public Sub Test_RemoveFile_NotInDictionary_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_res As Boolean
    actual_res = fs_stub.RemoveFile("C:\NoSuchFile.log")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("RemoveFile", "C:\NoSuchFile.log", False)
End Sub

Public Sub Test_RemoveFile_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_res As Boolean
    actual_res = fs_stub.RemoveFile("C:\NoSuchFile.log")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("RemoveFile", "C:\NoSuchFile.log", False)
End Sub
' ----------------------------------------------------------------------------
' GetNewestFile
' ----------------------------------------------------------------------------

Public Sub Test_GetNewestFile_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_path As String
    expected_path = "C:\Folder\Newest.docx"

    ' (Directory, catch, ignore)
    Call fs_stub.Store.SetReturn("GetNewestFile", expected_path, "C:\Folder", ".*\.docx", "", True)

    ' Act
    Dim actual_path As String
    actual_path = fs_stub.GetNewestFile("C:\Folder", ".*\.docx")

    ' Assert
    Assert.Equals expected_path, actual_path
End Sub

Public Sub Test_GetNewestFile_NoEntry_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_path As String
    actual_path = fs_stub.GetNewestFile("Z:\None", ".*\.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("GetNewestFile", "Z:\None", ".*\.txt", "", True)
End Sub

Public Sub Test_Calls_ReadOnlyMethods_RecordCalls(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim expected_files(0 To 0) As String
    expected_files(0) = "C:\Temp\File1.txt"
    Call fs_stub.Store.SetReturn("GetFileList", expected_files, "C:\Temp", "^.*\.txt$", "", True)

    Dim expected_dirs(0 To 0) As String
    expected_dirs(0) = "C:\Temp\Sub1"
    Call fs_stub.Store.SetReturn("GetDirectoryList", expected_dirs, "C:\Temp", "", "^ignore$", False)

    Call fs_stub.Store.SetReturn("PathExists", True, "C:\ExistsPath")
    Call fs_stub.Store.SetReturn("IsFile", True, "C:\Temp\SomeFile.txt")
    Call fs_stub.Store.SetReturn("IsDirectory", True, "C:\Temp\KnownDir")
    Call fs_stub.Store.SetReturn("GetLastModified", DateSerial(2026, 5, 22), "C:\fileA.txt")
    Call fs_stub.Store.SetReturn("GetFileSize", 2048#, "C:\fileA.txt")
    Call fs_stub.Store.SetReturn("GetNewestFile", "C:\Folder\Newest.docx", "C:\Folder", ".*\.docx", "", True)

    ' Act
    Dim actual_files() As String
    actual_files = fs_stub.GetFileList("C:\Temp", "^.*\.txt$")

    Dim actual_dirs() As String
    actual_dirs = fs_stub.GetDirectoryList("C:\Temp", IgnoreRegExp:="^ignore$", IgnoreCase:=False)

    Dim actual_exists As Boolean
    actual_exists = fs_stub.PathExists("C:\ExistsPath")

    Dim actual_is_file As Boolean
    actual_is_file = fs_stub.IsFile("C:\Temp\SomeFile.txt")

    Dim actual_is_dir As Boolean
    actual_is_dir = fs_stub.IsDirectory("C:\Temp\KnownDir")

    Dim actual_modified As Date
    actual_modified = fs_stub.GetLastModified("C:\fileA.txt")

    Dim actual_size As Double
    actual_size = fs_stub.GetFileSize("C:\fileA.txt")

    Dim actual_newest As String
    actual_newest = fs_stub.GetNewestFile("C:\Folder", ".*\.docx")

    ' Assert
    Assert.Equals "C:\Temp\File1.txt", actual_files(0)
    Assert.Equals "C:\Temp\Sub1", actual_dirs(0)
    Assert.IsTrue actual_exists
    Assert.IsTrue actual_is_file
    Assert.IsTrue actual_is_dir
    Assert.Equals DateSerial(2026, 5, 22), actual_modified
    Assert.EqualsNumeric 2048#, actual_size
    Assert.Equals "C:\Folder\Newest.docx", actual_newest

    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileList", "C:\Temp", "^.*\.txt$", "", True)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetDirectoryList", "C:\Temp", "", "^ignore$", False)
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("PathExists", "C:\ExistsPath")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsFile", "C:\Temp\SomeFile.txt")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("IsDirectory", "C:\Temp\KnownDir")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetLastModified", "C:\fileA.txt")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetFileSize", "C:\fileA.txt")
    Assert.EqualsNumeric 1, fs_stub.Store.GetCallCount("GetNewestFile", "C:\Folder", ".*\.docx", "", True)
End Sub

' ----------------------------------------------------------------------------
' CreateBackupFile
' ----------------------------------------------------------------------------

Public Sub Test_CreateBackupFile_RecordsArgsAndReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    ' Register the CreateBackupFile return value in Store.
    Call fs_stub.Store.SetReturn("CreateBackupFile", "MyBackup", "C:\source.txt", "D:\backups", "_v2", True, True, True)

    ' Act
    Dim actual_str As String
    actual_str = fs_stub.CreateBackupFile("C:\source.txt", "D:\backups", "_v2", True, True, True)

    ' Assert
    ' 1) Return value.
    Assert.Equals "MyBackup", actual_str

    ' 2) The latest return value is recorded in Store.
    Dim stored_val As TestDoubleCallRecord
    Set stored_val = fs_stub.Store.GetLatestCall("CreateBackupFile", "C:\source.txt", "D:\backups", "_v2", True, True, True)
    Assert.EqualsNumeric 6, stored_val.ArgumentCount
End Sub

Public Sub Test_CreateBackupFile_Unregistered_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_path As String
    actual_path = fs_stub.CreateBackupFile("C:\noRegFile.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateBackupFile", "C:\noRegFile.txt", "", "", False, False, True)
End Sub

Public Sub Test_CreateBackupFile_NotInDictionary_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble

    Dim actual_str As String
    actual_str = fs_stub.CreateBackupFile("C:\noRegFile.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, fs_stub.Store.GetCallCount("CreateBackupFile", "C:\noRegFile.txt", "", "", False, False, True)
End Sub
