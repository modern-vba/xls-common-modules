VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFileSystemService

' #############################################################################
'!
'! @brief
'! Test double for the service that manages file system access.
'!
' #############################################################################

'* Stub settings and history of successfully completed calls.
'* Members that need return values and depend on external state or parent resources are strict stubs; members where a conservative default is natural, such as existence checks, use safe defaults.
'* Operations without return values are spy only, and unregistered strict stub errors and explicit errors from SetError are not recorded in call history.
'* spy only records call arguments in the order of the original method definition.
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* Strict stub for GetAbsolutePath. Register the return value explicitly in Store.
Public Function GetAbsolutePath(ByVal PathLikeString As String) As String
    Call Store.RaiseIfError("GetAbsolutePath", PathLikeString)

    Dim result_value As String
    result_value = Store.GetReturn("GetAbsolutePath", PathLikeString)
    Call Store.RecordCall("GetAbsolutePath", PathLikeString)

    GetAbsolutePath = result_value
End Function
Private Function IFileSystemService_GetAbsolutePath(ByVal PathLikeString As String) As String
    IFileSystemService_GetAbsolutePath = GetAbsolutePath(PathLikeString)
End Function

'* Strict stub for GetTemporaryDirectoryPath. Register the return value explicitly in Store.
Public Function GetTemporaryDirectoryPath() As String
    Call Store.RaiseIfError("GetTemporaryDirectoryPath")

    Dim result_value As String
    result_value = Store.GetReturn("GetTemporaryDirectoryPath")
    Call Store.RecordCall("GetTemporaryDirectoryPath")

    GetTemporaryDirectoryPath = result_value
End Function
Private Function IFileSystemService_GetTemporaryDirectoryPath() As String
    IFileSystemService_GetTemporaryDirectoryPath = GetTemporaryDirectoryPath()
End Function

'* Strict stub for CreateTemporaryDirectory. Register the return value explicitly in Store.
Public Function CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
    Call Store.RaiseIfError("CreateTemporaryDirectory", NamePrefix)

    Dim result_value As String
    result_value = Store.GetReturn("CreateTemporaryDirectory", NamePrefix)
    Call Store.RecordCall("CreateTemporaryDirectory", NamePrefix)

    CreateTemporaryDirectory = result_value
End Function
Private Function IFileSystemService_CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
    IFileSystemService_CreateTemporaryDirectory = CreateTemporaryDirectory(NamePrefix:=NamePrefix)
End Function

'* Strict stub for GetFileList. Because it does not assume the parent directory exists, explicitly register even an empty array in Store.
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Call Store.RaiseIfError("GetFileList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim result_value() As String
    result_value = Store.GetReturn("GetFileList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)
    Call Store.RecordCall("GetFileList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    GetFileList = result_value
End Function
Private Function IFileSystemService_GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetFileList = GetFileList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

'* Strict stub for GetDirectoryList. Because it does not assume the parent directory exists, explicitly register even an empty array in Store.
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Call Store.RaiseIfError("GetDirectoryList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim result_value() As String
    result_value = Store.GetReturn("GetDirectoryList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)
    Call Store.RecordCall("GetDirectoryList", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    GetDirectoryList = result_value
End Function
Private Function IFileSystemService_GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetDirectoryList = GetDirectoryList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

'* Safe default for PathExists. Returns False when no value is registered.
Public Function PathExists(ByVal PathLikeString As String) As Boolean
    Call Store.RaiseIfError("PathExists", PathLikeString)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("PathExists", False, PathLikeString)
    Call Store.RecordCall("PathExists", PathLikeString)

    PathExists = result_value
End Function
Private Function IFileSystemService_PathExists(ByVal PathLikeString As String) As Boolean
    IFileSystemService_PathExists = PathExists(PathLikeString)
End Function

'* Safe default for IsFile. Returns False when no value is registered.
Public Function IsFile(ByVal FilePath As String) As Boolean
    Call Store.RaiseIfError("IsFile", FilePath)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("IsFile", False, FilePath)
    Call Store.RecordCall("IsFile", FilePath)

    IsFile = result_value
End Function
Private Function IFileSystemService_IsFile(ByVal FilePath As String) As Boolean
    IFileSystemService_IsFile = IsFile(FilePath)
End Function

'* Safe default for IsDirectory. Returns False when no value is registered.
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
    Call Store.RaiseIfError("IsDirectory", DirectoryPath)

    Dim result_value As Boolean
    result_value = Store.GetReturnOrDefault("IsDirectory", False, DirectoryPath)
    Call Store.RecordCall("IsDirectory", DirectoryPath)

    IsDirectory = result_value
End Function
Private Function IFileSystemService_IsDirectory(ByVal DirectoryPath As String) As Boolean
    IFileSystemService_IsDirectory = IsDirectory(DirectoryPath)
End Function

'* Strict stub for GetLastModified. Register the return value explicitly in Store.
Public Function GetLastModified(ByVal PathLikeString As String) As Date
    Call Store.RaiseIfError("GetLastModified", PathLikeString)

    Dim result_value As Date
    result_value = Store.GetReturn("GetLastModified", PathLikeString)
    Call Store.RecordCall("GetLastModified", PathLikeString)

    GetLastModified = result_value
End Function
Private Function IFileSystemService_GetLastModified(ByVal PathLikeString As String) As Date
    IFileSystemService_GetLastModified = GetLastModified(PathLikeString)
End Function

'* Strict stub for GetFileSize. Register the return value explicitly in Store.
Public Function GetFileSize(ByVal FilePath As String) As Double
    Call Store.RaiseIfError("GetFileSize", FilePath)

    Dim result_value As Double
    result_value = Store.GetReturn("GetFileSize", FilePath)
    Call Store.RecordCall("GetFileSize", FilePath)

    GetFileSize = result_value
End Function
Private Function IFileSystemService_GetFileSize(ByVal FilePath As String) As Double
    IFileSystemService_GetFileSize = GetFileSize(FilePath)
End Function

'* Strict stub for CreateDirectory. Register the return value explicitly in Store.
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    Call Store.RaiseIfError("CreateDirectory", DirectoryPath, Force, Recursive)

    Dim result_value As Boolean
    result_value = Store.GetReturn("CreateDirectory", DirectoryPath, Force, Recursive)
    Call Store.RecordCall("CreateDirectory", DirectoryPath, Force, Recursive)

    CreateDirectory = result_value
End Function
Private Function IFileSystemService_CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    IFileSystemService_CreateDirectory = CreateDirectory(DirectoryPath, Force:=Force, Recursive:=Recursive)
End Function

'* spy only for MoveDirectory. Records the call; inject failures with Store.SetError.
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call Store.RaiseIfError("MoveDirectory", SourcePath, DestinationPath, Force)
    Call Store.RecordCall("MoveDirectory", SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

'* spy only for CopyDirectory. Records the call; inject failures with Store.SetError.
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call Store.RaiseIfError("CopyDirectory", SourcePath, DestinationPath, Force)
    Call Store.RecordCall("CopyDirectory", SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

'* Strict stub for RemoveDirectory. Register the return value explicitly in Store.
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    Call Store.RaiseIfError("RemoveDirectory", DirectoryPath, Force)

    Dim result_value As Boolean
    result_value = Store.GetReturn("RemoveDirectory", DirectoryPath, Force)
    Call Store.RecordCall("RemoveDirectory", DirectoryPath, Force)

    RemoveDirectory = result_value
End Function
Private Function IFileSystemService_RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveDirectory = RemoveDirectory(DirectoryPath, Force:=Force)
End Function

'* spy only for MoveFile. Records the call; inject failures with Store.SetError.
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call Store.RaiseIfError("MoveFile", SourcePath, DestinationPath, Force)
    Call Store.RecordCall("MoveFile", SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveFile(SourcePath, DestinationPath, Force:=Force)
End Sub

'* spy only for CopyFile. Records the call; inject failures with Store.SetError.
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call Store.RaiseIfError("CopyFile", SourcePath, DestinationPath, Force)
    Call Store.RecordCall("CopyFile", SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyFile(SourcePath, DestinationPath, Force:=Force)
End Sub

'* Strict stub for RemoveFile. Register the return value explicitly in Store.
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    Call Store.RaiseIfError("RemoveFile", FilePath, Force)

    Dim result_value As Boolean
    result_value = Store.GetReturn("RemoveFile", FilePath, Force)
    Call Store.RecordCall("RemoveFile", FilePath, Force)

    RemoveFile = result_value
End Function
Private Function IFileSystemService_RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveFile = RemoveFile(FilePath, Force:=Force)
End Function

'* Strict stub for GetNewestFile. Because it does not assume the parent directory exists, explicitly register an empty string as well when there is no match.
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    Call Store.RaiseIfError("GetNewestFile", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim result_value As String
    result_value = Store.GetReturn("GetNewestFile", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)
    Call Store.RecordCall("GetNewestFile", DirectoryPath, CatchRegExp, IgnoreRegExp, IgnoreCase)

    GetNewestFile = result_value
End Function
Private Function IFileSystemService_GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    IFileSystemService_GetNewestFile = GetNewestFile(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

'* Strict stub for CreateBackupFile. Register the return value explicitly in Store.
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String

    Call Store.RaiseIfError("CreateBackupFile", SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound)

    Dim result_value As String
    result_value = Store.GetReturn("CreateBackupFile", SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound)
    Call Store.RecordCall("CreateBackupFile", SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound)

    CreateBackupFile = result_value
End Function
Private Function IFileSystemService_CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String

    IFileSystemService_CreateBackupFile = CreateBackupFile( _
            SourceFilePath:=SourceFilePath, _
            DestinationDirectoryPath:=DestinationDirectoryPath, _
            BackupSuffix:=BackupSuffix, _
            Force:=Force, Move:=Move, FailIfSourceNotFound:=FailIfSourceNotFound)
End Function
