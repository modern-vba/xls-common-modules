VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSystemService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for a service that manages file system access.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IFileSystemService", Description:="This class is intended for interfaces."
End Sub

'* Converts the specified path to an absolute path.
'*
'* @param PathLikeString Path string.
'* @return Absolute path.
'*
'* @details
'* Converts the specified path to an absolute path. If the path is already absolute, returns a normalized absolute path.
'* Relative paths are resolved against the result of WorkbookService.GetThisWorkbookDirectoryPath.
Public Function GetAbsolutePath(ByVal PathLikeString As String) As String
End Function

'* Gets the path of the OS user temporary folder.
'*
'* @return Absolute path of the OS user temporary folder.
'*
'* @details
'* Returns the absolute Windows path of the user temporary folder reported by the OS. A trailing path separator is not appended.
Public Function GetTemporaryDirectoryPath() As String
End Function

'* Creates a temporary directory directly under the OS user temporary folder.
'*
'* @param NamePrefix [Optional] String to prepend to the temporary directory name.
'* @return Absolute path of the created temporary directory.
'*
'* @details
'* Creates a directory with a low-collision name directly under the OS user temporary folder and returns the created absolute Windows path.
'* NamePrefix is not transformed. If it produces an invalid folder name, a creation error is raised.
Public Function CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
End Function

'* Gets a list of absolute paths for files in the specified directory.
'*
'* @param DirectoryPath Directory path.
'* @param CatchRegExp [Optional] Regular expression for file names to include in the list.
'* @param IgnoreRegExp [Optional] Regular expression for file names to exclude from the list.
'* @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
'* @return Array of absolute paths.
'*
'* @details
'* Gets an array of absolute paths for files in the specified directory. Returns a 0-based empty array when there are no matches.
'* Return order follows FileSystemObject / OS enumeration order and is not sorted.
'* If an individual file is deleted or renamed during enumeration and getting its name or path raises a file-or-path-not-found error, that item is skipped and enumeration continues.
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
End Function

'* Gets a list of absolute paths for subdirectories in the specified directory.
'*
'* @param DirectoryPath Directory path.
'* @param CatchRegExp [Optional] Regular expression for subdirectory names to include in the list.
'* @param IgnoreRegExp [Optional] Regular expression for subdirectory names to exclude from the list.
'* @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
'* @return Array of absolute paths.
'*
'* @details
'* Gets an array of absolute paths for subdirectories in the specified directory. Returns a 0-based empty array when there are no matches.
'* Return order follows FileSystemObject / OS enumeration order and is not sorted.
'* If an individual subdirectory is deleted or renamed during enumeration and getting its name or path raises a file-or-path-not-found error, that item is skipped and enumeration continues.
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
End Function

'* Checks whether an entity exists at the specified path.
'*
'* @param PathLikeString Path string.
'* @return Whether an entity exists.
'*
'* @details
'* Checks whether an entity exists at the specified path. Returns True when a file or directory exists.
Public Function PathExists(ByVal PathLikeString As String) As Boolean
End Function

'* Checks whether a file exists at the specified path.
'*
'* @param FilePath File path.
'* @return Whether the entity exists.
'*
'* @details
'* Checks whether a file exists at the specified path.
Public Function IsFile(ByVal FilePath As String) As Boolean
End Function

'* Checks whether a directory exists at the specified path.
'*
'* @param DirectoryPath Directory path.
'* @return Whether the entity exists.
'*
'* @details
'* Checks whether a directory exists at the specified path.
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
End Function

'* Gets the last modified date and time.
'*
'* @param PathLikeString Path string.
'* @return Last modified date and time.
'*
'* @details
'* Gets the last modified date and time.
Public Function GetLastModified(ByVal PathLikeString As String) As Date
End Function

'* Gets the file size in bytes.
'*
'* @param FilePath File path.
'* @return File size, in bytes.
'*
'* @details
'* Gets the size of the specified file in bytes. Raises an error if the target does not exist or is not a file.
Public Function GetFileSize(ByVal FilePath As String) As Double
End Function

'* Creates a directory.
'*
'* @param DirectoryPath Directory path.
'* @param Force [Optional] Whether to overwrite (delete) when something already exists at the destination path.
'* @param Recursive [Optional] Whether to create parent directories recursively when they do not exist.
'* @return Whether a directory was actually created.
'*
'* @details
'* Creates a directory at the specified path. If a directory already exists, does nothing and returns False.
'* Returns True when a directory was actually created.
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
End Function

'* Moves a directory.
'*
'* @param SourcePath Path of the source directory.
'* @param DestinationPath Path of the destination directory.
'* @param Force [Optional] Whether to overwrite when something already exists at the destination path.
'*
'* @details
'* Moves the directory to the specified path. If DestinationPath ends with a path separator, moves it under that directory using the current directory name.
'* When Force:=True, copies to the destination using the same overwrite rules as CopyDirectory, then deletes the source. Extra files under an existing destination directory are not deleted.
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* Copies a directory.
'*
'* @param SourcePath Path of the source directory.
'* @param DestinationPath Path of the destination directory.
'* @param Force [Optional] Whether to overwrite when something already exists at the destination path.
'*
'* @details
'* Copies the directory to the specified path. If DestinationPath ends with a path separator, copies it under that directory using the current directory name.
'* When Force:=True, existing files are overwritten. Extra files under an existing destination directory are not deleted.
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* Deletes a directory.
'*
'* @param DirectoryPath Directory path.
'* @param Force [Optional] Whether to delete even read-only attributes.
'* @return Whether a directory was actually deleted.
'*
'* @details
'* Deletes the directory at the specified path. If the directory does not already exist, does nothing and returns False.
'* Returns True when a directory was actually deleted.
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
End Function

'* Moves a file.
'*
'* @param SourcePath Path of the source file.
'* @param DestinationPath Path of the destination file.
'* @param Force [Optional] Whether to overwrite when something already exists at the destination path.
'*
'* @details
'* Moves the file to the specified path. If DestinationPath ends with a path separator, moves it under that directory using the current file name.
'* When Force:=True, copies to the destination using the same overwrite rules as CopyFile, then deletes the source.
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* Copies a file.
'*
'* @param SourcePath Path of the source file.
'* @param DestinationPath Path of the destination file.
'* @param Force [Optional] Whether to overwrite when something already exists at the destination path.
'*
'* @details
'* Copies the file to the specified path. If DestinationPath ends with a path separator, copies it under that directory using the current file name.
'* When Force:=True, existing files are overwritten.
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
End Sub

'* Deletes a file.
'*
'* @param FilePath File path.
'* @param Force [Optional] Whether to delete even read-only attributes.
'* @return Whether a file was actually deleted.
'*
'* @details
'* Deletes the file at the specified path. If the file does not already exist, does nothing and returns False.
'* Returns True when a file was actually deleted.
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
End Function

'* Gets the newest file among files with the specified name directly under the specified folder.
'*
'* @param DirectoryPath Path of the folder to search.
'* @param CatchRegExp [Optional] Regular expression for file names to check.
'* @param IgnoreRegExp [Optional] Regular expression for file names not to check.
'* @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
'* @return Full path of the newest file. Empty string if there are no matching files.
'*
'* @details
'* Searches files in the specified folder and, among files that match the specified pattern,
'* returns the full path of the newest one. Traversal order follows FileSystemObject / OS enumeration order, and ordering for identical timestamps is not guaranteed.
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
End Function

'* Creates a backup file.
'*
'* @param SourceFilePath [Optional] File path to back up (default is the current workbook path).
'* @param DestinationDirectoryPath [Optional] Destination directory path for the backup (default is the target file directory).
'* @param BackupSuffix [Optional] Suffix to append to the backup file name (default is a timestamp).
'* @param Force [Optional] Whether to forcibly overwrite an existing backup file.
'* @param Move [Optional] Whether to move the source file during backup (True means move).
'* @param FailIfSourceNotFound [Optional] Whether to raise an error when the target file does not exist (True means error).
'* @return Full path of the created backup file.
'*
'* @details
'* Creates a backup of the specified file. The suffix is appended to the original file name for the destination backup file name.
'* If an existing file is present, handling is based on the `Force` setting.
'* `Force:=True` overwrites the existing backup file.
'* If `Force:=False` and the default suffix from `BackupSuffix:=""` collides, returns the existing backup path without comparing contents or moving with `Move:=True`.
'* If `Force:=False` and an explicit `BackupSuffix` collides, an error is raised.
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String
End Function
