VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFileSystemService

' #############################################################################
'!
'! @brief
'! Implementation of the service that manages file system access.
'!
' #############################################################################

'Private pFSO As Object
Private pFSO As Scripting.FileSystemObject

Private Const C_ERR_FILE_NOT_FOUND As Long = 53
Private Const C_ERR_PATH_NOT_FOUND As Long = 76

Private Sub Class_Initialize()
    Set pFSO = CreateObject("Scripting.FileSystemObject")
    'Set pFSO = New Scripting.FileSystemObject
End Sub

'* Converts the specified path to an absolute path.
'*
'* @param PathLikeString Path string.
'* @return Absolute path.
'*
'* @details
'* Converts the specified path to an absolute path. If the path is already absolute, returns a normalized absolute path.
'* For local paths that are not URLs, environment variables are expanded using OS behavior before conversion to an absolute path.
'* Relative paths are resolved against the result of WorkbookService.GetThisWorkbookDirectoryPath.
Public Function GetAbsolutePath(ByVal PathLikeString As String) As String
    Dim target_path As String
    target_path = PathLikeString
    If Not IsUrlPath(target_path) Then target_path = pExpandEnvironmentStrings(target_path)

    If IsAbsolutePath(target_path) Then
        GetAbsolutePath = GetAbsolutePathFromParent(target_path, target_path)
    Else
        Call pRequireWorkbookService("GetAbsolutePath")
        GetAbsolutePath = GetAbsolutePathFromParent(WbSrv.GetThisWorkbookDirectoryPath(), target_path)
    End If
End Function
Private Function IFileSystemService_GetAbsolutePath(ByVal PathLikeString As String) As String
    IFileSystemService_GetAbsolutePath = GetAbsolutePath(PathLikeString)
End Function

Private Function pExpandEnvironmentStrings(ByVal PathLikeString As String) As String
    Dim shell_obj As Object
    Set shell_obj = CreateObject("WScript.Shell")

    pExpandEnvironmentStrings = shell_obj.ExpandEnvironmentStrings(PathLikeString)
End Function

Private Sub pRequireWorkbookService(ByVal MethodName As String)
    If WbSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService" & "." & MethodName, Description:="InitializeCommonService has not been called."
    End If
End Sub

'* Gets the path of the OS user temporary folder.
'*
'* @return Absolute path of the OS user temporary folder.
'*
'* @details
'* Returns the absolute Windows path of the user temporary folder reported by the OS. A trailing path separator is not appended.
Public Function GetTemporaryDirectoryPath() As String
    Dim result_value As String
    result_value = pFSO.GetSpecialFolder(TemporaryFolder).Path

    Do While Len(result_value) > Len(GetPathRoot(result_value)) And Right$(result_value, 1) = G_FS_PATH_SEP
        result_value = Left$(result_value, Len(result_value) - 1)
    Loop

    GetTemporaryDirectoryPath = result_value
End Function
Private Function IFileSystemService_GetTemporaryDirectoryPath() As String
    IFileSystemService_GetTemporaryDirectoryPath = GetTemporaryDirectoryPath()
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
    Const C_MAX_ATTEMPT As Long = 100

    Dim parent_path As String
    parent_path = GetTemporaryDirectoryPath()

    Dim attempt_idx As Long
    For attempt_idx = 1 To C_MAX_ATTEMPT
        Dim dir_path As String
        dir_path = JoinPath(parent_path, NamePrefix & pFSO.GetTempName())

        If Not PathExists(dir_path) Then
            On Error Resume Next
            Call pFSO.CreateFolder(dir_path)

            Dim err_num As Long
            Dim err_source As String
            Dim err_desc As String
            Dim err_help_file As String
            Dim err_help_context As Long
            err_num = Err.Number
            err_source = Err.Source
            err_desc = Err.Description
            err_help_file = Err.HelpFile
            err_help_context = Err.HelpContext
            Err.Clear
            On Error GoTo 0

            If err_num = 0 Then
                CreateTemporaryDirectory = dir_path
                Exit Function
            End If

            If Not PathExists(dir_path) Then
                Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
            End If
        End If
    Next attempt_idx

    Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService.CreateTemporaryDirectory", Description:="Could not create the temporary directory."
End Function
Private Function IFileSystemService_CreateTemporaryDirectory(Optional ByVal NamePrefix As String = "") As String
    IFileSystemService_CreateTemporaryDirectory = CreateTemporaryDirectory(NamePrefix:=NamePrefix)
End Function

' Gets a list of absolute file paths in the specified directory.
'
' @param DirectoryPath Directory path.
' @param CatchRegExp [Optional] Regular expression for file names to include in the list.
' @param IgnoreRegExp [Optional] Regular expression for file names to exclude from the list.
' @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
' @return Array of absolute paths.
'
' @details
' Gets a list of absolute file paths in the specified directory as an array. Returns a zero-based empty array when there are no matches.
' The return order follows the FileSystemObject / OS enumeration order and is not sorted.
' If an individual file is deleted or renamed during enumeration and getting its name or path raises a file or path not found error, that item is excluded and enumeration continues.
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    Dim file_list As Object 'Scripting.Files
    Set file_list = pFSO.GetFolder(dir_path).Files


    GetFileList = pFilterFiles(file_list, CatchRegExp, IgnoreRegExp, IgnoreCase)
End Function
Private Function IFileSystemService_GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetFileList = GetFileList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

' Gets a list of absolute subdirectory paths in the specified directory.
'
' @param DirectoryPath Directory path.
' @param CatchRegExp [Optional] Regular expression for subdirectory names to include in the list.
' @param IgnoreRegExp [Optional] Regular expression for subdirectory names to exclude from the list.
' @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
' @return Array of absolute paths.
'
' @details
' Gets a list of absolute subdirectory paths in the specified directory as an array. Returns a zero-based empty array when there are no matches.
' The return order follows the FileSystemObject / OS enumeration order and is not sorted.
' If an individual subdirectory is deleted or renamed during enumeration and getting its name or path raises a file or path not found error, that item is excluded and enumeration continues.
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    Dim folder_list As Object 'Scripting.Folders
    Set folder_list = pFSO.GetFolder(dir_path).SubFolders

    'GetDirectoryList = pFilterFolders(folder_list, CatchRegExp, IgnoreRegExp)
    GetDirectoryList = pFilterFiles(folder_list, CatchRegExp, IgnoreRegExp, IgnoreCase)
End Function
Private Function IFileSystemService_GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String()
    IFileSystemService_GetDirectoryList = GetDirectoryList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

Private Sub pCreateRegExpObject(ByRef CatchRegExpObject As Object, ByRef IgnoreRegExpObject As Object, ByVal CatchRegExp As String, ByVal IgnoreRegExp As String, ByVal IgnoreCase As Boolean)
    If CatchRegExp <> "" Then
        Set CatchRegExpObject = CreateObject("VBScript.RegExp")
        CatchRegExpObject.Pattern = CatchRegExp
        CatchRegExpObject.IgnoreCase = IgnoreCase
    End If

    If IgnoreRegExp <> "" Then
        Set IgnoreRegExpObject = CreateObject("VBScript.RegExp")
        IgnoreRegExpObject.Pattern = IgnoreRegExp
        IgnoreRegExpObject.IgnoreCase = IgnoreCase
    End If
End Sub

Private Function pIsMatchRegExp(ByVal Expression As String, ByVal CatchRegExpObject As Object, ByVal IgnoreRegExpObject As Object) As Boolean
    If pIsMatchRegExpCore(Expression, CatchRegExpObject, True) Then
        If pIsMatchRegExpCore(Expression, IgnoreRegExpObject, False) Then
            pIsMatchRegExp = False
        Else
            pIsMatchRegExp = True
        End If
    Else
        pIsMatchRegExp = False
    End If
End Function

Private Function pIsMatchRegExpCore(ByVal Expression As String, ByVal RegExpObject As Object, ByVal ResultWhenNothing As Boolean) As Boolean
    If RegExpObject Is Nothing Then
        pIsMatchRegExpCore = ResultWhenNothing
    Else
        If RegExpObject.Test(Expression) Then
            pIsMatchRegExpCore = True
        Else
            pIsMatchRegExpCore = False
        End If
    End If
End Function

Private Function pFilterFiles(ByVal FileList As Object, ByVal CatchRegExp As String, ByVal IgnoreRegExp As String, ByVal IgnoreCase As Boolean) As String()
    Dim catch_re As Object, ignore_re As Object 'RegExp
    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim result_arr() As String
    Dim arr_item As Object 'Scripting.File / Scripting.Folder
    For Each arr_item In FileList
        Dim item_name As String
        Dim item_path As String
        If pTryReadFileListItem(item_name, item_path, arr_item) Then
            If pIsMatchRegExp(item_name, catch_re, ignore_re) Then
                Call pAddItemToStringArray(result_arr, item_path)
            End If
        End If
    Next arr_item

    If IsEmptyArray(result_arr) Then
        pFilterFiles = EmptyStringArray()
    Else
        pFilterFiles = result_arr
    End If
End Function

Private Function pTryReadFileListItem(ByRef ItemName As String, ByRef ItemPath As String, ByVal FileListItem As Object) As Boolean
    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    Dim err_help_file As String
    Dim err_help_context As Long

    Err.Clear
    On Error Resume Next
    ItemName = FileListItem.Name
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num <> 0 Then
        If pIsMissingEnumeratedItemError(err_num) Then
            pTryReadFileListItem = False
            Exit Function
        End If
        Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
    End If

    Err.Clear
    On Error Resume Next
    ItemPath = FileListItem.Path
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    err_help_file = Err.HelpFile
    err_help_context = Err.HelpContext
    Err.Clear
    On Error GoTo 0

    If err_num <> 0 Then
        If pIsMissingEnumeratedItemError(err_num) Then
            pTryReadFileListItem = False
            Exit Function
        End If
        Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc, HelpFile:=err_help_file, HelpContext:=err_help_context
    End If

    pTryReadFileListItem = True
End Function

Private Function pIsMissingEnumeratedItemError(ByVal ErrorNumber As Long) As Boolean
    Select Case ErrorNumber
        Case C_ERR_FILE_NOT_FOUND, C_ERR_PATH_NOT_FOUND
            pIsMissingEnumeratedItemError = True
        Case Else
            pIsMissingEnumeratedItemError = False
    End Select
End Function

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

'Private Function pFilterFolders(FolderList As Object, CatchRegExp As String, IgnoreRegExp As String) As String()
'    Dim catch_re As Object, ignore_re As Object 'RegExp
'    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)
'
'    Dim result_arr() As String
'
'    Dim arr_item As Object 'Scripting.Folder
'    For Each arr_item In FolderList
'        If pIsMatchRegExp(arr_item.Name, catch_re, ignore_re) Then
'
'            Call pAddItemToStringArray(result_arr, arr_item.Path)
'        End If
'    Next arr_item
'
'    pFilterFolders = result_arr
'End Function

' Checks whether an entity exists at the specified path.
'
' @param PathLikeString Path string.
' @return Whether an entity exists.
'
' @details
' Checks whether an entity exists at the specified path. Returns True when a file or directory exists.
Public Function PathExists(ByVal PathLikeString As String) As Boolean
    Dim fs_path As String
    fs_path = GetAbsolutePath(PathLikeString)

    If pFSO.FileExists(fs_path) Or pFSO.FolderExists(fs_path) Then
        PathExists = True
    Else
        PathExists = False
    End If
End Function
Private Function IFileSystemService_PathExists(ByVal PathLikeString As String) As Boolean
    IFileSystemService_PathExists = PathExists(PathLikeString)
End Function

' Checks whether a file exists at the specified path.
'
' @param FilePath File path.
' @return Whether an entity exists.
'
' @details
' Checks whether a file exists at the specified path.
Public Function IsFile(ByVal FilePath As String) As Boolean
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If pFSO.FileExists(file_path) Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
Private Function IFileSystemService_IsFile(ByVal FilePath As String) As Boolean
    IFileSystemService_IsFile = IsFile(FilePath)
End Function

' Checks whether a directory exists at the specified path.
'
' @param DirectoryPath Directory path.
' @return Whether an entity exists.
'
' @details
' Checks whether a directory exists at the specified path.
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If pFSO.FolderExists(dir_path) Then
        IsDirectory = True
    Else
        IsDirectory = False
    End If
End Function
Private Function IFileSystemService_IsDirectory(ByVal DirectoryPath As String) As Boolean
    IFileSystemService_IsDirectory = IsDirectory(DirectoryPath)
End Function

' Gets the last modified date and time.
'
' @param PathLikeString Path string.
' @return Last modified date and time.
'
' @details
' Gets the last modified date and time.
Public Function GetLastModified(ByVal PathLikeString As String) As Date
    Dim path_str As String
    path_str = GetAbsolutePath(PathLikeString)
    Dim path_obj As Object
    If IsFile(path_str) Then
        Set path_obj = pFSO.GetFile(path_str)
    ElseIf IsDirectory(path_str) Then
        Set path_obj = pFSO.GetFolder(path_str)
    Else
        Err.Raise vbObjectError + 1, "Class FileSystemService.GetLastModified", "The file or directory to get the modified date from does not exist. (" & path_str & ")"
    End If
    GetLastModified = path_obj.DateLastModified
End Function
Private Function IFileSystemService_GetLastModified(ByVal PathLikeString As String) As Date
    IFileSystemService_GetLastModified = GetLastModified(PathLikeString)
End Function

' Gets the file size in bytes.
'
' @param FilePath File path.
' @return File size, in bytes.
'
' @details
' Gets the size of the specified file in bytes.
Public Function GetFileSize(ByVal FilePath As String) As Double
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If pFSO.FileExists(file_path) Then
        GetFileSize = pFSO.GetFile(file_path).Size
    ElseIf pFSO.FolderExists(file_path) Then
        Err.Raise vbObjectError + 1, "Class FileSystemService.GetFileSize", "The target for getting the file size is not a file. (" & file_path & ")"
    Else
        Err.Raise vbObjectError + 1, "Class FileSystemService.GetFileSize", "The file to get the file size from does not exist. (" & file_path & ")"
    End If
End Function
Private Function IFileSystemService_GetFileSize(ByVal FilePath As String) As Double
    IFileSystemService_GetFileSize = GetFileSize(FilePath)
End Function

' Creates a directory.
'
' @param DirectoryPath Directory path.
' @param Force [Optional] Whether to overwrite (delete) anything that already exists at the destination path.
' @param Recursive [Optional] Whether to create parent directories recursively when they do not exist.
' @return Whether a directory was actually created.
'
' @details
' Creates a directory at the specified path. If the directory already exists, does nothing and returns False.
' Returns True when a directory was actually created.
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If IsDirectory(dir_path) Then
        Debug.Print dir_path & " was requested for creation, but it already exists."
        CreateDirectory = False
        Exit Function
    End If

    Dim parent_dir As String
    parent_dir = GetParentPath(dir_path, IgnoreEndSep:=True)

    If IsDirectory(parent_dir) Then
        ' If Force is True, delete the existing file.
        If Force And IsFile(dir_path) Then Call RemoveFile(dir_path, Force:=True)
        ' The parent directory exists, so create the directory normally.
        Call pFSO.CreateFolder(dir_path)
        CreateDirectory = True
    Else
        ' When the parent directory does not exist.
        If Not Recursive Then Err.Raise vbObjectError + 1, "Class FileSystemService.CreateDirectory", "The parent directory does not exist. (" & parent_dir & ")"

        Dim dir_arr() As String

        ' Walk up the parent directories while storing directory names in an array.
        Do
            Call pAddItemToStringArray(dir_arr, GetLeafFromPath(dir_path, IgnoreEndSep:=True))
            dir_path = parent_dir
            parent_dir = GetParentPath(dir_path, IgnoreEndSep:=True)
        Loop While dir_path <> "" And Not IsDirectory(parent_dir)

        If dir_path = "" Then
            ' What if no root is found?
            Err.Raise vbObjectError + 1, "Class FileSystemService.CreateDirectory", "Could not create the directory. (" & dir_path & ")"
        Else
            ' If Force is True, delete the existing file.
            If Force And IsFile(dir_path) Then Call RemoveFile(dir_path, Force:=True)
            ' Create the directories in order.
            Dim item_idx As Long
            For item_idx = UBound(dir_arr) To LBound(dir_arr) Step -1
                Call pFSO.CreateFolder(dir_path)
                dir_path = JoinPath(dir_path, dir_arr(item_idx))
            Next item_idx
            Call pFSO.CreateFolder(dir_path)
            CreateDirectory = True
        End If
    End If
End Function
Private Function IFileSystemService_CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    IFileSystemService_CreateDirectory = CreateDirectory(DirectoryPath, Force:=Force, Recursive:=Recursive)
End Function

' Moves a directory.
'
' @param SourcePath Source directory path.
' @param DestinationPath Destination directory path.
' @param Force [Optional] Whether to overwrite anything that already exists at the destination path.
'
' @details
' Moves the directory at the specified path. If DestinationPath ends with a path separator, moves it under that directory with the current directory name.
' When Force:=True, copies to the destination using the same overwrite rules as CopyDirectory, then deletes the source. Extra files under the existing destination directory are not deleted.
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsDirectory(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService.MoveDirectory", "The directory to move does not exist. (" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=True)

    If Force Then
        Call pFSO.CopyFolder(src_path, real_dst_path, OverWriteFiles:=True)
        Call RemoveDirectory(src_path, Force:=True)
    Else
        Call pFSO.MoveFolder(src_path, real_dst_path)
    End If
End Sub
Private Sub IFileSystemService_MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

' Copies a directory.
'
' @param SourcePath Source directory path.
' @param DestinationPath Destination directory path.
' @param Force [Optional] Whether to overwrite anything that already exists at the destination path.
'
' @details
' Copies the directory at the specified path. If DestinationPath ends with a path separator, copies it under that directory with the current directory name.
' When Force:=True, overwrites existing files. Extra files under the existing destination directory are not deleted.
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsDirectory(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService.CopyDirectory", "The directory to copy does not exist. (" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=True)

    Call pFSO.CopyFolder(src_path, real_dst_path, OverWriteFiles:=Force)
End Sub
Private Sub IFileSystemService_CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

' Deletes a directory.
'
' @param DirectoryPath Directory path.
' @param Force [Optional] Whether to delete even when the read-only attribute is set.
' @return Whether a directory was actually deleted.
'
' @details
' Deletes the directory at the specified path. If the directory does not already exist, does nothing and returns False.
' Returns True when a directory was actually deleted.
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)

    If Not IsDirectory(dir_path) Then
        Debug.Print dir_path & " was requested for deletion, but it did not exist."
        RemoveDirectory = False
        Exit Function
    End If

    Call pFSO.DeleteFolder(dir_path, Force:=Force) ' Deletes non-empty directories even when Force is False; this controls whether read-only attributes are ignored.
    RemoveDirectory = True
End Function
Private Function IFileSystemService_RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveDirectory = RemoveDirectory(DirectoryPath, Force:=Force)
End Function

' Moves a file.
'
' @param SourcePath Source file path.
' @param DestinationPath Destination file path.
' @param Force [Optional] Whether to overwrite anything that already exists at the destination path.
'
' @details
' Moves the file at the specified path. If DestinationPath ends with a path separator, moves it under that directory with the current file name.
' When Force:=True, copies to the destination using the same overwrite rules as CopyFile, then deletes the source.
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsFile(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService.MoveFile", "The file to move does not exist. (" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=False)

    If Force Then
        Call pFSO.CopyFile(src_path, real_dst_path, OverWriteFiles:=True)
        Call RemoveFile(src_path, Force:=True)
    Else
        Call pFSO.MoveFile(src_path, real_dst_path)
    End If
End Sub
Private Sub IFileSystemService_MoveFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveFile(SourcePath, DestinationPath, Force:=Force)
End Sub

' Copies a file.
'
' @param SourcePath Source file path.
' @param DestinationPath Destination file path.
' @param Force [Optional] Whether to overwrite anything that already exists at the destination path.
'
' @details
' Copies the file at the specified path. If DestinationPath ends with a path separator, copies it under that directory with the current file name.
' When Force:=True, overwrites existing files.
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Dim src_path As String, dst_path As String
    src_path = GetAbsolutePath(SourcePath)
    dst_path = GetAbsolutePath(DestinationPath)

    If Not IsFile(src_path) Then Err.Raise vbObjectError + 1, "Class FileSystemService.CopyFile", "The file to copy does not exist. (" & src_path & ")"

    Dim ends_with_sep As Boolean
    ends_with_sep = pEndsWithPathSeparator(DestinationPath)

    Dim real_dst_path As String
    real_dst_path = pResolveOperationDestinationPath(src_path, dst_path, ends_with_sep)
    Call pValidateOperationDestination(src_path, real_dst_path, SourceIsDirectory:=False)

    Call pFSO.CopyFile(src_path, real_dst_path, OverWriteFiles:=Force)
End Sub
Private Sub IFileSystemService_CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call CopyFile(SourcePath, DestinationPath, Force:=Force)
End Sub

' Deletes a file.
'
' @param FilePath File path.
' @param Force [Optional] Whether to delete even when the read-only attribute is set.
' @return Whether a file was actually deleted.
'
' @details
' Deletes the file at the specified path. If the file does not already exist, does nothing and returns False.
' Returns True when a file was actually deleted.
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    Dim file_path As String
    file_path = GetAbsolutePath(FilePath)

    If Not IsFile(file_path) Then
        Debug.Print file_path & " was requested for deletion, but it did not exist."
        RemoveFile = False
        Exit Function
    End If

    Call pFSO.DeleteFile(file_path, Force:=Force) ' Locks are irrelevant; this controls whether read-only attributes are ignored.
    RemoveFile = True
End Function
Private Function IFileSystemService_RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveFile = RemoveFile(FilePath, Force:=Force)
End Function

' Gets the newest file among files with the specified name directly under the specified folder.
'
' @param DirectoryPath Path of the folder to search.
' @param CatchRegExp [Optional] Regular expression for file names to check.
' @param IgnoreRegExp [Optional] Regular expression for file names not to check.
' @param IgnoreCase [Optional] Whether to ignore case in the regular expressions.
' @return Full path of the newest file. Returns an empty string when no matching file exists.
'
' @details
' Searches files in the specified folder and returns the full path of the newest file among those matching the specified pattern.
' The traversal order follows the FileSystemObject / OS enumeration order, and ordering for identical timestamps is not guaranteed.
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    Dim catch_re As Object, ignore_re As Object 'RegExp
    Call pCreateRegExpObject(catch_re, ignore_re, CatchRegExp, IgnoreRegExp, IgnoreCase)

    Dim dir_obj As Object 'Folder
    Dim dir_path As String
    dir_path = GetAbsolutePath(DirectoryPath)
    Set dir_obj = pFSO.GetFolder(dir_path)

    Dim newest_file As Object 'File
    Dim file_item As Object 'File
    For Each file_item In dir_obj.Files
        If pIsMatchRegExp(file_item.Name, catch_re, ignore_re) Then
            If newest_file Is Nothing Then
                Set newest_file = file_item
            Else
                If newest_file.DateLastModified < file_item.DateLastModified Then
                    Set newest_file = file_item
                End If
            End If
        End If
    Next file_item

    If newest_file Is Nothing Then
        GetNewestFile = ""
    Else
        GetNewestFile = newest_file.Path
    End If
End Function
Private Function IFileSystemService_GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "", Optional ByVal IgnoreCase As Boolean = True) As String
    IFileSystemService_GetNewestFile = GetNewestFile(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp, IgnoreCase:=IgnoreCase)
End Function

' Creates a backup file.
'
' @param SourceFilePath [Optional] File path to back up. The default is the current workbook path.
' @param DestinationDirectoryPath [Optional] Destination directory path for the backup. The default is the target file's directory.
' @param BackupSuffix [Optional] Suffix to add to the backup file name. The default is a timestamp.
' @param Force [Optional] Whether to force overwriting an existing backup file.
' @param Move [Optional] Whether to move the original file during backup. True means move.
' @param FailIfSourceNotFound Whether to raise an error when the target file does not exist. True means raise an error.
' @return Full path of the created backup file.
'
' @details
' Creates a backup of the specified file. The backup destination file name is the original file name with
' the suffix appended. If an existing file is present, it is handled based on the Force setting.
' Force:=True overwrites the existing backup file.
' When Force:=False and the default suffix from BackupSuffix:="" collides, returns the existing backup path and does not compare contents or move with Move:=True.
' A collision when Force:=False and BackupSuffix is explicitly specified is an error.
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String

    Dim src_file_path As String
    Dim src_file_name As String
    Dim src_file_obj As File
    Dim file_base As String
    Dim file_ext As String
    Dim dst_file_path As String
    Dim dst_dir_path As String
    Dim suffix_str As String

    ' If no source file is specified for backup, use this workbook.
    If SourceFilePath = "" Then
        src_file_path = ThisWorkbook.FullName
    Else
        src_file_path = GetAbsolutePath(SourceFilePath)
    End If

    ' Raise an error when the source file for backup does not exist.
    If Not IsFile(src_file_path) Then
        If FailIfSourceNotFound Then
            Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService.CreateBackupFile", Description:="The file to back up does not exist. (" & SourceFilePath & ")"
        Else
            Debug.Print "The file to back up did not exist. (" & src_file_path & ")"
            Exit Function
        End If
    End If

    ' If no backup destination directory is specified, use the same directory as the target file.
    If DestinationDirectoryPath = "" Then
        dst_dir_path = GetParentPath(src_file_path, IgnoreEndSep:=True)
    Else
        If IsAbsolutePath(DestinationDirectoryPath) Then
            dst_dir_path = DestinationDirectoryPath
        Else
            dst_dir_path = JoinPath(GetParentPath(src_file_path), DestinationDirectoryPath)
        End If
    End If

    ' Create the backup destination directory when it does not exist.
    If IsFile(dst_dir_path) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService.CreateBackupFile", Description:="The backup destination is not a directory. (" & DestinationDirectoryPath & ")"
    ElseIf Not IsDirectory(dst_dir_path) Then
        Call CreateDirectory(dst_dir_path, Recursive:=True)
    End If

    ' If no suffix is specified, use the last modified date and time.
    If BackupSuffix = "" Then
        suffix_str = Format(GetLastModified(src_file_path), "_yyyy-MM-dd_HH-mm-ss")
    Else
        suffix_str = BackupSuffix
    End If

    ' Get the file name from the original path.
    file_base = GetLeafFromPath(src_file_path, BaseName:=True, Extension:=False)
    file_ext = GetLeafFromPath(src_file_path, BaseName:=False, Extension:=True)

    ' Get the backup destination file path.
    dst_file_path = JoinPath(dst_dir_path, file_base & suffix_str & file_ext)

    If IsFile(dst_file_path) Then
        If Not Force Then
            ' If the force flag is not set and the backup target file exists, raise an error.
            If BackupSuffix = "" Then
                ' If a time-based suffix duplicates an existing backup, backup is already complete, so finish processing.
                Debug.Print "The file was already backed up. (" & src_file_path & ")"
                CreateBackupFile = dst_file_path
                Exit Function
            Else
                Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService.CreateBackupFile", Description:="The backup destination file already exists. (" & dst_file_path & ")"
            End If
        End If
    ElseIf IsDirectory(dst_file_path) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class FileSystemService.CreateBackupFile", Description:="A directory with the same name as the backup destination file exists. (" & dst_file_path & ")"
    End If

    ' Back up the file.
    If Move Then
        Call MoveFile(src_file_path, dst_file_path, Force:=Force)
    Else
        Call CopyFile(src_file_path, dst_file_path, Force:=Force)
    End If

    CreateBackupFile = dst_file_path
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

Private Function pEndsWithPathSeparator(ByVal PathLikeString As String) As Boolean
    Dim last_char As String
    last_char = Right$(PathLikeString, 1)
    pEndsWithPathSeparator = (last_char = G_FS_PATH_SEP Or last_char = "/")
End Function

Private Function pResolveOperationDestinationPath(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal EndsWithSep As Boolean) As String
    If EndsWithSep Then
        pResolveOperationDestinationPath = JoinPath(DestinationPath, GetLeafFromPath(SourcePath, IgnoreEndSep:=True))
    Else
        pResolveOperationDestinationPath = DestinationPath
    End If
End Function

Private Sub pValidateOperationDestination(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal SourceIsDirectory As Boolean)
    If pPathsEqual(SourcePath, DestinationPath) Then
        Err.Raise vbObjectError + 1, "Class FileSystemService.pValidateOperationDestination", "The copy/move source and copy/move destination are the same. (" & SourcePath & ")"
    End If

    If SourceIsDirectory Then
        If pIsPathInside(DestinationPath, SourcePath) Then
            Err.Raise vbObjectError + 1, "Class FileSystemService.pValidateOperationDestination", "The copy/move destination is under the copy/move source. (" & DestinationPath & ")"
        End If
        If pIsPathInside(SourcePath, DestinationPath) Then
            Err.Raise vbObjectError + 1, "Class FileSystemService.pValidateOperationDestination", "The copy/move destination is the parent directory of the copy/move source. (" & DestinationPath & ")"
        End If
    End If
End Sub

Private Function pPathsEqual(ByVal Path1 As String, ByVal Path2 As String) As Boolean
    pPathsEqual = (StrComp(pNormalizeComparePath(Path1), pNormalizeComparePath(Path2), vbTextCompare) = 0)
End Function

Private Function pIsPathInside(ByVal ChildPath As String, ByVal ParentPath As String) As Boolean
    Dim child_path_norm As String
    child_path_norm = pNormalizeComparePath(ChildPath)
    Dim parent_path_norm As String
    parent_path_norm = pNormalizeComparePath(ParentPath)

    If Right$(parent_path_norm, 1) <> G_FS_PATH_SEP Then parent_path_norm = parent_path_norm & G_FS_PATH_SEP
    pIsPathInside = (Len(parent_path_norm) < Len(child_path_norm) And StrComp(Left$(child_path_norm, Len(parent_path_norm)), parent_path_norm, vbTextCompare) = 0)
End Function

Private Function pNormalizeComparePath(ByVal PathLikeString As String) As String
    Dim path_value As String
    path_value = GetAbsolutePath(PathLikeString)
    Dim root_value As String
    root_value = GetPathRoot(path_value)

    Do While Len(path_value) > Len(root_value) And Right$(path_value, 1) = G_FS_PATH_SEP
        path_value = Left$(path_value, Len(path_value) - 1)
    Loop

    pNormalizeComparePath = path_value
End Function
