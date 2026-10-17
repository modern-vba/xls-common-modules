Attribute VB_Name = "Lib_FileSystem"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Module that manages the shared FileSystemService instance.
'!
' #############################################################################

'* FileSystemService. Replace with a test double during unit tests.
Public FsSrv As IFileSystemService

'* Initializes FileSystemService.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Creates FileSystemService when FsSrv is unset or Force is True. If it has been replaced in tests, keeps it only when Force is False.
Public Sub InitializeFileSystemService(Optional ByVal Force As Boolean = False)
    If Force Or (FsSrv Is Nothing) Then Set FsSrv = New FileSystemService
End Sub
