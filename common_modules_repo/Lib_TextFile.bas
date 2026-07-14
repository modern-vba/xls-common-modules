Attribute VB_Name = "Lib_TextFile"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Module that manages the shared TextFileService instance.
'!
' #############################################################################

'* TextFileService. Replace with a test double during unit tests.
Public TfSrv As ITextFileService

'* Initializes TextFileService.
'*
'* @param Force [Optional] When True, recreates a replaced service as the production service.
'*
'* @details
'* Creates TextFileService when TfSrv is unset or Force is True. If it has been replaced in tests, keeps it only when Force is False.
Public Sub InitializeTextFileService(Optional ByVal Force As Boolean = False)
    If Force Or (TfSrv Is Nothing) Then Set TfSrv = New TextFileService
End Sub
