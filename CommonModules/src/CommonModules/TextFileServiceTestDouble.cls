VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileService

' #############################################################################
'!
'! @brief
'! Test double for the service that creates TextFileEntity objects.
'!
'! @details
'! GetTextFileEntity is the entry point for lightweight fakes. When unregistered, returns a new initialized TextFileEntityTestDouble.
'! File contents and open failures are not reproduced automatically. Configure arbitrary failures on GetTextFileEntity with Store.SetError.
'!
' #############################################################################

'* Stub settings and history of successfully completed calls.
'* Unregistered strict stub errors and explicit errors from SetError are not recorded in call history.
Public Store As TestDoubleBehaviorStore

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* Lightweight fake for GetTextFileEntity. When unregistered, creates a new initialized TextFileEntityTestDouble.
Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Call Store.RaiseIfError("GetTextFileEntity", FilePath)

    Dim result_value As ITextFileEntity
    If Store.HasReturn("GetTextFileEntity", FilePath) Then
        Set result_value = Store.GetReturn("GetTextFileEntity", FilePath)
        Call pValidateEntityFilePath(result_value, FilePath)
    Else
        Dim file_stub As TextFileEntityTestDouble
        Set file_stub = New TextFileEntityTestDouble
        Call file_stub.Initialize(FilePath)
        Set result_value = file_stub
    End If

    Call Store.RecordCall("GetTextFileEntity", FilePath)
    Set GetTextFileEntity = result_value
End Function
Private Function ITextFileService_GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Set ITextFileService_GetTextFileEntity = GetTextFileEntity(FilePath)
End Function

Private Sub pValidateEntityFilePath(ByVal TextFileEntity As ITextFileEntity, ByVal ExpectedFilePath As String)
    On Error GoTo INVALID_ENTITY

    Dim actual_path As String
    actual_path = TextFileEntity.FilePath

    On Error GoTo 0
    If actual_path <> ExpectedFilePath Then
        Err.Raise vbObjectError + 1, "Class TextFileServiceTestDouble.pValidateEntityFilePath", "The registered TextFileEntity FilePath does not match. (expected: " & ExpectedFilePath & ", actual: " & actual_path & ")"
    End If
    Exit Sub

INVALID_ENTITY:
    Err.Raise vbObjectError + 1, "Class TextFileServiceTestDouble.pValidateEntityFilePath", "The registered TextFileEntity has not been initialized with the specified path. (" & ExpectedFilePath & ")"
End Sub
