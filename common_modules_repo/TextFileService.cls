VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileService

' #############################################################################
'!
'! @brief
'! Implementation of the service that creates TextFileEntity objects.
'!
' #############################################################################

Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Dim result_value As TextFileEntity
    Set result_value = New TextFileEntity
    Call result_value.Initialize(FilePath)
    Set GetTextFileEntity = result_value
End Function
Private Function ITextFileService_GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Set ITextFileService_GetTextFileEntity = GetTextFileEntity(FilePath)
End Function
