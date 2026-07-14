VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITextFileService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for a service that creates ITextFileEntity objects.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface ITextFileService", Description:="This class is intended for interfaces."
End Sub

'* Creates a TextFileEntity.
'*
'* @param FilePath File path.
'*
'* @details
'* Creates a TextFileEntity.
Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
End Function
