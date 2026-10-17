VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_WorksheetChangeTestProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Test probe class that fires events for tests of Worksheet_Change protection behavior.
'!
' #############################################################################

Private WithEvents pTargetSheet As Worksheet
Attribute pTargetSheet.VB_VarHelpID = -1
Private pTargetAddress As String
Private pMarkerAddress As String

Public Sub Initialize( _
        ByVal TargetSheet As Worksheet, _
        ByVal TargetAddress As String, _
        ByVal MarkerAddress As String)

    Set pTargetSheet = TargetSheet
    pTargetAddress = TargetAddress
    pMarkerAddress = MarkerAddress
End Sub

Private Sub pTargetSheet_Change(ByVal Target As Range)
    If Intersect(Target, pTargetSheet.Range(pTargetAddress)) Is Nothing Then Exit Sub

    If pTargetSheet.Range(pMarkerAddress).Value <> "PROTECTED" Then
        pTargetSheet.Range(pMarkerAddress).Value = "PROTECTED"
        Call pTargetSheet.Protect(Contents:=True, AllowFormattingCells:=False)
    End If
End Sub
