VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUserInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for accessing an input sheet.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IUserInputSheet", Description:="This class is intended for interfaces."
End Sub

'* Target range from which item names and values are read in the input sheet.
Public Property Get InputArea() As WorksheetRangeBounds
End Property

'* Gets the range of the item value that corresponds to the specified item name.
'*
'* @param FirstItemName First item name.
'* @param SecondItemName Second item name. When omitted, gets the value range for the first item name.
'* @return Corresponding item value range. Nothing if not found.
'*
'* @details
'* Returns the value range that corresponds to the item header in the input area.
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds
End Function
