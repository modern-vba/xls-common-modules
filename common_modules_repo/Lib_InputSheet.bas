Attribute VB_Name = "Lib_InputSheet"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Module that provides a factory for UserInputSheet.
'!
' #############################################################################

'* Creates a new UserInputSheet instance.
'*
'* @param InputArea Target range from which item names and values are read in the input sheet.
'* @return New UserInputSheet instance.
'*
'* @details
'* Creates a new UserInputSheet instance.
Public Function New_InputSheet(ByVal InputArea As WorksheetRangeBounds) As UserInputSheet
    Dim result_value As UserInputSheet
    Set result_value = New UserInputSheet
    Call result_value.Initialize(InputArea)
    Set New_InputSheet = result_value
End Function
