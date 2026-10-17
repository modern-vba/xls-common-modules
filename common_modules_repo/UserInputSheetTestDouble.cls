VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputSheetTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUserInputSheet

' #############################################################################
'!
'! @brief
'! Test double for accessing an input sheet.
'!
'! @details
'! GetItemRange is a strict stub, and unregistered return value errors and explicit errors from SetError are not recorded in call history.
'!
' #############################################################################

'* Stub settings and history of successfully completed calls.
Public Store As TestDoubleBehaviorStore

'* Safe default for the input area. Returns the entire InputSheet sheet when unregistered.
Public Property Get InputArea() As WorksheetRangeBounds
    Dim default_area As WorksheetRangeBounds
    Set default_area = New_RangeBounds(Sheet:="InputSheet", Book:="ThisWorkbook.xlsm")
    Set InputArea = Store.GetReturnOrDefault("InputArea", default_area)
End Property
Private Property Get IUserInputSheet_InputArea() As WorksheetRangeBounds
    Set IUserInputSheet_InputArea = InputArea
End Property

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* Strict stub for GetItemRange. Register the return value explicitly in Store.
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Call Store.RaiseIfError("GetItemRange", FirstItemName, SecondItemName)

    Dim result_value As WorksheetRangeBounds
    Set result_value = Store.GetReturn("GetItemRange", FirstItemName, SecondItemName)
    Call Store.RecordCall("GetItemRange", FirstItemName, SecondItemName)

    Set GetItemRange = result_value
End Function
Private Function IUserInputSheet_GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Set IUserInputSheet_GetItemRange = GetItemRange(FirstItemName, SecondItemName:=SecondItemName)
End Function
