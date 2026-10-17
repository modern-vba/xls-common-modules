VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUserInputSheet

' #############################################################################
'!
'! @brief
'! Class for accessing a user interface sheet.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pRangeBounds As WorksheetRangeBounds

'* Target range from which item names and values are read in the input sheet.
Public Property Get InputArea() As WorksheetRangeBounds
    Call pCheckInit
    Set InputArea = pRangeBounds.CopyObject()
End Property
Private Property Get IUserInputSheet_InputArea() As WorksheetRangeBounds
    Set IUserInputSheet_InputArea = InputArea
End Property

'* Sets each property.
'*
'* @param InputArea Target range from which item names and values are read in the input sheet.
'*
'* @details
'* Sets each property.
Public Sub Initialize(ByVal InputArea As WorksheetRangeBounds)
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "The object is already initialized."
    If InputArea Is Nothing Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "InputArea is not specified."
    If InputArea.IsEmpty Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "InputArea is an empty range."
    If Not WbSrv.WorksheetExists(InputArea.WorksheetName, Book:=InputArea.WorkbookName) Then Err.Raise vbObjectError + 1, "Class UserInputSheet.Initialize", "The target sheet was not found. (" & InputArea.WorksheetName & "@" & InputArea.WorkbookName & ")"

    Set pRangeBounds = InputArea.CopyObject()
    pIsInitialized = True
End Sub

'* Gets the range of the item value that corresponds to the specified item name.
'*
'* @param FirstItemName First item name.
'* @param SecondItemName Second item name (optional).
'* @return Corresponding item value range (Range object).
'*
'* @details
'* Gets the row or column range based on the position of the first item name.
'* If the second item name is omitted, returns the entire row range for the first item name.
'* If the second item name is specified, returns the range corresponding to that column.
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Call pCheckInit

    Dim used_range As WorksheetRangeBounds
    Set used_range = WsSrv.GetUsedRangeBounds(pRangeBounds)

    ' Treat the first column as the first item name column and split off the rest.
    Dim first_col As WorksheetRangeBounds
    Set first_col = used_range.TransformAbsolute(FinishColumn:=used_range.Column)

    Dim remain_range As WorksheetRangeBounds
    Set remain_range = used_range.TransformAbsolute(Column:=used_range.Column + 1)

    ' Search the first item name column by name.
    Dim found_cells As ObjectList
    Set found_cells = WsSrv.Find(FirstItemName, first_col)

    ' If an empty list is returned, the item was not found, so end with Nothing.
    If found_cells.Count = 0 Then
        Set GetItemRange = Nothing
        Exit Function
    End If

    ' Row immediately to the right of the found first item.
    Dim second_row As WorksheetRangeBounds
    Dim first_found_cell As WorksheetRangeBounds
    Set first_found_cell = found_cells.Item(0)

    Set second_row = remain_range.TransformAbsolute(Row:=first_found_cell.Row, FinishRow:=first_found_cell.Row)

    ' Trim the range.
    Set second_row = WsSrv.GetUsedRangeBounds(second_row)

    ' When no second item name is specified, second_row is the value range.
    If SecondItemName = "" Then
        If second_row.IsEmpty Then
            ' For an empty range, return it as one cell.
            Set GetItemRange = second_row.TransformAbsolute(FinishRow:=second_row.Row, FinishColumn:=second_row.Column)
        Else
            Set GetItemRange = second_row
        End If
        Exit Function
    End If

    ' Search for the second item name.
    Set found_cells = WsSrv.Find(SecondItemName, second_row)

    ' If an empty list is returned, the item was not found, so end with Nothing.
    If found_cells.Count = 0 Then
        Set GetItemRange = Nothing
        Exit Function
    End If

    ' Find the next first item and calculate the row count for the second item's values.
    Dim first_remain_col As WorksheetRangeBounds
    Dim second_found_cell As WorksheetRangeBounds
    Set second_found_cell = found_cells.Item(0)

    Set first_remain_col = first_col.TransformAbsolute(Row:=second_found_cell.Row + 1)
    Dim finish_row_idx As Long
    finish_row_idx = second_found_cell.Row
    Dim enum_obj As IEnumerator
    Set enum_obj = first_remain_col.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current
        If WsSrv.IsEmptyCell(cell_item, IgnoreEmptyString:=True) Then
            finish_row_idx = finish_row_idx + 1
        Else
            Exit Do
        End If
    Loop
    finish_row_idx = MinLng(finish_row_idx, used_range.FinishRow)

    ' Build the result range.
    Dim result_value As WorksheetRangeBounds
    Set result_value = remain_range.TransformAbsolute(Row:=second_found_cell.Row + 1, Column:=second_found_cell.Column, FinishRow:=finish_row_idx, FinishColumn:=second_found_cell.Column)
    Set result_value = WsSrv.GetUsedRangeBounds(result_value, IgnoreEmpty:=True)
    Set GetItemRange = result_value
End Function
Private Function IUserInputSheet_GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds

    Set IUserInputSheet_GetItemRange = GetItemRange(FirstItemName, SecondItemName:=SecondItemName)
End Function

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet.pCheckInit", "The object is not initialized."
End Sub
