VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ObjectListComparableStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IComparable

' #############################################################################
'!
'! @brief
'! Stub for verifying ObjectList IComparable sorting.
'!
' #############################################################################

Private pSortKey As Long
Private pItemName As String

'* Comparison key for sorting.
Public Property Get SortKey() As Long
    SortKey = pSortKey
End Property
'* Comparison key for sorting.
Public Property Let SortKey(ByVal NewValue As Long)
    pSortKey = NewValue
End Property

'* Name for identification.
Public Property Get ItemName() As String
    ItemName = pItemName
End Property
'* Name for identification.
Public Property Let ItemName(ByVal NewValue As String)
    pItemName = NewValue
End Property

Private Function IComparable_CompareTo(ByVal OtherObject As Object) As Integer
    If TypeName(OtherObject) <> TypeName(Me) Then
        Err.Raise vbObjectError + 1, "Class Test_ObjectListComparableStub.IComparable_CompareTo", "The comparison target type differs."
    End If

    Dim other_item As Test_ObjectListComparableStub
    Set other_item = OtherObject

    If pSortKey < other_item.SortKey Then
        IComparable_CompareTo = -1
    ElseIf pSortKey > other_item.SortKey Then
        IComparable_CompareTo = 1
    Else
        IComparable_CompareTo = 0
    End If
End Function

Private Function IComparable_IsGreaterThan(ByVal OtherObject As Object) As Boolean
    IComparable_IsGreaterThan = False
End Function

Private Function IComparable_IsLessThan(ByVal OtherObject As Object) As Boolean
    IComparable_IsLessThan = False
End Function

Private Function IComparable_IsEqualTo(ByVal OtherObject As Object) As Boolean
    IComparable_IsEqualTo = False
End Function
