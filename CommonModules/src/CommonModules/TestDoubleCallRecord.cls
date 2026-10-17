VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDoubleCallRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Record of one test double call.
'!
' #############################################################################

Private pMethodName As String
Private pCallIndex As Long
Private pObjectKeyMode As Long
Private pArguments As Collection

'* Called method name.
Public Property Get MethodName() As String
    MethodName = pMethodName
End Property
Public Property Let MethodName(ByVal NewValue As String)
    pMethodName = NewValue
End Property

'* 0-based call number in call history.
Public Property Get CallIndex() As Long
    CallIndex = pCallIndex
End Property
Public Property Let CallIndex(ByVal NewValue As Long)
    pCallIndex = NewValue
End Property

'* Object keying mode used for call matching.
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property
Public Property Let ObjectKeyMode(ByVal NewValue As Long)
    pObjectKeyMode = NewValue
End Property

'* Returns the number of call arguments.
Public Property Get ArgumentCount() As Long
    ArgumentCount = pArguments.Count
End Property

Private Sub Class_Initialize()
    Set pArguments = New Collection
End Sub

'* Stores call arguments.
Public Sub StoreArguments(ByVal Args As Variant)
    If Not IsArray(Args) Then
        Err.Raise vbObjectError + 1, "Class TestDoubleCallRecord.StoreArguments", "Call arguments must be specified as an array."
    End If

    Set pArguments = New Collection
    If IsEmptyArray(Args) Then Exit Sub

    Dim arg_item As Variant
    For Each arg_item In Args
        Call pAddArgument(arg_item)
    Next arg_item
End Sub

Private Sub pAddArgument(ByVal NewValue As Variant)
    Dim value_holder As Dictionary
    Set value_holder = New Dictionary

    If IsObject(NewValue) Then
        Set value_holder.Item("Value") = NewValue
    Else
        value_holder.Item("Value") = NewValue
    End If

    Call pArguments.Add(value_holder)
End Sub

'* Returns a call argument by its 0-based argument number.
Public Function GetArgument(ByVal ArgumentIndex As Long) As Variant
    If ArgumentIndex < 0 Or pArguments.Count <= ArgumentIndex Then
        Err.Raise vbObjectError + 1, "Class TestDoubleCallRecord.GetArgument", "Call argument index is out of range. (" & CStr(ArgumentIndex) & ")"
    End If

    Dim value_holder As Dictionary
    Set value_holder = pArguments.Item(ArgumentIndex + 1)

    If IsObject(value_holder.Item("Value")) Then
        Set GetArgument = value_holder.Item("Value")
    Else
        GetArgument = value_holder.Item("Value")
    End If
End Function
