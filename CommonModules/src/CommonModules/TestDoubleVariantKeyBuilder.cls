VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDoubleVariantKeyBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class that generates keys for test doubles from Variant values.
'!
' #############################################################################

Private pObjectKeyMode As Long

'* Keying mode applied to top-level object arguments.
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property
Public Property Let ObjectKeyMode(ByVal NewValue As Long)
    Call pValidateObjectKeyMode(NewValue)
    pObjectKeyMode = NewValue
End Property

Private Sub Class_Initialize()
    pObjectKeyMode = G_OBJECT_KEY_MODE_REFERENCE
End Sub

'* Generates a key from the specified arguments.
'*
'* @details
'* Zero arguments are keyed as an empty string.
Public Function BuildKey(ParamArray Args() As Variant) As String
    BuildKey = BuildKeyFromArray(Args)
End Function

'* Generates a key from an array in the same format as ParamArray.
'*
'* @details
'* Specify an initialized one-dimensional array for Args.
'* Non-arrays, multidimensional arrays, and uninitialized dynamic arrays are errors.
'* ObjectKeyMode is applied only to top-level object arguments.
'* Array arguments are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
Public Function BuildKeyFromArray(ByVal Args As Variant) As String
    BuildKeyFromArray = GetTypedMultiKeyByModeFromArray(pObjectKeyMode, Args)
End Function

'* Generates a key from one value.
'*
'* @details
'* ObjectKeyMode is applied only to top-level object values.
'* Array values are keyed with `G_OBJECT_KEY_MODE_REFERENCE`, so ObjectKeyMode is not applied to objects inside array elements.
Public Function BuildItemKey(ByVal Value As Variant) As String
    BuildItemKey = pBuildItemKey(Value)
End Function

Private Function pBuildItemKey(ByVal Value As Variant) As String
    Dim result_value As String
    If IsArray(Value) Then
        result_value = GetTypedValueKey(Value, G_OBJECT_KEY_MODE_REFERENCE)
    ElseIf IsObject(Value) Then
        result_value = pBuildObjectItemKey(Value)
    Else
        result_value = GetTypedValueKey(Value, G_OBJECT_KEY_MODE_REFERENCE)
    End If

    pBuildItemKey = result_value
End Function

Private Function pBuildObjectItemKey(ByVal Value As Variant) As String
    Dim effective_object_key_mode As Long
    Select Case pObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE

        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IEquatable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_I_EQUATABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Value Is Nothing Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            ElseIf TypeOf Value Is IDuplicateCheckable Then
                effective_object_key_mode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Else
                effective_object_key_mode = G_OBJECT_KEY_MODE_REFERENCE
            End If

        Case Else
            Err.Raise vbObjectError + 1, "Class TestDoubleVariantKeyBuilder.pBuildObjectItemKey", "Unsupported object key mode. (" & CStr(pObjectKeyMode) & ")"
    End Select

    pBuildObjectItemKey = GetTypedValueKey(Value, effective_object_key_mode)
End Function

Private Sub pValidateObjectKeyMode(ByVal NewValue As Long)
    Select Case NewValue
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Exit Sub
        Case Else
            Err.Raise vbObjectError + 1, "Class TestDoubleVariantKeyBuilder.pValidateObjectKeyMode", "Unsupported object key mode. (" & CStr(NewValue) & ")"
    End Select
End Sub
