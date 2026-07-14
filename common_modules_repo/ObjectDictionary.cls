VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class that manages a typed object dictionary whose elements can be referenced by key.
'!
' #############################################################################

Private pKeyToIndex As Object
Private pKeys As Collection
Private pValues As ObjectList
Private pHasEverAdded As Boolean
Private pNewEnumSnapshot As Collection
Private pNewEnumSnapshotVersion As Long
Private pEnumerationVersion As Long

'* Current number of items in the dictionary.
Public Property Get Count() As Long
    Count = pValues.Count
End Property

'* Comparison mode for string keys.
Public Property Get CompareMode() As Long
    CompareMode = pKeyToIndex.CompareMode
End Property
Public Property Let CompareMode(ByVal NewValue As Long)
    If pHasEverAdded Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.CompareMode", "CompareMode cannot be changed after elements have been added."
    End If

    pKeyToIndex.CompareMode = NewValue
End Property

'* Gets the key enumerator for standard For Each.
'*
'* @return Enumerator used by standard For Each.
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Or pNewEnumSnapshotVersion <> pEnumerationVersion Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
        pNewEnumSnapshotVersion = pEnumerationVersion
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* Returns an ObjectList copy holding values in key order.
'*
'* @return ObjectList copy of values.
Public Property Get Items() As ObjectList
    Set Items = pValues.CopyList()
End Property

'* Indexer for this collection. When referenced, returns the element corresponding to the specified key. When assigned, behaves like AddOrUpdate.
'*
'* @param Key Key to reference.
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Item", "The specified key does not exist. (" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))

    If pValues.IsObjectItemType Then
        Set Item = pValues.Item(item_idx)
    Else
        Item = pValues.Item(item_idx)
    End If
End Property
Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    Call AddOrUpdate(Key, Value)
End Property
Public Property Set Item(ByVal Key As Variant, ByVal Value As Object)
    Call AddOrUpdate(Key, Value)
End Property

Private Sub Class_Initialize()
    Set pKeyToIndex = CreateObject("Scripting.Dictionary")
    Set pKeys = New Collection
    Set pValues = New ObjectList
    pHasEverAdded = False
    Set pNewEnumSnapshot = Nothing
    pNewEnumSnapshotVersion = -1
    pEnumerationVersion = 0
End Sub

'* Initializes the element type contract of the dictionary.
'*
'* @param ElementTypeName Element type contract name.
'* @param RequireComparable Whether values must implement IComparable.
'* @param ObjectKeyMode Object identity / duplicate check mode.
Public Sub Initialize( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE)

    Call pValues.Initialize( _
            ElementTypeName:=ElementTypeName, _
            RequireComparable:=RequireComparable, _
            ObjectKeyMode:=ObjectKeyMode)
End Sub

'* Initializes with the same value element type state and key comparison mode as another dictionary.
'*
'* @param Other Source dictionary for initialization.
'*
'* @details
'* Inherits the value element type state and key comparison mode managed by another dictionary into this dictionary.
Public Sub InitializeFromOther(ByVal Other As ObjectDictionary)
    Call pValues.InitializeFromOther(Other.Items)
    Me.CompareMode = Other.CompareMode
End Sub

'* Adds a new key-value pair to the dictionary.
'*
'* @param Key Key to add.
'* @param Value Value to add.
Public Sub Add(ByVal Key As Variant, ByVal Value As Variant)
    If pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Add", "The specified key already exists. (" & pKeyToDisplayString(Key) & ")"
    End If

    Dim new_index As Long
    new_index = pValues.Count

    Dim key_registered As Boolean
    Dim key_stored As Boolean

    On Error GoTo ADD_ERROR
    Call pKeyToIndex.Add(Key, new_index)
    key_registered = True
    Call pAddKey(Key)
    key_stored = True
    Call pValues.Add(Value)
    pHasEverAdded = True
    Call pIncrementEnumerationVersion
    On Error GoTo 0

    Exit Sub

ADD_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If key_registered Then Call pKeyToIndex.Remove(Key)
    If key_stored And 0 < pKeys.Count Then Call pKeys.Remove(pKeys.Count)
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

Private Sub pAddKey(ByVal Key As Variant)
    If IsObject(Key) Then
        Dim key_obj As Object
        Set key_obj = Key
        Call pKeys.Add(key_obj)
    Else
        Call pKeys.Add(Key)
    End If
End Sub

'* Imports key-value pairs from another dictionary into this dictionary.
'*
'* @param OtherDictionary Other dictionary containing key-value pairs to import.
'* @param ErrorIfExists [Optional] Whether to raise an error if the key already exists.
'* @param UpdateIfExists [Optional] Whether to update the value if the key already exists.
'*
'* @details
'* Imports key-value pairs from another dictionary into this dictionary. When UpdateIfExists is True, updates values for existing keys regardless of ErrorIfExists.
Public Sub AddOther( _
        ByVal OtherDictionary As ObjectDictionary, _
        Optional ByVal ErrorIfExists As Boolean = True, _
        Optional ByVal UpdateIfExists As Boolean = False)

    If pCanInitializeFromOther Then Call InitializeFromOther(OtherDictionary)
    If ErrorIfExists And Not UpdateIfExists Then Call pCheckNoExistingKeys(OtherDictionary)

    Dim other_keys() As Variant
    other_keys = OtherDictionary.Keys

    If IsEmptyArray(other_keys) Then Exit Sub

    Dim dict_key As Variant
    For Each dict_key In other_keys
        If Me.Exists(dict_key) Then
            If UpdateIfExists Then
                Call Me.Update(dict_key, OtherDictionary.Item(dict_key))
            ElseIf ErrorIfExists Then
                Err.Raise vbObjectError + 1, "Class ObjectDictionary.AddOther", "The specified key already exists. (" & pKeyToDisplayString(dict_key) & ")"
            End If
        Else
            Call Me.Add(dict_key, OtherDictionary.Item(dict_key))
        End If
    Next dict_key
End Sub

Private Function pCanInitializeFromOther() As Boolean
    pCanInitializeFromOther = (pValues.Count = 0 And pValues.ItemTypeName = "" And Not pValues.HasItemTypeContract)
End Function

Private Sub pCheckNoExistingKeys(ByVal OtherDictionary As ObjectDictionary)
    Dim other_keys() As Variant
    other_keys = OtherDictionary.Keys

    If IsEmptyArray(other_keys) Then Exit Sub

    Dim dict_key As Variant
    For Each dict_key In other_keys
        If Me.Exists(dict_key) Then
            Err.Raise vbObjectError + 1, "Class ObjectDictionary.AddOther", "The specified key already exists. (" & pKeyToDisplayString(dict_key) & ")"
        End If
    Next dict_key
End Sub

'* Returns whether the specified key exists.
'*
'* @param Key Key to check.
'* @return True when the key exists.
Public Function Exists(ByVal Key As Variant) As Boolean
    Exists = pKeyToIndex.Exists(Key)
End Function

'* Updates the element corresponding to the specified key.
'*
'* @param Key Key to update.
'* @param Value New value.
Public Sub Update(ByVal Key As Variant, ByVal Value As Variant)
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Update", "The specified key does not exist. (" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))
    Call pValues.Update(item_idx, Value)
End Sub

'* Adds the specified key if it does not exist, or updates the value if it exists.
'*
'* @param Key Key to add or update.
'* @param Value Value to add or update.
Public Sub AddOrUpdate(ByVal Key As Variant, ByVal Value As Variant)
    If pKeyToIndex.Exists(Key) Then
        Call Update(Key, Value)
    Else
        Call Add(Key, Value)
    End If
End Sub

'* Removes the element corresponding to the specified key.
'*
'* @param Key Key to remove.
'* @return Removed element.
Public Function Remove(ByVal Key As Variant) As Variant
    If Not pKeyToIndex.Exists(Key) Then
        Err.Raise vbObjectError + 1, "Class ObjectDictionary.Remove", "The specified key does not exist. (" & pKeyToDisplayString(Key) & ")"
    End If

    Dim item_idx As Long
    item_idx = CLng(pKeyToIndex.Item(Key))

    Dim removed_value As Variant
    If pValues.IsObjectItemType Then
        Set removed_value = pValues.Remove(item_idx)
    Else
        removed_value = pValues.Remove(item_idx)
    End If

    Call pKeyToIndex.Remove(Key)
    Call pKeys.Remove(item_idx + 1)
    Call pRebuildIndexesFrom(item_idx)
    Call pIncrementEnumerationVersion

    If IsObject(removed_value) Then
        Set Remove = removed_value
    Else
        Remove = removed_value
    End If
End Function

Private Sub pRebuildIndexesFrom(ByVal StartIndex As Long)
    Dim key_idx As Long
    For key_idx = StartIndex To pKeys.Count - 1
        Dim key_value As Variant
        If IsObject(pKeys.Item(key_idx + 1)) Then
            Set key_value = pKeys.Item(key_idx + 1)
        Else
            key_value = pKeys.Item(key_idx + 1)
        End If

        pKeyToIndex.Item(key_value) = key_idx
    Next key_idx
End Sub

'* Removes all key-value pairs.
'*
'* @details
'* Preserves the value element type contract and CompareMode.
Public Sub RemoveAll()
    Call pKeyToIndex.RemoveAll
    Set pKeys = New Collection
    Call pValues.RemoveAll
    Call pIncrementEnumerationVersion
End Sub

'* Returns keys as a 0-based array in addition order.
'*
'* @return Array of keys.
Public Function Keys() As Variant()
    If pKeys.Count = 0 Then
        Keys = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    ReDim result_value(0 To pKeys.Count - 1)

    Dim key_idx As Long
    For key_idx = 1 To pKeys.Count
        If IsObject(pKeys.Item(key_idx)) Then
            Set result_value(key_idx - 1) = pKeys.Item(key_idx)
        Else
            result_value(key_idx - 1) = pKeys.Item(key_idx)
        End If
    Next key_idx

    Keys = result_value
End Function

'* Returns values as a 0-based array in key order.
'*
'* @return Array of values.
Public Function ConvertToArray() As Variant()
    ConvertToArray = pValues.ConvertToArray()
End Function

'* Returns values as a 0-based string array in key order.
'*
'* @return String array of values.
Public Function ConvertToStringArray() As String()
    ConvertToStringArray = pValues.ConvertToStringArray()
End Function

Private Sub pIncrementEnumerationVersion()
    pEnumerationVersion = pEnumerationVersion + 1
End Sub

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim key_idx As Long
    For key_idx = 1 To pKeys.Count
        If IsObject(pKeys.Item(key_idx)) Then
            Dim key_obj As Object
            Set key_obj = pKeys.Item(key_idx)
            Call result_value.Add(key_obj)
        Else
            Call result_value.Add(pKeys.Item(key_idx))
        End If
    Next key_idx

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Function pKeyToDisplayString(ByVal Key As Variant) As String
    On Error GoTo TO_STRING_ERROR

    If IsObject(Key) Then
        pKeyToDisplayString = TypeName(Key) & "@0x" & Hex(ObjPtr(Key))
    ElseIf IsArray(Key) Then
        pKeyToDisplayString = "Array"
    Else
        pKeyToDisplayString = CStr(Key)
    End If

    Exit Function

TO_STRING_ERROR:
    pKeyToDisplayString = TypeName(Key)
End Function
