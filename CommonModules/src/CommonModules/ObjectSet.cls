VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Provides a set for storing arbitrary objects or values and accessing them efficiently.
'! Each element is managed by a unique key.
'!
' #############################################################################

Private pDict As Dictionary
Private pTypeName As String
Private pElementTypeName As String
Private pHasItemTypeContract As Boolean
Private pRequireComparable As Boolean
Private pObjectKeyMode As Long
Private pIsObject As Boolean
Private pIsComparable As Boolean
Private pIsStringable As Boolean
Private pIsEquatable As Boolean
Private pIsDuplicateCheckable As Boolean
Private pNewEnumSnapshot As Collection
Private pNewEnumSnapshotVersion As Long
Private pEnumerationVersion As Long

'* Total number of elements in the set.
Public Property Get Count() As Long
    Count = pDict.Count
End Property

'* Type of elements in the set.
Public Property Get ItemTypeName() As String
    ItemTypeName = pTypeName
End Property

'* Whether an element type contract is set on the set.
Public Property Get HasItemTypeContract() As Boolean
    HasItemTypeContract = pHasItemTypeContract
End Property

'* Element type contract name that can be specified in Initialize.
Public Property Get ElementTypeName() As String
    ElementTypeName = pElementTypeName
End Property

'* Whether elements must implement IComparable.
Public Property Get RequireComparable() As Boolean
    RequireComparable = pRequireComparable
End Property

'* Object identity / duplicate check mode.
Public Property Get ObjectKeyMode() As Long
    ObjectKeyMode = pObjectKeyMode
End Property

'* Whether the element type is an object type.
Public Property Get IsObjectItemType() As Boolean
    IsObjectItemType = pIsObject
End Property

'* Whether the element type is treated as IComparable.
Public Property Get IsComparableItemType() As Boolean
    IsComparableItemType = pIsComparable
End Property

'* Whether the element type is treated as IStringable.
Public Property Get IsStringableItemType() As Boolean
    IsStringableItemType = pIsStringable
End Property

'* Whether the element type is treated as IEquatable.
Public Property Get IsEquatableItemType() As Boolean
    IsEquatableItemType = pIsEquatable
End Property

'* Whether the element type is treated as IDuplicateCheckable.
Public Property Get IsDuplicateCheckableItemType() As Boolean
    IsDuplicateCheckableItemType = pIsDuplicateCheckable
End Property

'* Gets the enumerator for standard For Each.
'*
'* @return Enumerator used by standard For Each.
'*
'* @details
'* Enumerates a snapshot of elements from the start of enumeration.
'* If this set is updated during standard For Each enumeration, continuation of the running enumeration is not guaranteed.
'* Use GetEnumerator for enumeration that updates or removes elements.
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If pNewEnumSnapshot Is Nothing Or pNewEnumSnapshotVersion <> pEnumerationVersion Then
        Set pNewEnumSnapshot = pCreateNewEnumSnapshot()
        pNewEnumSnapshotVersion = pEnumerationVersion
    End If

    Set NewEnum = pNewEnumSnapshot.[_NewEnum]
End Property

'* Indexer for this collection. Read-only. Returns the element at the specified index. Elements are accessed internally based on keys.
'*
'* @param ItemIndex Element index.
Public Property Get Item(ByVal ItemIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If pIsObject Then
        Set Item = pDict(pDict.Keys(ItemIndex))
    Else
        Item = pDict(pDict.Keys(ItemIndex))
    End If
End Property

Private Sub Class_Initialize()
    Set pDict = New Dictionary
    pTypeName = ""
    pElementTypeName = ""
    pHasItemTypeContract = False
    pRequireComparable = False
    pObjectKeyMode = G_OBJECT_KEY_MODE_REFERENCE
    pIsObject = False
    pIsComparable = False
    pIsEquatable = False
    pIsDuplicateCheckable = False
    pIsStringable = False
    pEnumerationVersion = 0
    pNewEnumSnapshotVersion = -1
    Set pNewEnumSnapshot = Nothing
End Sub

'* Initializes the element type contract of the set.
'*
'* @param ElementTypeName Element type contract name.
'* @param RequireComparable Whether elements must implement IComparable.
'* @param ObjectKeyMode Object identity / duplicate check mode.
'*
'* @details
'* Initializes this as a set that accepts only object-type elements.
'* ElementTypeName is compared with the self-reported ElementTypeKey or TypeName.
Public Sub Initialize( _
        ByVal ElementTypeName As String, _
        Optional ByVal RequireComparable As Boolean = False, _
        Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE)

    If pDict.Count <> 0 Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Initialize", "An initialized set cannot be reinitialized."
    End If
    If Not IsValidElementTypeKey(ElementTypeName) Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Initialize", "ElementTypeName must contain only characters allowed in class module names. (" & ElementTypeName & ")"
    End If

    Call pCheckObjectKeyMode(ObjectKeyMode)
    pElementTypeName = ElementTypeName
    pTypeName = "Object@" & ElementTypeName
    pHasItemTypeContract = True
    pRequireComparable = RequireComparable
    pObjectKeyMode = ObjectKeyMode
    pIsObject = True
    pIsComparable = RequireComparable
    pIsEquatable = (ObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE)
    pIsDuplicateCheckable = (ObjectKeyMode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE)
End Sub

'* Initializes with the same element type state as another set.
'*
'* @param Other Source set for initialization.
'*
'* @details
'* Inherits the element type state managed by another set into this set.
Public Sub InitializeFromOther(ByVal Other As ObjectSet)
    If pDict.Count <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.InitializeFromOther", "An initialized set cannot be reinitialized."
    End If

    pTypeName = Other.ItemTypeName
    pElementTypeName = Other.ElementTypeName
    pHasItemTypeContract = Other.HasItemTypeContract
    pRequireComparable = Other.RequireComparable
    pObjectKeyMode = Other.ObjectKeyMode
    pIsObject = Other.IsObjectItemType
    pIsComparable = Other.IsComparableItemType
    pIsStringable = Other.IsStringableItemType
    pIsEquatable = Other.IsEquatableItemType
    pIsDuplicateCheckable = Other.IsDuplicateCheckableItemType
End Sub

'* Initializes with the same element type state as another list.
'*
'* @param OtherList Source list for initialization.
'*
'* @details
'* Inherits the element type state managed by another list into this set.
Public Sub InitializeFromList(ByVal OtherList As ObjectList)
    If pDict.Count <> 0 Or pTypeName <> "" Or pHasItemTypeContract Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.InitializeFromList", "An initialized set cannot be reinitialized."
    End If

    pTypeName = OtherList.ItemTypeName
    pElementTypeName = OtherList.ElementTypeName
    pHasItemTypeContract = OtherList.HasItemTypeContract
    pRequireComparable = OtherList.RequireComparable
    pObjectKeyMode = OtherList.ObjectKeyMode
    pIsObject = OtherList.IsObjectItemType
    pIsComparable = OtherList.IsComparableItemType
    pIsStringable = OtherList.IsStringableItemType
    pIsEquatable = OtherList.IsEquatableItemType
    pIsDuplicateCheckable = OtherList.IsDuplicateCheckableItemType
End Sub

Private Sub pCheckObjectKeyMode(ByVal ObjectKeyMode As Long)
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectKeyMode", "Unsupported object key mode. (" & CStr(ObjectKeyMode) & ")"
    End Select
End Sub

'* Updates the element at the specified index with a new object.
'*
'* @param ItemIndex Index of the element to update.
'* @param ItemObject New element.
'*
'* @details
'* Updates the element at the specified index with a new object. Raises an error if the key changes.
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call pCheckItemType(ItemObject)

    Dim old_key As Variant
    If IsObject(pDict.Keys(ItemIndex)) Then
        Set old_key = pDict.Keys(ItemIndex)
    Else
        old_key = pDict.Keys(ItemIndex)
    End If

    Dim new_key As Variant
    Call pGetKey(new_key, ItemObject)

    If Not pKeyEquals(old_key, new_key) Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.Update", "The original object key and updated object key differ. If this change is required, Remove the original element and then Add the new element. (old_key: " & old_key & " / new_key: " & new_key & ")"
        Exit Sub
    End If

    If pIsObject Then
        Set pDict(old_key) = ItemObject
    Else
        pDict(old_key) = ItemObject
    End If
    Call pIncrementEnumerationVersion
End Sub

Private Function pKeyEquals(ByVal ItemKey1 As Variant, ByVal ItemKey2 As Variant) As Boolean
    Dim result_value As Boolean
    result_value = (CStr(ItemKey1) = CStr(ItemKey2))

    pKeyEquals = result_value
End Function

'* Returns the set as an array.
'*
'* @return Elements in array form.
'*
'* @details
'* Returns all elements of the set in array form. If the element count is 0, returns a 0-based empty array.
Public Function ConvertToArray() As Variant()
    If pDict.Count < 1 Then
        ConvertToArray = EmptyVariantArray()
        Exit Function
    End If

    Dim result_value() As Variant
    Dim item_idx As Long

    ReDim result_value(0 To pDict.Count - 1)

    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set result_value(item_idx) = pDict(dict_key)
        Else
            result_value(item_idx) = pDict(dict_key)
        End If

        item_idx = item_idx + 1
    Next dict_key

    ConvertToArray = result_value
End Function

'* Returns the set as a string array.
'*
'* @return Array in string form.
'*
'* @details
'* Returns all elements of the set as an array in string form. If the element count is 0, returns a 0-based empty array.
Public Function ConvertToStringArray() As String()
    If pDict.Count < 1 Then
        ConvertToStringArray = EmptyStringArray()
        Exit Function
    End If

    Dim result_value() As String
    Dim item_idx As Long
    Dim stringable_obj As IStringable
    Dim dict_item_obj As Object

    ReDim result_value(0 To pDict.Count - 1)

    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set dict_item_obj = pDict(dict_key)
            If dict_item_obj Is Nothing Then
                result_value(item_idx) = GetTypedValueKey(dict_item_obj)
            Else
                If TypeOf dict_item_obj Is IStringable Then
                    Set stringable_obj = dict_item_obj
                    result_value(item_idx) = stringable_obj.ToString()
                Else
                    result_value(item_idx) = GetTypedValueKey(dict_item_obj)
                End If
            End If
        Else
            If IsError(pDict(dict_key)) Then
                result_value(item_idx) = ExcelErrorToString(pDict(dict_key))
            ElseIf IsArray(pDict(dict_key)) Then
                result_value(item_idx) = GetTypedValueKey(pDict(dict_key))
            Else
                result_value(item_idx) = CStr(pDict(dict_key))
            End If
        End If
        item_idx = item_idx + 1
    Next dict_key

    ConvertToStringArray = result_value
End Function

'* Adds a new element to the set.
'*
'* @param ItemObject Element to add.
'* @param Optional ErrorIfExists Whether to raise an error if it already exists.
'* @return Boolean value indicating whether the element was added.
'*
'* @details
'* Adds a new element to the set. Optionally specifies behavior when the element already exists.
Public Function Add(ByVal ItemObject As Variant, Optional ByVal ErrorIfExists As Boolean = True) As Boolean
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not ErrorIfExists Then
        If pDict.Exists(key_var) Then
            Add = False
            Exit Function
        End If
    End If

    Call pDict.Add(key_var, ItemObject)
    Call pIncrementEnumerationVersion
    Add = True
End Function

Private Sub pInitializeItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    Dim item_is_obj As Boolean
    item_is_obj = IsObject(ItemObject)

    If item_is_obj Then
        pIsObject = True
        If Not ItemObject Is Nothing Then
            pElementTypeName = pGetElementTypeName(ItemObject)
            pTypeName = "Object@" & pElementTypeName
            pHasItemTypeContract = True
            If TypeOf ItemObject Is IComparable Then
                pIsComparable = True
            End If
            If TypeOf ItemObject Is IEquatable Then
                pIsEquatable = True
                pObjectKeyMode = G_OBJECT_KEY_MODE_I_EQUATABLE
            End If
            If TypeOf ItemObject Is IDuplicateCheckable Then
                pIsDuplicateCheckable = True
                pIsEquatable = False
                pObjectKeyMode = G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            End If
            If TypeOf ItemObject Is IStringable Then
                pIsStringable = True
            End If
        End If
    Else
        If pIsObject Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.pInitializeItemType", "The first element was an object, but the new element is not an object. (" & pTypeName & ")"
        Else
            ' Primitive type.
            pTypeName = GetTypeString(ItemObject)
            pHasItemTypeContract = True
        End If
    End If
End Sub

'* Adds each element of an array to the set.
'*
'* @param ItemArray Array of elements to add.
'* @param Optional ErrorIfExists Whether to raise an error if an element already exists.
'*
'* @details
'* Adds each element of an array to the set. Optionally specifies behavior when an element already exists.
Public Sub AddArray(ByVal ItemArray As Variant, Optional ByVal ErrorIfExists As Boolean = True)
    If Not IsArray(ItemArray) Then Err.Raise vbObjectError + 1, "Class ObjectSet.AddArray", "The argument is not an array."
    If IsEmptyArray(ItemArray) Then Exit Sub

    Dim key_var As Variant
    Dim set_item_obj As IDuplicateCheckable
    Dim item_obj As Variant

    For Each item_obj In ItemArray
        Call Add(item_obj, ErrorIfExists)
    Next item_obj
End Sub

'* Imports elements from another set into this set.
'*
'* @param OtherSet Other set containing elements to import.
'* @param Optional ErrorIfExists Whether to raise an error if an element already exists.
'*
'* @details
'* Imports elements from another set into this set. Optionally specifies behavior when an element already exists.
Public Sub AddOther(ByVal OtherSet As ObjectSet, Optional ByVal ErrorIfExists As Boolean = True)
    If pTypeName = "" And OtherSet.HasItemTypeContract Then
        Call InitializeFromOther(OtherSet)
    Else
        Call pPrepareToAddObjectContract( _
                OtherSet.HasItemTypeContract, _
                OtherSet.ElementTypeName, _
                OtherSet.RequireComparable, _
                OtherSet.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(item_idx), ErrorIfExists:=ErrorIfExists)
    Next item_idx
End Sub

'* Imports elements from another list into this set.
'*
'* @param OtherList Other list containing elements to import.
'* @param Optional ErrorIfExists Whether to raise an error if an element already exists.
'*
'* @details
'* Imports elements from another list into this set. Optionally specifies behavior when an element already exists.
Public Sub AddList(ByVal OtherList As ObjectList, Optional ByVal ErrorIfExists As Boolean = True)
    If pTypeName = "" And OtherList.HasItemTypeContract Then
        Call InitializeFromList(OtherList)
    Else
        Call pPrepareToAddObjectContract( _
                OtherList.HasItemTypeContract, _
                OtherList.ElementTypeName, _
                OtherList.RequireComparable, _
                OtherList.ObjectKeyMode)
    End If

    Dim item_idx As Long
    For item_idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(item_idx), ErrorIfExists)
    Next item_idx
End Sub

'* Creates a duplicate of this set.
'*
'* @return New duplicated set.
'*
'* @details
'* Creates a duplicate of this set.
Public Function CopySet() As ObjectSet
    Dim result_value As ObjectSet
    Set result_value = New ObjectSet
    Call result_value.InitializeFromOther(Me)

    Dim key_var As Variant
    For Each key_var In pDict.Keys
        Call result_value.Add(pDict(key_var))
    Next key_var

    Set CopySet = result_value
End Function

'* Removes the element at the specified index.
'*
'* @param Index Index of the element to remove.
'* @return Removed element.
'*
'* @details
'* Removes the element at the specified index.
Public Function Remove(ByVal Index As Long) As Variant
    Dim old_key As Variant
    Dim old_obj As Variant

    If pIsObject Then
        Set old_obj = pDict.Item(pDict.Keys(Index))
    Else
        old_obj = pDict.Item(pDict.Keys(Index))
    End If

    If pIsObject Then
        Set Remove = RemoveItem(old_obj)
    Else
        Remove = RemoveItem(old_obj)
    End If
End Function

'* Removes the specified element.
'*
'* @param ItemObject Element to remove.
'* @param Optional IgnoreNotExists Whether to ignore when the element does not exist.
'* @return Removed element.
'*
'* @details
'* Removes the element judged duplicate of the specified element. Optionally specifies behavior when the element does not exist.
Public Function RemoveItem(ByVal ItemObject As Variant, Optional ByVal IgnoreNotExists As Boolean = True) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.RemoveItem", "Attempted to remove an element that does not exist."
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.RemoveItem", "Attempted to remove an element that does not exist."
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Call pDict.Remove(key_var)
        Call pIncrementEnumerationVersion
        Set RemoveItem = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)
    Call pDict.Remove(key_var)
    Call pIncrementEnumerationVersion

    RemoveItem = result_value
End Function

'* Removes all elements of the set.
'*
'* @details
'* Removes all elements of the set. Preserves the element type state.
Public Sub RemoveAll()
    If pDict.Count = 0 Then Exit Sub
    Call pDict.RemoveAll
    Call pIncrementEnumerationVersion
End Sub

'* Checks whether the argument element is already included.
'*
'* @param ItemObject Element to check.
'* @return Boolean value indicating whether the element is included.
'*
'* @details
'* Checks whether an element that duplicates the specified object is included.
Public Function Exists(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    If Not pCanSearchItem(ItemObject) Then
        Exists = result_value
        Exit Function
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    result_value = pDict.Exists(key_var)

    Exists = result_value
End Function

'* Returns the element considered duplicate of the argument.
'*
'* @param ItemObject Element to check.
'* @return Element considered duplicate.
'*
'* @details
'* Returns the element considered duplicate of the argument. Returns Nothing when the element does not exist.
Public Function FindItem(ByVal ItemObject As Variant) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set FindItem = Nothing
            Exit Function
        Else
            FindItem = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set FindItem = Nothing
            Exit Function
        Else
            FindItem = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Set FindItem = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)

    FindItem = result_value
End Function

'* Sorts the contents of this set.
'*
'* @param Optional Descending Whether to sort in descending order.
'*
'* @details
'* Sorts the contents of this set.
'* Because this uses ObjectList.Sort() internally, see that method for behavior details.
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.InitializeFromSet(Me)
    Call obj_list.AddSet(Me)
    Call obj_list.Sort(Descending:=Descending)
    Call Me.RemoveAll
    Call Me.AddList(obj_list)
End Sub

'* Gets the enumerator for this set.
'*
'* @param Descending [Optional] Whether to enumerate in descending order.
'* @param IsReadOnly [Optional] Whether to make the enumerator read-only.
'* @return Enumerator object.
'*
'* @details
'* Gets the enumerator for this set.
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result_value
End Function

Private Function pCreateNewEnumSnapshot() As Collection
    Dim result_value As Collection
    Set result_value = New Collection

    Dim enum_obj As IEnumerator
    Set enum_obj = Me.GetEnumerator(IsReadOnly:=True)
    Do While enum_obj.MoveNext()
        Call result_value.Add(enum_obj.Current)
    Loop

    Set pCreateNewEnumSnapshot = result_value
End Function

Private Sub pIncrementEnumerationVersion()
    pEnumerationVersion = pEnumerationVersion + 1
End Sub

Private Sub pPrepareToAddObjectContract( _
        ByVal SourceHasItemTypeContract As Boolean, _
        ByVal SourceElementTypeName As String, _
        ByVal SourceRequireComparable As Boolean, _
        ByVal SourceObjectKeyMode As Long)

    If Not SourceHasItemTypeContract Then Exit Sub
    If SourceElementTypeName = "" Then Exit Sub

    If Not pHasItemTypeContract Then
        Call Initialize( _
                ElementTypeName:=SourceElementTypeName, _
                RequireComparable:=SourceRequireComparable, _
                ObjectKeyMode:=SourceObjectKeyMode)
        Exit Sub
    End If

    If StrComp(pElementTypeName, SourceElementTypeName, vbTextCompare) <> 0 _
            Or pRequireComparable <> SourceRequireComparable _
            Or pObjectKeyMode <> SourceObjectKeyMode Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.pPrepareToAddObjectContract", "The source and destination element type contracts differ."
    End If
End Sub

Private Function pGetElementTypeName(ByVal ItemObject As Object) As String
    Dim result_value As String
    If TypeOf ItemObject Is IElementTypeProvider Then
        Dim element_type_provider As IElementTypeProvider
        Set element_type_provider = ItemObject
        result_value = element_type_provider.ElementTypeKey
        If Not IsValidElementTypeKey(result_value) Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.pGetElementTypeName", "ElementTypeKey must contain only characters allowed in class module names. (" & result_value & ")"
        End If
    Else
        result_value = TypeName(ItemObject)
    End If

    pGetElementTypeName = result_value
End Function

Private Sub pCheckObjectItemAbilities(ByVal ItemObject As Object)
    If pRequireComparable And Not TypeOf ItemObject Is IComparable Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "Attempted to add an object that does not implement IComparable to a RequireComparable=True set. (" & pTypeName & ")"
    End If

    Select Case pObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Not TypeOf ItemObject Is IEquatable Then
                Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "Attempted to add an object that does not implement IEquatable to a set in IEquatable mode. (" & pTypeName & ")"
            End If
        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Not TypeOf ItemObject Is IDuplicateCheckable Then
                Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "Attempted to add an object that does not implement IDuplicateCheckable to a set in IDuplicateCheckable mode. (" & pTypeName & ")"
            End If
        Case Else
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckObjectItemAbilities", "Unsupported object key mode. (" & CStr(pObjectKeyMode) & ")"
    End Select
End Sub

Private Function pCanSearchItem(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pTypeName <> "" Then
        Call pCheckItemType(ItemObject)
        result_value = True
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not pIsObject Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not IsObject(ItemObject) Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If ItemObject Is Nothing Then
        result_value = True
    Else
        result_value = False
    End If

    pCanSearchItem = result_value
End Function

Private Sub pCheckItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pIsObject Then
        If Not IsObject(ItemObject) Then
            Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckItemType", "The first element was an object, but the new element is not an object. (" & pTypeName & ")"
        End If
        If ItemObject Is Nothing Then Exit Sub
    End If

    Dim item_type_name As String
    If pIsObject Then
        item_type_name = GetTypeString(ItemObject, UseElementTypeKey:=True)
    Else
        item_type_name = GetTypeString(ItemObject)
    End If

    If StrComp(item_type_name, pTypeName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckItemType", "The item type differs from the initially set element type. (list: " & pTypeName & ", item: " & item_type_name & ")"
    End If

    If pIsObject Then
        Call pCheckObjectItemAbilities(ItemObject)
    End If
End Sub

Private Sub pGetKey(ByRef ObjectKey As Variant, ByVal ObjectItem As Variant)
    Call pCheckUnsupportedSpecialValue(ObjectItem)

    ObjectKey = GetTypedValueKey(ObjectItem, pObjectKeyMode, UseElementTypeKey:=True)
End Sub

Private Sub pCheckUnsupportedSpecialValue(ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then Exit Sub
    If IsArray(ItemObject) Then Exit Sub
    If IsEmpty(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckUnsupportedSpecialValue", "Empty cannot be used as an element."
    If IsNull(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet.pCheckUnsupportedSpecialValue", "Null cannot be used as an element."
End Sub

Private Function pErrorValuesEqual(ByVal ErrorValue1 As Variant, ByVal ErrorValue2 As Variant) As Boolean
    Dim error_values As Dictionary
    Set error_values = New Dictionary
    Call error_values.Add(ErrorValue1, True)
    pErrorValuesEqual = error_values.Exists(ErrorValue2)
End Function
