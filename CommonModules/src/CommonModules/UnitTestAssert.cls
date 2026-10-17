VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Assert class for unit tests.
'!
' #############################################################################

Private Const C_ARRAY_COMPARISON_ERROR As String = "Array arguments are not supported by scalar assertions. Use EqualsArray or NotEqualsArray."
Private Const C_NUMERIC_COMPARISON_ERROR As String = "Numeric assertions require Byte, Integer, Long, LongLong, Single, or Double values."
Private Const C_ARRAY_ASSERTION_ARGUMENT_ERROR As String = "Array assertions require array arguments."

Private pCallCount As Long
Private pIsFailed As Boolean
Private pResultMessage As String

'* Whether it differed from the assertion.
Public Property Get IsFailed() As Boolean
    IsFailed = pIsFailed
End Property

'* Details when it differed from the assertion.
Public Property Get ResultMessage() As String
    ResultMessage = pResultMessage
End Property

'* Number of assertions executed.
Public Property Get AssertionCount() As Long
    AssertionCount = pCallCount
End Property

Private Sub Class_Initialize()
    pCallCount = 0
    pIsFailed = False
    pResultMessage = ""
End Sub

'* Expected value and actual value are equal.
'*
'* @param ExpectedValue Expected value.
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected value and actual value are equal.
'* Compares not only values but also types.
Public Sub Equals(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasArrayArgument(ExpectedValue, ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_ARRAY_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, False)

    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected value and actual value are not equal.
'*
'* @param ExpectedValue Expected value.
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected value and actual value differ.
'* Compares not only values but also types.
Public Sub NotEquals(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasArrayArgument(ExpectedValue, ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_ARRAY_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, False)

    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected number and actual number are equal.
'*
'* @param ExpectedValue Expected value.
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected number and actual number are equal.
'* Does not check type as long as the values are numeric.
'* Note: Currency and Decimal are not currently included in type-ignoring numeric comparison targets.
Public Sub EqualsNumeric(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasArrayArgument(ExpectedValue, ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_ARRAY_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    If Not pIsNumericComparisonValue(ExpectedValue) Or Not pIsNumericComparisonValue(ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_NUMERIC_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, True)

    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected number and actual number are not equal.
'*
'* @param ExpectedValue Expected value.
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected number and actual number are not equal.
'* Does not check type as long as the values are numeric.
'* Note: Currency and Decimal are not currently included in type-ignoring numeric comparison targets.
Public Sub NotEqualsNumeric(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasArrayArgument(ExpectedValue, ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_ARRAY_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    If Not pIsNumericComparisonValue(ExpectedValue) Or Not pIsNumericComparisonValue(ActualValue) Then
        Call pSetUnsupportedComparisonResult(ExpectedValue, ActualValue, C_NUMERIC_COMPARISON_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, True)

    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected object and actual object are equal under the specified ObjectKeyMode.
'*
'* @param ExpectedObject Expected object.
'* @param ActualObject Actual object.
'* @param ObjectKeyMode Object comparison method.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Compares with ObjectKeyMode specified as one of `G_OBJECT_KEY_MODE_REFERENCE`, `G_OBJECT_KEY_MODE_I_EQUATABLE`, or `G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE`.
Public Sub EqualsObject(ByVal ExpectedObject As Object, ByVal ActualObject As Object, Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsObjectCore(is_equal, expected_str, actual_str, ExpectedObject, ActualObject, ObjectKeyMode)

    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected object and actual object are not equal under the specified ObjectKeyMode.
'*
'* @param ExpectedObject Expected object.
'* @param ActualObject Actual object.
'* @param ObjectKeyMode Object comparison method.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Compares with ObjectKeyMode specified as one of `G_OBJECT_KEY_MODE_REFERENCE`, `G_OBJECT_KEY_MODE_I_EQUATABLE`, or `G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE`.
Public Sub NotEqualsObject(ByVal ExpectedObject As Object, ByVal ActualObject As Object, Optional ByVal ObjectKeyMode As Long = G_OBJECT_KEY_MODE_REFERENCE, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean

    Call pEqualsObjectCore(is_equal, expected_str, actual_str, ExpectedObject, ActualObject, ObjectKeyMode)

    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* True is expected.
'*
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the result is true.
Public Sub IsTrue(ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    Call Equals(True, ActualValue, CaseName:=CaseName)
End Sub

'* False is expected.
'*
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the result is false.
Public Sub IsFalse(ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    Call Equals(False, ActualValue, CaseName:=CaseName)
End Sub

'* Expected type and actual type are equal.
'*
'* @param ExpectedTypeName String of the expected type name.
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected type and actual type are equal.
'* Checks whether the types match, not whether casting is possible.
Public Sub IsTypeOf(ByVal ExpectedTypeName As String, ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If TypeName(ActualValue) <> ExpectedTypeName Then
        Call pSetResultMessage("<" & ExpectedTypeName & ">", "<" & TypeName(ActualValue) & ">", CaseName:=CaseName)
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Nothing is expected.
'*
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Nothing is expected.
'* Same as IsTypeOf("Nothing", ActualValue).
Public Sub IsNothing(ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    Call IsTypeOf("Nothing", ActualValue, CaseName:=CaseName)
End Sub

'* Empty is expected.
'*
'* @param ActualValue Actual value.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Empty is expected.
'* Same as IsTypeOf("Empty", ActualValue).
Public Sub IsEmpty(ByVal ActualValue As Variant, Optional ByVal CaseName As String = "")
    Call IsTypeOf("Empty", ActualValue, CaseName:=CaseName)
End Sub

'* The specified error is expected to occur.
'*
'* @param ExpectedCode Expected error code. Passing 0 targets all errors.
'* @param ActualCodeCode Actual error code. Pass Err.Number.
'* @param ErrorSource Actual error source. Pass Err.Source.
'* @param ErrorDescription Actual error details. Pass Err.Description.
'* @param CaseName Case name to display on failure.
'* @param ClearAfterAssert When True, runs Err.Clear after inspection.
'* @return Whether an error occurred.
'*
'* @details
'* Indicates that the specified error is expected to occur.
'* When 0 is specified as ExpectedCode, indicates that some error is expected to occur.
'* When a nonzero value is specified, indicates that that error is expected to occur (other errors are ignored).
Public Function ErrorRaised(ByVal ExpectedCode As Long, ByVal ActualCode As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String, Optional ByVal CaseName As String = "", Optional ByVal ClearAfterAssert As Boolean = False) As Boolean
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Function

    Dim expected_str As String
    Dim actual_str As String
    Dim is_error As Boolean

    Call pErrorRaisedCore(is_error, expected_str, actual_str, ExpectedCode, ActualCode, ErrorSource, ErrorDescription, False)

    If is_error Then
        ErrorRaised = True
    Else
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        ErrorRaised = False
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
    If ClearAfterAssert Then Err.Clear
End Function

'* The specified error is expected not to occur.
'*
'* @param ExpectedCode Expected error code. Passing 0 targets all errors.
'* @param ActualCodeCode Actual error code. Pass Err.Number.
'* @param ErrorSource Actual error source. Pass Err.Source.
'* @param ErrorDescription Actual error details. Pass Err.Description.
'* @param CaseName Case name to display on failure.
'* @param ClearAfterAssert When True, runs Err.Clear after inspection.
'* @return Whether an error did not occur.
'*
'* @details
'* Indicates that the specified error is expected not to occur.
'* When 0 is specified as ExpectedCode, indicates that no error of any kind is expected to occur.
'* When a nonzero value is specified, indicates that that error is not expected to occur (other errors may occur).
Public Function ErrorNotRaised(ByVal ExpectedCode As Long, ByVal ActualCode As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String, Optional ByVal CaseName As String = "", Optional ByVal ClearAfterAssert As Boolean = False) As Boolean
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Function

    Dim expected_str As String
    Dim actual_str As String
    Dim is_error As Boolean

    Call pErrorRaisedCore(is_error, expected_str, actual_str, ExpectedCode, ActualCode, ErrorSource, ErrorDescription, True)

    If is_error Then
        Call pSetResultMessage(expected_str, actual_str, CaseName:=CaseName)
        ErrorNotRaised = False
    Else
        ErrorNotRaised = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
    If ClearAfterAssert Then Err.Clear
End Function

'* Expected array and actual array are equal.
'*
'* @param ExpectedValue Expected array.
'* @param ActualValue Actual array.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected array and actual array are equal.
'* Compares not only values but also types.
Public Sub EqualsArray(ByVal ExpectedArray As Variant, ByVal ActualArray As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasNonArrayArgument(ExpectedArray, ActualArray) Then
        Call pSetUnsupportedComparisonResult(ExpectedArray, ActualArray, C_ARRAY_ASSERTION_ARGUMENT_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim additional_str As String
    Dim is_equal As Boolean

    Call pEqualsArrayCore(is_equal, expected_str, actual_str, additional_str, ExpectedArray, ActualArray, True)

    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str, AdditionalInformation:=additional_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

'* Expected array and actual array are not equal.
'*
'* @param ExpectedValue Expected array.
'* @param ActualValue Actual array.
'* @param CaseName Case name to display on failure.
'*
'* @details
'* Indicates that the expected array and actual array differ.
'* Compares not only values but also types.
Public Sub NotEqualsArray(ByVal ExpectedArray As Variant, ByVal ActualArray As Variant, Optional ByVal CaseName As String = "")
    ' Do nothing if the assertion has already failed.
    If pIsFailed Then Exit Sub

    If pHasNonArrayArgument(ExpectedArray, ActualArray) Then
        Call pSetUnsupportedComparisonResult(ExpectedArray, ActualArray, C_ARRAY_ASSERTION_ARGUMENT_ERROR, CaseName)
        Exit Sub
    End If

    Dim expected_str As String
    Dim actual_str As String
    Dim additional_str As String
    Dim is_equal As Boolean

    Call pEqualsArrayCore(is_equal, expected_str, actual_str, additional_str, ExpectedArray, ActualArray, True)

    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str, AdditionalInformation:=additional_str, CaseName:=CaseName)
        pIsFailed = True
    End If

    ' Processing that always runs.
    pCallCount = pCallCount + 1
End Sub

Private Sub pEqualsObjectCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedObject As Object, _
        ByVal ActualObject As Object, _
        ByVal ObjectKeyMode As Long)

    Call pValidateObjectAssertionMode(ObjectKeyMode)
    Call pRaiseIfObjectKeyModeInterfaceMissing(ExpectedObject, ObjectKeyMode, "ExpectedObject")
    Call pRaiseIfObjectKeyModeInterfaceMissing(ActualObject, ObjectKeyMode, "ActualObject")

    ExpectedValueString = GetTypedValueKey(ExpectedObject, ObjectKeyMode)
    ActualValueString = GetTypedValueKey(ActualObject, ObjectKeyMode)

    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE
            CheckResult = pReferenceObjectsEqual(ExpectedObject, ActualObject)

        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If ExpectedObject Is Nothing Or ActualObject Is Nothing Then
                CheckResult = pReferenceObjectsEqual(ExpectedObject, ActualObject)
            Else
                Dim expected_eq As IEquatable
                Set expected_eq = ExpectedObject
                CheckResult = expected_eq.Equals(ActualObject)
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If ExpectedObject Is Nothing Or ActualObject Is Nothing Then
                CheckResult = pReferenceObjectsEqual(ExpectedObject, ActualObject)
            Else
                Dim expected_dup As IDuplicateCheckable
                Set expected_dup = ExpectedObject
                CheckResult = expected_dup.IsDuplicateOf(ActualObject)
            End If
    End Select
End Sub

Private Function pReferenceObjectsEqual(ByVal ExpectedObject As Object, ByVal ActualObject As Object) As Boolean
    If ExpectedObject Is Nothing Or ActualObject Is Nothing Then
        pReferenceObjectsEqual = (ExpectedObject Is Nothing) And (ActualObject Is Nothing)
    Else
        pReferenceObjectsEqual = (ExpectedObject Is ActualObject)
    End If
End Function

Private Sub pValidateObjectAssertionMode(ByVal ObjectKeyMode As Long)
    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_REFERENCE, G_OBJECT_KEY_MODE_I_EQUATABLE, G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            Exit Sub
        Case Else
            Err.Raise vbObjectError + 1, "Class UnitTestAssert.pValidateObjectAssertionMode", "Unsupported object comparison mode. (" & CStr(ObjectKeyMode) & ")"
    End Select
End Sub

Private Sub pRaiseIfObjectKeyModeInterfaceMissing(ByVal TargetObject As Object, ByVal ObjectKeyMode As Long, ByVal TargetName As String)
    If TargetObject Is Nothing Then Exit Sub

    Select Case ObjectKeyMode
        Case G_OBJECT_KEY_MODE_I_EQUATABLE
            If Not TypeOf TargetObject Is IEquatable Then
                Err.Raise vbObjectError + 1, "Class UnitTestAssert.pRaiseIfObjectKeyModeInterfaceMissing", TargetName & " is not IEquatable. (" & TypeName(TargetObject) & ")"
            End If

        Case G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE
            If Not TypeOf TargetObject Is IDuplicateCheckable Then
                Err.Raise vbObjectError + 1, "Class UnitTestAssert.pRaiseIfObjectKeyModeInterfaceMissing", TargetName & " is not IDuplicateCheckable. (" & TypeName(TargetObject) & ")"
            End If
    End Select
End Sub

Private Sub pEqualsCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant, _
        ByVal IgnoreNumericType As Boolean)

    Dim actual_is_object As Boolean
    Dim expected_is_object As Boolean

    expected_is_object = IsObject(ExpectedValue)
    actual_is_object = IsObject(ActualValue)

    If pHasArrayArgument(ExpectedValue, ActualValue) Then
        ExpectedValueString = pAssertValueString(ExpectedValue)
        ActualValueString = pAssertValueString(ActualValue)
        CheckResult = False
        Exit Sub
    End If

    If expected_is_object And Not actual_is_object _
            Or Not expected_is_object And actual_is_object Then

        ExpectedValueString = "Entity of " & TypeName(ExpectedValue)
        ActualValueString = "Entity of " & TypeName(ActualValue)
        CheckResult = False
        Exit Sub
    End If

    ' By this point, cases where one side is primitive and the other is an object have already returned early.

    If actual_is_object Then
        ' Both are object types.

        Dim actual_is_nothing As Boolean
        Dim expected_is_nothing As Boolean

        actual_is_nothing = (ActualValue Is Nothing)
        expected_is_nothing = (ExpectedValue Is Nothing)

        If actual_is_nothing Then
            ActualValueString = "Nothing"
        Else
            ActualValueString = "Entity of " & TypeName(ActualValue)
        End If

        If expected_is_nothing Then
            ExpectedValueString = "Nothing"
        Else
            ExpectedValueString = "Entity of " & TypeName(ExpectedValue)
        End If

        If expected_is_nothing Or actual_is_nothing Then
            If expected_is_nothing And actual_is_nothing Then
                CheckResult = True
            Else
                CheckResult = False
            End If
            Exit Sub
        End If

        ' By this point, cases where either side or both sides are Nothing have already returned early.

        If Not ExpectedValue Is ActualValue Then
            CheckResult = False
        Else
            CheckResult = True
        End If
    Else
        ' Both are primitive types.
        If IgnoreNumericType Then
            Call pEqualsPrimitiveNumericCore(CheckResult, ExpectedValueString, ActualValueString, ExpectedValue, ActualValue)
        Else
            Call pEqualsPrimitiveCore(CheckResult, ExpectedValueString, ActualValueString, ExpectedValue, ActualValue)
        End If
    End If
End Sub

Private Function pHasArrayArgument(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant) As Boolean
    pHasArrayArgument = IsArray(ExpectedValue) Or IsArray(ActualValue)
End Function

Private Function pHasNonArrayArgument(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant) As Boolean
    pHasNonArrayArgument = Not IsArray(ExpectedValue) Or Not IsArray(ActualValue)
End Function

Private Sub pSetUnsupportedComparisonResult( _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant, _
        ByVal AdditionalInformation As String, _
        Optional ByVal CaseName As String = "")

    Call pSetResultMessage(pAssertValueString(ExpectedValue), pAssertValueString(ActualValue), AdditionalInformation, CaseName)
    pCallCount = pCallCount + 1
End Sub

Private Function pAssertValueString(ByVal TargetValue As Variant) As String
    If IsObject(TargetValue) Then
        If TargetValue Is Nothing Then
            pAssertValueString = "Nothing"
        Else
            pAssertValueString = "Entity of " & TypeName(TargetValue)
        End If
    ElseIf IsArray(TargetValue) Then
        pAssertValueString = TypeName(TargetValue) & "<Array>"
    Else
        pAssertValueString = pPrimitiveValueString(TargetValue)
    End If
End Function

Private Function pIsNumericComparisonValue(ByVal TargetValue As Variant) As Boolean
    If IsObject(TargetValue) Then Exit Function
    If IsArray(TargetValue) Then Exit Function
    If pIsSpecialPrimitive(TargetValue) Then Exit Function

    pIsNumericComparisonValue = pIsNumericType(TypeName(TargetValue))
End Function

Private Sub pEqualsPrimitiveCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant)


    Dim expect_type As String
    expect_type = TypeName(ExpectedValue)

    Dim actual_type As String
    actual_type = TypeName(ActualValue)

    ExpectedValueString = pPrimitiveValueString(ExpectedValue)
    ActualValueString = pPrimitiveValueString(ActualValue)

    If pIsSpecialPrimitive(ExpectedValue) Or pIsSpecialPrimitive(ActualValue) Then
        CheckResult = pSpecialPrimitivesEqual(ExpectedValue, ActualValue)
        Exit Sub
    End If

    If expect_type = "String" And Not actual_type = "String" _
            Or Not expect_type = "String" And actual_type = "String" Then
        CheckResult = False
        Exit Sub
    End If

    If ExpectedValue = ActualValue Then
        If TypeName(ExpectedValue) <> TypeName(ActualValue) Then
            CheckResult = False
        Else
            CheckResult = True
        End If
    Else
        CheckResult = False
    End If
End Sub

Private Sub pEqualsPrimitiveNumericCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant)

    ExpectedValueString = pPrimitiveValueString(ExpectedValue)
    ActualValueString = pPrimitiveValueString(ActualValue)

    If Not pIsNumericComparisonValue(ExpectedValue) Or Not pIsNumericComparisonValue(ActualValue) Then
        CheckResult = False
        Exit Sub
    End If

    If pIsSpecialPrimitive(ExpectedValue) Or pIsSpecialPrimitive(ActualValue) Then
        CheckResult = False
        Exit Sub
    End If

    If ExpectedValue = ActualValue Then
        If pIsNumericType(TypeName(ExpectedValue)) And pIsNumericType(TypeName(ActualValue)) Then
            CheckResult = True
        Else
            CheckResult = False
        End If
    Else
        CheckResult = False
    End If
End Sub

Private Function pPrimitiveValueString(ByVal TargetValue As Variant) As String
    If VBA.IsError(TargetValue) Then
        pPrimitiveValueString = "CVErr(" & CStr(CLng(TargetValue)) & ")<" & TypeName(TargetValue) & ">"
    ElseIf VBA.IsNull(TargetValue) Then
        pPrimitiveValueString = "Null<" & TypeName(TargetValue) & ">"
    ElseIf VBA.IsEmpty(TargetValue) Then
        pPrimitiveValueString = "<Empty>"
    Else
        pPrimitiveValueString = CStr(TargetValue) & "<" & TypeName(TargetValue) & ">"
    End If
End Function

Private Function pIsSpecialPrimitive(ByVal TargetValue As Variant) As Boolean
    pIsSpecialPrimitive = VBA.IsError(TargetValue) Or VBA.IsNull(TargetValue) Or VBA.IsEmpty(TargetValue)
End Function

Private Function pSpecialPrimitivesEqual(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant) As Boolean
    If VBA.IsError(ExpectedValue) Or VBA.IsError(ActualValue) Then
        If VBA.IsError(ExpectedValue) And VBA.IsError(ActualValue) Then
            pSpecialPrimitivesEqual = CLng(ExpectedValue) = CLng(ActualValue)
        Else
            pSpecialPrimitivesEqual = False
        End If
    ElseIf VBA.IsNull(ExpectedValue) Or VBA.IsNull(ActualValue) Then
        pSpecialPrimitivesEqual = VBA.IsNull(ExpectedValue) And VBA.IsNull(ActualValue)
    ElseIf VBA.IsEmpty(ExpectedValue) Or VBA.IsEmpty(ActualValue) Then
        pSpecialPrimitivesEqual = VBA.IsEmpty(ExpectedValue) And VBA.IsEmpty(ActualValue)
    Else
        pSpecialPrimitivesEqual = False
    End If
End Function

Private Function pIsNumericType(ByVal TypeNameString As String) As Boolean
    Select Case TypeNameString
     Case "Byte"
        pIsNumericType = True
     Case "Short"
        pIsNumericType = True
     Case "Integer"
        pIsNumericType = True
     Case "Long"
        pIsNumericType = True
     Case "LongLong"
        pIsNumericType = True
     Case "Single"
        pIsNumericType = True
     Case "Double"
        pIsNumericType = True
     Case Else
        pIsNumericType = False
    End Select
End Function

Private Sub pErrorRaisedCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedCode As Variant, _
        ByVal ActualCode As Variant, _
        ByVal ErrorSource As String, _
        ByVal ErrorDescription As String, _
        ByVal AsNotRaised As Boolean)

    If ActualCode = 0 Then
        CheckResult = False
        If ExpectedCode = 0 Then
            ExpectedValueString = "Some error is raised."
            ActualValueString = "Any error is not raised."
            CheckResult = False
        Else
            ExpectedValueString = "The error [0x" & Hex(ExpectedCode) & "] is raised."
            ActualValueString = "Any error is not raised."
            CheckResult = False
        End If
    Else
        If ExpectedCode = 0 Then
            ' When the expected code is 0, any raised error is OK.
            If AsNotRaised Then
                ExpectedValueString = "Any error is not raised."
                ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
                CheckResult = True
            Else
                ExpectedValueString = "Some error is raised."
                ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
                CheckResult = True
            End If
        Else
            ExpectedValueString = "The error [0x" & Hex(ExpectedCode) & "] is raised."
            ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
            If ActualCode = ExpectedCode Then
                CheckResult = True
            Else
                CheckResult = False
            End If
        End If
    End If
End Sub

Private Sub pEqualsArrayCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByRef AdditionalInformation As String, _
        ByVal ExpectedArray As Variant, _
        ByVal ActualArray As Variant, _
        ByVal IgnoreNumericType As Boolean)

    ExpectedValueString = TypeName(ExpectedArray)
    ActualValueString = TypeName(ActualArray)

    ' Fail if the types differ.
    If ExpectedValueString <> ActualValueString Then
        CheckResult = False
        Exit Sub
    End If

    ' Fail if either value is not an array.
    If Not IsArray(ExpectedArray) Or Not IsArray(ActualArray) Then
        CheckResult = False
        Exit Sub
    End If

    Dim expect_is_empty As Boolean
    expect_is_empty = IsEmptyArray(ExpectedArray)

    Dim actual_is_empty As Boolean
    actual_is_empty = IsEmptyArray(ActualArray)

    If expect_is_empty Or actual_is_empty Then
        CheckResult = expect_is_empty And actual_is_empty
        If Not CheckResult Then
            If expect_is_empty Then
                ExpectedValueString = "empty array"
            End If
            If actual_is_empty Then
                ActualValueString = "empty array"
            End If
        End If
        Exit Sub
    End If

    ' Get the array LBound and UBound values.
    Dim expect_lb_arr() As Long
    Dim expect_ub_arr() As Long
    Call GetArrayBounds(expect_lb_arr, expect_ub_arr, ExpectedArray)

    Dim actual_lb_arr() As Long
    Dim actual_ub_arr() As Long
    Call GetArrayBounds(actual_lb_arr, actual_ub_arr, ActualArray)

    ' Calculate the number of array dimensions.
    Dim expect_dimension As Long
    expect_dimension = UBound(expect_lb_arr) - LBound(expect_lb_arr) + 1

    Dim actual_dimension As Long
    actual_dimension = UBound(actual_lb_arr) - LBound(actual_lb_arr) + 1

    ' Compare the number of array dimensions.
    If expect_dimension <> actual_dimension Then
        ExpectedValueString = "dimension is " & expect_dimension
        ActualValueString = "dimension is " & actual_dimension
        CheckResult = False
        Exit Sub
    End If

    ' Compare LBound and UBound for each dimension.
    Dim item_idx As Long
    For item_idx = LBound(expect_lb_arr) To UBound(expect_lb_arr)
        If expect_lb_arr(item_idx) <> actual_lb_arr(item_idx) Then
            ExpectedValueString = "lbound of dimension " & (item_idx + 1) & " is " & expect_lb_arr(item_idx)
            ActualValueString = "lbound of dimension " & (item_idx + 1) & " is " & actual_lb_arr(item_idx)
            CheckResult = False
            Exit Sub
        End If

        If expect_ub_arr(item_idx) <> actual_ub_arr(item_idx) Then
            ExpectedValueString = "ubound of dimension " & (item_idx + 1) & " is " & expect_ub_arr(item_idx)
            ActualValueString = "ubound of dimension " & (item_idx + 1) & " is " & actual_ub_arr(item_idx)
            CheckResult = False
            Exit Sub
        End If
    Next item_idx

    ' Convert expect to a one-dimensional array.
    Dim expect_item As Variant
    Dim expect_arr() As Variant
    item_idx = 0
    For Each expect_item In ExpectedArray
        ReDim Preserve expect_arr(0 To item_idx)
        If IsObject(expect_item) Then
            Set expect_arr(item_idx) = expect_item
        Else
            expect_arr(item_idx) = expect_item
        End If
        item_idx = item_idx + 1
    Next expect_item

    ' Compare all elements with Equals.
    ' Note: array comparison requires an exact match including types, so element comparison does not ignore numeric type differences regardless of IgnoreNumericType.
    Dim actual_item As Variant
    item_idx = 0
    For Each actual_item In ActualArray
        Dim check_result As Boolean
        Dim expect_str As String
        Dim actual_str As String
        Call pEqualsCore(check_result, expect_str, actual_str, expect_arr(item_idx), actual_item, False)
        If Not check_result Then
            ExpectedValueString = expect_str
            ActualValueString = actual_str
            AdditionalInformation = "@" & item_idx
            Exit Sub
        End If

        item_idx = item_idx + 1
    Next actual_item

    CheckResult = True
End Sub

'Private Function pEqualsMultidimensionalArrayCore( _
'        ByRef TargetArray() As Variant, _
'        ByRef CurrentDimension As Integer, _
'        ByRef MaxDimension As Integer, _
'        ByVal ExpectedArray As Variant, _
'        ByVal ActualArray As Variant, _
'        ByVal IgnoreNumericType As Boolean)
'
'    Dim expected_value_string As String
'    Dim actual_value_string As String
'    expected_value_string = TypeName(ExpectedArray)
'    actual_value_string = TypeName(ActualArray)
'
'    If expected_value_string <> actual_value_string Then
'        CheckResult = False
'        Exit Function
'    End If
'
'    If Not IsArray(ExpectedArray) Or Not IsArray(ActualArray) Then
'        CheckResult = False
'        Exit Function
'    End If
'
'    Dim expect_dimension As Long
'    expect_dimension = pGetArrayDimension(ExpectedArray)
'
'    Dim actual_dimension As Long
'    actual_dimension = pGetArrayDimension(ActualArray)
'
'    If expect_dimension <> actual_dimension Then
'        expected_value_string = "dimension is " & expect_dimension
'        actual_value_string = "dimension is " & actual_dimension
'        CheckResult = False
'        Exit Function
'    End If
'
'    'CheckResult = True
'End Function

Private Sub pSetResultMessage(ByVal ExpectedValueString As String, ByVal ActualValueString As String, Optional ByVal AdditionalInformation As String = "", Optional ByVal CaseName As String = "")
    Dim additional_info As String
    additional_info = AdditionalInformation

    If CaseName <> "" Then
        If additional_info <> "" Then additional_info = additional_info & " | "
        additional_info = additional_info & "case: " & CaseName
    End If

    If additional_info = "" Then
        pResultMessage = "[" & pCallCount & "] expected: " & ExpectedValueString & " | actual: " & ActualValueString
    Else
        pResultMessage = "[" & pCallCount & "] expected: " & ExpectedValueString & " | actual: " & ActualValueString & " | " & additional_info
    End If
    pIsFailed = True
End Sub
