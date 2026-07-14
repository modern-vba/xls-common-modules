VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! [Interface] Interface for comparable objects.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IComparable", Description:="This class is intended for interfaces."
End Sub

'* Compares objects.
'*
'* @param OtherObject Object to compare against.
'* @return Integer value that represents the comparison result.
'*
'* @details
'* Compares objects and returns the comparison result as an integer.
'* - Positive return value: the calling object is greater than the compared object.
'* - Return value of 0: both objects are equal.
'* - Negative return value: the calling object is less than the compared object.
'*
'* Here, "equal" means neither greater nor less for sorting, and does not necessarily match identity or equivalence.
'* See the IEquatable class description for identity and equivalence.
Public Function CompareTo(ByVal OtherObject As Object) As Integer
End Function

'* Determines whether this object is greater than OtherObject.
'*
'* @param OtherObject Object to compare against.
'* @return Boolean value. True when the calling object is greater than the compared object.
'*
'* @details
'* Returns True when the CompareTo function returns a positive value.
Public Function IsGreaterThan(ByVal OtherObject As Object) As Boolean
End Function

'* Determines whether this object is less than OtherObject.
'*
'* @param OtherObject Object to compare against.
'* @return Boolean value. True when the calling object is less than the compared object.
'*
'* @details
'* Returns True when the CompareTo function returns a negative value.
Public Function IsLessThan(ByVal OtherObject As Object) As Boolean
End Function

'* Determines whether this object is equal to OtherObject.
'*
'* @param OtherObject Object to compare against.
'* @return Boolean value. True when the calling object is equal to the compared object.
'*
'* @details
'* Returns True when the CompareTo function returns 0.
'* Here, "equal" means neither greater nor less for sorting, and does not necessarily match identity or equivalence.
Public Function IsEqualTo(ByVal OtherObject As Object) As Boolean
End Function
