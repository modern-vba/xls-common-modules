VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEquatable"
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
'! Interface indicating that an object can check whether it is identical to another object.
'!
'! To check equivalence, where essential characteristics match while incidental information may differ, rather than identity where all properties match, provide a member function named IsEquivalentTo instead of implementing this interface.
'! Implement IDuplicateCheckable for duplicate checks such as elements in a set.
'!
'! For equivalence, no interface is currently provided because polymorphic use is not expected; comparisons are generally between objects of the same concrete class.
'! If identity, equivalence, and duplicate checks all use the same criteria, IsEquivalentTo() and IsDuplicateOf() may be implemented by calling Equals().
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IEquatable", Description:="This class is intended for interfaces."
End Sub

'* Returns whether this object and another object are identical.
'*
'* @param OtherObject Other object.
'*
'* @details
'* Returns whether this object and another object are identical.
'* Unless there is a specific reason not to, first check that the other object has the same type as this object,
'* then return the result of comparing GetIdentityString() on this object and the other object with =.
Public Function Equals(ByVal OtherObject As Object) As Boolean
End Function

'* Returns a string used to check identity.
'*
'* @details
'* Returns a string used to check identity. Return the same string for objects judged equal and different strings for objects judged different.
'* For example, an implementation may return a string made by joining all member variables.
Public Function GetIdentityString() As String
End Function


