VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDuplicateCheckable"
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
'! Interface indicating that an object can check whether it is a duplicate of another object.
'!
'! Checks whether this object is a duplicate of another object based on GetKey().
'! Unless there is a specific reason not to, call GetKey() on this object and the other object, and return whether the strings match.
'!
'! To check equivalence, where essential characteristics match while incidental information may differ, rather than identity where all properties match, provide a member function named IsEquivalentTo instead of implementing this interface.
'! Implement IDuplicateCheckable for duplicate checks such as elements in a set.
'!
'! For equivalence, no interface is currently provided because polymorphic use is not expected; comparisons are generally between objects of the same concrete class.
'! If identity, equivalence, and duplicate checks all use the same criteria, IsEquivalentTo() and IsDuplicateOf() may be implemented by calling Equals().
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface IDuplicateCheckable", Description:="This class is intended for interfaces."
End Sub

'* Checks for duplication with another object.
'*
'* @param OtherObject Object to compare against.
'* @return Whether the objects are duplicates.
'*
'* @details
'* Returns whether this object and another object are duplicates.
'* Unless there is a specific reason not to, first check that the other object has the same type as this object,
'* then return the result of comparing GetKey() on this object and the other object with =.
Public Function IsDuplicateOf(ByVal OtherObject As Object) As Boolean
End Function

'* Returns a string representation used to control object duplication.
'*
'* @details
'* Returns a string representation used to control object duplication.
Public Function GetKey() As String
End Function
