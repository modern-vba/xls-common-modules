VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CounterSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Class that groups Counter objects.
'!
' #############################################################################

Private pDict As Dictionary

Private Sub Class_Initialize()
    Set pDict = New Dictionary
End Sub

'* Gets a counter by name.
'*
'* @param CounterObjectName The name of the counter object.
'* @return Counter The counter object that corresponds to the specified name.
'* @details
'* Gets the counter object with the specified name.
Public Function GetCounter(ByVal CounterObjectName As String) As Counter
    Call pCheckExists(CounterObjectName)
    Set GetCounter = pDict(CounterObjectName)
End Function

'* Adds a counter by name.
'*
'* @param CounterObjectName The name of the counter object.
'* @param CounterObject The counter object to add.
'* @details
'* Adds the specified name and counter object to the set.
Public Sub AddCounter(ByVal CounterObjectName As String, ByVal CounterObject As Counter)
    If CounterObject Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class CounterSet.AddCounter", Description:="CounterObject is Nothing."
    End If

    Call pDict.Add(CounterObjectName, CounterObject)
End Sub

'* Creates a new counter with the specified name and initial value, and adds it to the set.
'*
'* @param CounterObjectName The name of the counter object.
'* @param InitialCount Initial value of the counter (default is 0).
'* @param CountStep Count interval of the counter (default is 1).
'* @return Counter The created counter object.
'* @details
'* Creates a new counter with the specified name, initial value, and interval, and adds it to the set.
Public Function AddNewCounter(ByVal CounterObjectName As String, Optional ByVal InitialCount As Long = 0, Optional ByVal CountStep As Long = 1) As Counter
    Dim count_obj As Counter
    Set count_obj = New Counter
    Call count_obj.Initialize(InitialCount, CountStep)
    Call AddCounter(CounterObjectName, count_obj)
    Set AddNewCounter = count_obj
End Function

'* Removes a counter by name.
'*
'* @param CounterObjectName The name of the counter object.
'* @return Counter The removed counter object.
'* @details
'* Removes the counter with the specified name and returns the removed counter object.
Public Function RemoveCounter(ByVal CounterObjectName As String) As Counter
    Call pCheckExists(CounterObjectName)

    Dim result_value As Counter
    Set result_value = pDict(CounterObjectName)
    Call pDict.Remove(CounterObjectName)
    Set RemoveCounter = result_value
End Function

Private Sub pCheckExists(ByVal CounterObjectName As String)
    If Not pDict.Exists(CounterObjectName) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class CounterSet.pCheckExists", Description:="Counter does not exist. (" & CounterObjectName & ")"
    End If
End Sub
