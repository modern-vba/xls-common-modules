Attribute VB_Name = "Test_TextFileServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the TextFileServiceTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_GetTextFileEntity_Unregistered_CreatesInitializedTestDouble(ByVal Assert As UnitTestAssert)
    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble

    Dim actual_obj As ITextFileEntity
    Set actual_obj = service_stub.GetTextFileEntity("C:\Data\myfile.txt")

    Assert.IsTypeOf "TextFileEntityTestDouble", actual_obj
    Assert.Equals "C:\Data\myfile.txt", actual_obj.FilePath
    Assert.Equals "C:\Data\myfile.txt", service_stub.Store.GetLatestCall("GetTextFileEntity", "C:\Data\myfile.txt").GetArgument(0)
    Assert.EqualsNumeric 1, service_stub.Store.GetCallCount("GetTextFileEntity", "C:\Data\myfile.txt")
End Sub

Public Sub Test_GetTextFileEntity_WithInitializedStub_ReturnsStub(ByVal Assert As UnitTestAssert)
    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble

    Dim expect_obj As TextFileEntityTestDouble
    Set expect_obj = New TextFileEntityTestDouble
    Call expect_obj.Initialize("C:\Data\myfile.txt")
    Call service_stub.Store.SetReturn("GetTextFileEntity", expect_obj, "C:\Data\myfile.txt")

    Dim actual_obj As ITextFileEntity
    Set actual_obj = service_stub.GetTextFileEntity("C:\Data\myfile.txt")

    Assert.Equals expect_obj, actual_obj
End Sub

Public Sub Test_GetTextFileEntity_WithUninitializedStub_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble

    Dim expect_obj As TextFileEntityTestDouble
    Set expect_obj = New TextFileEntityTestDouble
    Call service_stub.Store.SetReturn("GetTextFileEntity", expect_obj, "C:\Data\myfile.txt")

    Dim actual_obj As ITextFileEntity
    Set actual_obj = service_stub.GetTextFileEntity("C:\Data\myfile.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileServiceTestDouble.")
End Sub

Public Sub Test_GetTextFileEntity_WithDifferentPathStub_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim service_stub As TextFileServiceTestDouble
    Set service_stub = New TextFileServiceTestDouble

    Dim expect_obj As TextFileEntityTestDouble
    Set expect_obj = New TextFileEntityTestDouble
    Call expect_obj.Initialize("C:\Data\other.txt")
    Call service_stub.Store.SetReturn("GetTextFileEntity", expect_obj, "C:\Data\myfile.txt")

    Dim actual_obj As ITextFileEntity
    Set actual_obj = service_stub.GetTextFileEntity("C:\Data\myfile.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileServiceTestDouble.")
End Sub
