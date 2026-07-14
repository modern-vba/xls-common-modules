Attribute VB_Name = "Test_TextFileEntityTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the TextFileEntityTestDouble class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_Initialize_WithFilePath_StoresResult(ByVal Assert As UnitTestAssert)
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble

    Call file_stub.Initialize("C:\Temp\test.txt")

    Assert.Equals "C:\Temp\test.txt", file_stub.Store.GetLatestCall("Initialize", "C:\Temp\test.txt").GetArgument(0)
    Assert.Equals "C:\Temp\test.txt", file_stub.FilePath
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("Initialize", "C:\Temp\test.txt")
End Sub

Public Sub Test_FilePath_BeforeInitialize_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble

    Dim actual_path As String
    actual_path = file_stub.FilePath

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_Initialize_AlreadyInitialized_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\first.txt")
    Err.Clear

    Call file_stub.Initialize("C:\Temp\second.txt")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.Equals "C:\Temp\first.txt", file_stub.FilePath
End Sub

Public Sub Test_OpenFile_DefaultParamsAfterInitialize_RecordsCallAndState(ByVal Assert As UnitTestAssert)
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")

    Call file_stub.OpenFile

    Assert.EqualsNumeric 5, file_stub.Store.GetLatestCall("OpenFile", False, False, False, False, False).ArgumentCount
    Assert.IsTrue file_stub.AsRead
    Assert.IsFalse file_stub.AsWrite
    Assert.IsFalse file_stub.AsAppend
    Assert.IsFalse file_stub.GetReadLock
    Assert.IsFalse file_stub.GetWriteLock
    Assert.IsTrue file_stub.IsOpen
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("OpenFile", False, False, False, False, False)
End Sub

Public Sub Test_OpenFile_BeforeInitialize_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble

    Call file_stub.OpenFile

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_OpenFile_AlreadyOpen_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.OpenFile
    Err.Clear

    Call file_stub.OpenFile

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_OpenFile_AsAppendWithoutWrite_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")

    Call file_stub.OpenFile(AsWrite:=False, AsAppend:=True)

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_ReadLine_WithReadMode_ReturnsLinesAndUpdatesEof(ByVal Assert As UnitTestAssert)
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.Store.SetReturn("ReadLine", "First line", 0&)
    Call file_stub.Store.SetReturn("ReadLine", "Second line", 1&)
    Call file_stub.OpenFile

    Assert.IsFalse file_stub.IsEndOfFile
    Assert.Equals "First line", file_stub.ReadLine
    Assert.EqualsNumeric 1, file_stub.ReadCount
    Assert.IsFalse file_stub.IsEndOfFile
    Assert.Equals "Second line", file_stub.ReadLine
    Assert.EqualsNumeric 2, file_stub.ReadCount
    Assert.IsTrue file_stub.IsEndOfFile
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("ReadLine", 0&)
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("ReadLine", 1&)
End Sub

Public Sub Test_ReadLine_NotOpen_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.Store.SetReturn("ReadLine", "First line", 0&)

    Dim actual_line As String
    actual_line = file_stub.ReadLine

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_ReadLine_WriteMode_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.Store.SetReturn("ReadLine", "First line", 0&)
    Call file_stub.OpenFile(AsWrite:=True)

    Dim actual_line As String
    actual_line = file_stub.ReadLine

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_ReadLine_MissingReadCount_RaisesWithoutCreatingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.OpenFile

    Dim actual_line As String
    actual_line = file_stub.ReadLine

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
    Assert.EqualsNumeric 0, file_stub.ReadCount
    Assert.IsFalse file_stub.Store.HasReturn("ReadLine", 0&)
End Sub

Public Sub Test_WriteLine_WriteMode_RecordsLines(ByVal Assert As UnitTestAssert)
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.OpenFile(AsWrite:=True)

    Call file_stub.WriteLine("First line")
    Call file_stub.WriteLine("Second line")

    Assert.EqualsNumeric 2, file_stub.WriteCount
    Assert.Equals "First line", file_stub.Store.GetLatestCall("WriteLine", "First line").GetArgument(0)
    Assert.Equals "Second line", file_stub.Store.GetLatestCall("WriteLine", "Second line").GetArgument(0)
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("WriteLine", "First line")
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("WriteLine", "Second line")
End Sub

Public Sub Test_WriteLine_NotOpen_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")

    Call file_stub.WriteLine("Line")

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_WriteLine_ReadMode_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.OpenFile

    Call file_stub.WriteLine("Line")

    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_CloseFile_NotOpenAfterInitialize_NoErrorRecordsCall(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")

    Call file_stub.CloseFile

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse file_stub.IsOpen
    Assert.EqualsNumeric 0, file_stub.Store.GetLatestCall("CloseFile").ArgumentCount
    Assert.EqualsNumeric 1, file_stub.Store.GetCallCount("CloseFile")
End Sub

Public Sub Test_CloseFile_BeforeInitialize_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble

    Call file_stub.CloseFile

    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue StartsWith(Err.Source, "Class TextFileEntityTestDouble.")
End Sub

Public Sub Test_CloseFile_AfterRead_ResetsReadCounterAndState(ByVal Assert As UnitTestAssert)
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.Initialize("C:\Temp\test.txt")
    Call file_stub.Store.SetReturn("ReadLine", "test", 0&)
    Call file_stub.OpenFile

    Dim dummy_line As String
    dummy_line = file_stub.ReadLine

    Call file_stub.CloseFile

    Assert.EqualsNumeric 0, file_stub.ReadCount
    Assert.IsFalse file_stub.IsOpen
    Assert.IsFalse file_stub.AsRead
    Assert.IsFalse file_stub.AsWrite
End Sub
