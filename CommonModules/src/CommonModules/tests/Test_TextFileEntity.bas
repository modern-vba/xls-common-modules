Attribute VB_Name = "Test_TextFileEntity"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the TextFileEntity class.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_Initialize_UninitializedFsSrv_RaisesExplicitError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set FsSrv = Nothing

    Dim text_file As TextFileEntity
    Set text_file = New TextFileEntity

    ' Act
    Call text_file.Initialize("C:\codex_text_file_test.txt")

    Dim actual_err_num As Long
    Dim actual_err_source As String
    Dim actual_err_desc As String
    actual_err_num = Err.Number
    actual_err_source = Err.Source
    actual_err_desc = Err.Description

    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, actual_err_num, actual_err_source, actual_err_desc) Then Exit Sub
    Assert.Equals "Class TextFileEntity.pGetAbsolutePath", actual_err_source
    Assert.Equals "InitializeCommonService has not been called.", actual_err_desc
End Sub
