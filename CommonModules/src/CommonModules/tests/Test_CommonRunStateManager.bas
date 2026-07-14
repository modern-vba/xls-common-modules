Attribute VB_Name = "Test_CommonRunStateManager"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Unit tests for CommonRunStateManager.
'! Executed by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_Initialize_NotInitialized_CreatesRunState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set DbgInfo = Nothing
    Set ProgStat = Nothing

    ' Act
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "DebugInformation", DbgInfo
    Assert.IsTypeOf "ProgressStatus", ProgStat

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Initialize_ExistingRunState_OverwritesRunState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set DbgInfo = New DebugInformation
    Set ProgStat = New ProgressStatus

    Dim original_dbg_info As DebugInformation
    Set original_dbg_info = DbgInfo

    Dim original_prog_stat As ProgressStatus
    Set original_prog_stat = ProgStat

    ' Act
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.NotEquals original_dbg_info, DbgInfo
    Assert.NotEquals original_prog_stat, ProgStat
    Assert.IsTypeOf "DebugInformation", DbgInfo
    Assert.IsTypeOf "ProgressStatus", ProgStat

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Initialize_ClearedRunState_CreatesRunState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager
    Call manager.Clear
    Err.Clear

    ' Act
    Call manager.Initialize

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "DebugInformation", DbgInfo
    Assert.IsTypeOf "ProgressStatus", ProgStat

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Clear_Initialized_ClearsRunState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager

    ' Act
    Call manager.Clear

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing DbgInfo
    Assert.IsNothing ProgStat

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Clear_AlreadyCleared_DoesNotRaiseError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager
    Call manager.Clear
    Err.Clear

    ' Act
    Call manager.Clear

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing DbgInfo
    Assert.IsNothing ProgStat

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Terminate_InitializedRunState_ClearsRunState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager

    ' Act
    Set manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing DbgInfo
    Assert.IsNothing ProgStat

    Set DbgInfo = Nothing
    Set ProgStat = Nothing
End Sub

Public Sub Test_Initialize_CommonServicesNotInitialized_DoesNotInitializeCommonServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing
    Set WsSrv = Nothing
    Set FsSrv = Nothing
    Set TfSrv = Nothing

    ' Act
    Dim manager As CommonRunStateManager
    Set manager = New CommonRunStateManager

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing WbSrv
    Assert.IsNothing WsSrv
    Assert.IsNothing FsSrv
    Assert.IsNothing TfSrv

    Set manager = Nothing
    Set DbgInfo = Nothing
    Set ProgStat = Nothing
    Set WbSrv = Nothing
    Set WsSrv = Nothing
    Set FsSrv = Nothing
    Set TfSrv = Nothing
End Sub
