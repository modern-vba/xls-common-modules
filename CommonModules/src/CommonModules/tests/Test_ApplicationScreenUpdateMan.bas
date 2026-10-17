Attribute VB_Name = "Test_ApplicationScreenUpdateMan"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for the ApplicationScreenUpdateManager class.
'!
' #############################################################################

Public Sub Test_DisableUpdates_Default_DisablesAllUpdates(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager

    ' Act
    Call screen_manager.DisableUpdates

    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation

    Call screen_manager.Restore
    Set screen_manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals False, actual_screen_updating
    Assert.Equals False, actual_enable_events
    Assert.Equals False, actual_display_alerts
    Assert.EqualsNumeric xlCalculationManual, actual_calculation

CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_DisableUpdates_PartialFlags_DisablesSpecifiedSettingsOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager

    ' Act
    Call screen_manager.DisableUpdates( _
            StopScreenUpdating:=False, _
            StopEvents:=True, _
            StopDisplayAlerts:=False, _
            StopCalculation:=False)

    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation

    Call screen_manager.Restore
    Set screen_manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals True, actual_screen_updating
    Assert.Equals False, actual_enable_events
    Assert.Equals True, actual_display_alerts
    Assert.EqualsNumeric xlCalculationAutomatic, actual_calculation

CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_Restore_WithoutDisableUpdates_RestoresBackedUpSettings(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager

    ' Act
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Call screen_manager.Restore

    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation
    Set screen_manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals True, actual_screen_updating
    Assert.Equals True, actual_enable_events
    Assert.Equals True, actual_display_alerts
    Assert.EqualsNumeric xlCalculationAutomatic, actual_calculation

CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_Restore_AfterRestore_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager
    Call screen_manager.Restore
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Err.Clear

    ' Act
    Call screen_manager.Restore

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP

CLEANUP:
    Set screen_manager = Nothing
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_DisableUpdates_AfterRestore_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager
    Call screen_manager.Restore
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Err.Clear

    ' Act
    Call screen_manager.DisableUpdates

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP

CLEANUP:
    Set screen_manager = Nothing
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_Restore_NestedLifo_RestoresEachBackup(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Dim outer_manager As ApplicationScreenUpdateManager
    Set outer_manager = New ApplicationScreenUpdateManager
    Call outer_manager.DisableUpdates

    Dim inner_manager As ApplicationScreenUpdateManager
    Set inner_manager = New ApplicationScreenUpdateManager
    Call inner_manager.DisableUpdates

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    ' Act
    Call inner_manager.Restore

    Dim after_inner_screen_updating As Boolean
    Dim after_inner_enable_events As Boolean
    Dim after_inner_display_alerts As Boolean
    Dim after_inner_calculation As XlCalculation
    after_inner_screen_updating = Application.ScreenUpdating
    after_inner_enable_events = Application.EnableEvents
    after_inner_display_alerts = Application.DisplayAlerts
    after_inner_calculation = Application.Calculation

    Call outer_manager.Restore

    Dim after_outer_screen_updating As Boolean
    Dim after_outer_enable_events As Boolean
    Dim after_outer_display_alerts As Boolean
    Dim after_outer_calculation As XlCalculation
    after_outer_screen_updating = Application.ScreenUpdating
    after_outer_enable_events = Application.EnableEvents
    after_outer_display_alerts = Application.DisplayAlerts
    after_outer_calculation = Application.Calculation

    Set inner_manager = Nothing
    Set outer_manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals False, after_inner_screen_updating
    Assert.Equals False, after_inner_enable_events
    Assert.Equals False, after_inner_display_alerts
    Assert.EqualsNumeric xlCalculationManual, after_inner_calculation
    Assert.Equals True, after_outer_screen_updating
    Assert.Equals True, after_outer_enable_events
    Assert.Equals True, after_outer_display_alerts
    Assert.EqualsNumeric xlCalculationAutomatic, after_outer_calculation

CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_Restore_StatusBarChanged_DoesNotRestoreStatusBar(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_status_bar As Variant
    org_status_bar = Application.StatusBar
    Application.StatusBar = False

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager

    ' Act
    Application.StatusBar = "ASUM status bar test"
    Call screen_manager.Restore

    Dim actual_status_bar As Variant
    actual_status_bar = Application.StatusBar
    Set screen_manager = Nothing

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals "ASUM status bar test", actual_status_bar

CLEANUP:
    Application.StatusBar = org_status_bar
End Sub

Public Sub Test_ClassTerminate_NotRestored_RestoresBackedUpSettings(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager
    Call screen_manager.DisableUpdates

    ' Act
    Set screen_manager = Nothing

    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals True, actual_screen_updating
    Assert.Equals True, actual_enable_events
    Assert.Equals True, actual_display_alerts
    Assert.EqualsNumeric xlCalculationAutomatic, actual_calculation

CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub
