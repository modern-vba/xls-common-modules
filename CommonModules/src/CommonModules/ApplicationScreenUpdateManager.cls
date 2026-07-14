VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationScreenUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! One-shot scope manager that temporarily saves Excel runtime settings.
'! Backs up ScreenUpdating, EnableEvents, DisplayAlerts, and Calculation when the instance is created, and Restore returns them to that state.
'! DisableUpdates is a helper for changing settings after they have been saved; Restore is valid even if DisableUpdates is not called.
'! The same instance cannot be reused after Restore. Create a new instance to save settings again.
'! When using nested instances, call Restore in LIFO order, the reverse of creation order. Restoring out of order is treated as caller error.
'! StatusBar is owned by ProgressStatus; this class does not save or restore it.
'! When the instance is destroyed, Restore is attempted as a fallback, but restore errors during destruction are suppressed.
'! If the instance is not destroyed, such as when execution is interrupted by the Esc key or the End statement, call the Restore method explicitly.
'!
' #############################################################################

Private pSavedScreenUpdating As Boolean
Private pSavedEnableEvents As Boolean
Private pSavedDisplayAlerts As Boolean
Private pSavedCalculation As Integer
Private pIsBackedUp As Boolean

'* The ScreenUpdating value recorded on this object.
Public Property Get SavedScreenUpdating() As Boolean
    Call pEnsureBackupExists
    SavedScreenUpdating = pSavedScreenUpdating
End Property

'* The EnableEvents value recorded on this object.
Public Property Get SavedEnableEvents() As Boolean
    Call pEnsureBackupExists
    SavedEnableEvents = pSavedEnableEvents
End Property

'* The DisplayAlerts value recorded on this object.
Public Property Get SavedDisplayAlerts() As Boolean
    Call pEnsureBackupExists
    SavedDisplayAlerts = pSavedDisplayAlerts
End Property

'* The Calculation value recorded on this object.
Public Property Get SavedCalculation() As Integer
    Call pEnsureBackupExists
    SavedCalculation = pSavedCalculation
End Property

Private Sub Class_Initialize()
    Call pBackup
End Sub

' Back up each setting in this object.
'
' @details
' Records the current Excel settings (ScreenUpdating, EnableEvents, DisplayAlerts, Calculation),
' so the Restore method can return them to their original settings.
Private Sub pBackup()
    pSavedScreenUpdating = Application.ScreenUpdating
    pSavedEnableEvents = Application.EnableEvents
    pSavedDisplayAlerts = Application.DisplayAlerts
    pSavedCalculation = Application.Calculation
    pIsBackedUp = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Restore
    On Error GoTo 0
End Sub

'* Stops screen updating, event firing, and similar Excel updates.
'*
'* @param StopScreenUpdating Whether to stop screen updating (default is True).
'* @param StopEvents Whether to stop event firing (default is True).
'* @param StopDisplayAlerts Whether to suppress warning dialogs (default is True).
'* @param StopCalculation Whether to switch calculation to manual mode (default is True).
'*
'* @details
'* Changes only the specified Excel settings for work, without changing the values saved when the instance was created.
'* Cannot be used on the same instance after Restore.
Public Sub DisableUpdates( _
            Optional ByVal StopScreenUpdating As Boolean = True, _
            Optional ByVal StopEvents As Boolean = True, _
            Optional ByVal StopDisplayAlerts As Boolean = True, _
            Optional ByVal StopCalculation As Boolean = True)
    Call pEnsureBackupExists
    If StopScreenUpdating Then Application.ScreenUpdating = False
    If StopEvents Then Application.EnableEvents = False
    If StopDisplayAlerts Then Application.DisplayAlerts = False
    If StopCalculation Then Application.Calculation = xlCalculationManual
End Sub

'* Explicitly restores the settings saved in this object.
'*
'* @details
'* Restores the settings saved at instance creation to Excel.
'* Even if DisableUpdates has not been called, settings changed after creation are returned to their creation-time values.
'* After restoration, the saved state is discarded, so reusing the same instance or restoring twice raises an error.
Public Sub Restore()
    Call pEnsureBackupExists
    Application.ScreenUpdating = pSavedScreenUpdating
    Application.EnableEvents = pSavedEnableEvents
    Application.DisplayAlerts = pSavedDisplayAlerts
    Application.Calculation = pSavedCalculation
    pIsBackedUp = False
End Sub

Private Sub pEnsureBackupExists()
    If Not pIsBackedUp Then
        Err.Raise vbObjectError + 1, "Class ApplicationScreenUpdateManager.pEnsureBackupExists", "No backed-up settings are available."
    End If
End Sub
