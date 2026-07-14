VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileEntityTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileEntity

' #############################################################################
'!
'! @brief
'! Lightweight fake for opening, closing, and writing text files.
'!
'! @details
'! Does not operate on a real file; manages only initialization and open/read/write/close lifecycle state.
'! ReadLine is a strict stub; explicitly register read lines in Store.
'! OpenFile / WriteLine / CloseFile do not handle arbitrary error injection by Store.SetError.
'! To test arbitrary file open failure, configure an error on TextFileServiceTestDouble.GetTextFileEntity or test the target process separately.
'!
' #############################################################################

Private pFilePath As String
Private pIsInitialized As Boolean
Private pIsOpen As Boolean
Private pAsRead As Boolean
Private pAsWrite As Boolean
Private pAsAppend As Boolean
Private pGetReadLock As Boolean
Private pGetWriteLock As Boolean

'* Read count. Incremented by ReadLine and reset by CloseFile.
Public ReadCount As Long

'* Write count. Incremented by WriteLine. Not reset.
Public WriteCount As Long

'* Stub settings and history of successfully completed calls.
'* State violation errors from this lightweight fake are not recorded in call history.
Public Store As TestDoubleBehaviorStore

'* File path.
Public Property Get FilePath() As String
    Call pCheckInit
    FilePath = pFilePath
End Property
Private Property Get ITextFileEntity_FilePath() As String
    ITextFileEntity_FilePath = FilePath
End Property

'* Whether opened for reading. (False when not open.)
Public Property Get AsRead() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsRead = pAsRead
    Else
        AsRead = False
    End If
End Property
Private Property Get ITextFileEntity_AsRead() As Boolean
    ITextFileEntity_AsRead = AsRead
End Property

'* Whether opened for writing. (False when not open.)
Public Property Get AsWrite() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsWrite = pAsWrite
    Else
        AsWrite = False
    End If
End Property
Private Property Get ITextFileEntity_AsWrite() As Boolean
    ITextFileEntity_AsWrite = AsWrite
End Property

'* Whether opened for appending. (False when not open.)
Public Property Get AsAppend() As Boolean
    Call pCheckInit
    If pIsOpen Then
        AsAppend = pAsAppend
    Else
        AsAppend = False
    End If
End Property
Private Property Get ITextFileEntity_AsAppend() As Boolean
    ITextFileEntity_AsAppend = AsAppend
End Property

'* Whether a read lock has been acquired. (False when not open.)
Public Property Get GetReadLock() As Boolean
    Call pCheckInit
    If pIsOpen Then
        GetReadLock = pGetReadLock
    Else
        GetReadLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetReadLock() As Boolean
    ITextFileEntity_GetReadLock = GetReadLock
End Property

'* Whether a write lock has been acquired. (False when not open.)
Public Property Get GetWriteLock() As Boolean
    Call pCheckInit
    If pIsOpen Then
        GetWriteLock = pGetWriteLock
    Else
        GetWriteLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetWriteLock() As Boolean
    ITextFileEntity_GetWriteLock = GetWriteLock
End Property

'* Whether the file is currently open.
Public Property Get IsOpen() As Boolean
    Call pCheckInit
    IsOpen = pIsOpen
End Property
Private Property Get ITextFileEntity_IsOpen() As Boolean
    ITextFileEntity_IsOpen = IsOpen
End Property

' Whether the file is at end of file. (True when not open or when in write/append mode.)
Public Property Get IsEndOfFile() As Boolean
    Call pCheckInit
    If Not pIsOpen Or pAsWrite Then
        IsEndOfFile = True
    Else
        IsEndOfFile = Not Store.HasReturn("ReadLine", ReadCount)
    End If
End Property
Private Property Get ITextFileEntity_IsEndOfFile() As Boolean
    ITextFileEntity_IsEndOfFile = IsEndOfFile
End Property

Private Sub Class_Initialize()
    Set Store = New TestDoubleBehaviorStore
End Sub

'* spy only for Initialize. Records the initialization path.
Public Sub Initialize(ByVal FilePath As String)
    If pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble.Initialize", "An open file cannot be reinitialized. (" & pFilePath & ")"
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble.Initialize", "The file has already been initialized. (" & pFilePath & ")"

    Call Store.RecordCall("Initialize", FilePath)
    pFilePath = FilePath
    pIsInitialized = True
End Sub

'* Lightweight fake for OpenFile. Updates state based on arguments and records the call.
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call pCheckInit
    If pIsOpen Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble.OpenFile", "The file is already open. (" & pFilePath & ")"
    If AsAppend And Not AsWrite Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble.OpenFile", "AsAppend cannot be True when AsWrite is False."

    Call Store.RecordCall("OpenFile", AsWrite, AsAppend, GetReadLock, GetWriteLock, Force)

    pAsRead = Not AsWrite
    pAsWrite = AsWrite
    pAsAppend = AsAppend
    pGetReadLock = GetReadLock
    pGetWriteLock = GetWriteLock
    pIsOpen = True
End Sub
Private Sub ITextFileEntity_OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call OpenFile(AsWrite:=AsWrite, AsAppend:=AsAppend, GetReadLock:=GetReadLock, GetWriteLock:=GetWriteLock, Force:=Force)
End Sub

'* Strict stub for ReadLine. Explicitly register read lines in Store.
Public Function ReadLine() As String
    Call pCheckOpened
    If pAsWrite Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble.ReadLine", "The file is open for writing. (" & pFilePath & ")"

    If Not Store.HasReturn("ReadLine", ReadCount) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TextFileEntityTestDouble.ReadLine", Description:="No read line has been registered. (read_count: " & ReadCount & ")"
    End If

    Dim result_value As String
    result_value = Store.GetReturn("ReadLine", ReadCount)
    Call Store.RecordCallWithKeyArgs("ReadLine", Array(ReadCount))
    ReadCount = ReadCount + 1

    ReadLine = result_value
End Function
Private Function ITextFileEntity_ReadLine() As String
    ITextFileEntity_ReadLine = ReadLine()
End Function

'* spy only for WriteLine. Records written lines and updates WriteCount based on state.
Public Sub WriteLine(ByVal Expression As String)
    Call pCheckOpened
    If pAsRead Then Err.Raise vbObjectError, "Class TextFileEntityTestDouble.WriteLine", "The file is open for reading. (" & pFilePath & ")"

    Call Store.RecordCall("WriteLine", Expression)
    WriteCount = WriteCount + 1
End Sub
Private Sub ITextFileEntity_WriteLine(ByVal Expression As String)
    Call WriteLine(Expression)
End Sub

'* Lightweight fake for CloseFile. Records it as a 0-argument call and resets state if open.
Public Sub CloseFile()
    Call pCheckInit
    Call Store.RecordCall("CloseFile")

    If Not pIsOpen Then Exit Sub

    ReadCount = 0
    Call pResetOpenState
End Sub
Private Sub ITextFileEntity_CloseFile()
    Call CloseFile
End Sub

Private Sub pResetOpenState()
    pIsOpen = False
    pAsRead = False
    pAsWrite = False
    pAsAppend = False
    pGetReadLock = False
    pGetWriteLock = False
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble.pCheckInit", "The object is not initialized."
End Sub

Private Sub pCheckOpened()
    Call pCheckInit
    If Not pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntityTestDouble.pCheckOpened", "The file is not open."
End Sub
