VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITextFileEntity

' #############################################################################
'!
'! @brief
'! Implementation for opening, closing, and writing text files.
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pIsOpen As Boolean
Private pIsFirst As Boolean
Private pFilePath As String
Private pFileDesc As Integer
Private pAsRead As Boolean
Private pAsWrite As Boolean
Private pAsAppend As Boolean
Private pGetReadLock As Boolean
Private pGetWriteLock As Boolean

' File path.
Public Property Get FilePath() As String
    Call pCheckInit
    FilePath = pFilePath
End Property
Private Property Get ITextFileEntity_FilePath() As String
    ITextFileEntity_FilePath = FilePath
End Property

' Whether the file is opened for reading. (False when not open.)
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

' Whether the file is opened for writing. (False when not open.)
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

' Whether the file is opened for appending. (False when not open.)
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

' Whether a read lock has been acquired. (False when not open.)
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

' Whether a write lock has been acquired. (False when not open.)
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

' Whether the file is currently open.
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
    If Not pIsOpen Then
        IsEndOfFile = True
    ElseIf pAsWrite Then
        IsEndOfFile = True
    Else
        IsEndOfFile = EOF(pFileDesc)
    End If
End Property
Private Property Get ITextFileEntity_IsEndOfFile() As Boolean
    ITextFileEntity_IsEndOfFile = IsEndOfFile
End Property

'* Sets each value.
'*
'* @param FilePath File path.
'*
'* @details
'* Sets each value. An initialized instance cannot be reinitialized.
Public Sub Initialize(ByVal FilePath As String)
    If pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntity.Initialize", "An open file cannot be reinitialized. (" & pFilePath & ")"
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntity.Initialize", "The file has already been initialized. (" & pFilePath & ")"

    pFilePath = pGetAbsolutePath(FilePath)
    pIsInitialized = True
End Sub

Private Function pGetAbsolutePath(ByVal PathLikeString As String) As String
    Call pRequireFileSystemService("pGetAbsolutePath")
    pGetAbsolutePath = FsSrv.GetAbsolutePath(PathLikeString)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    If pIsOpen Then
        Close #pFileDesc
        Call pResetOpenState
    End If
    On Error GoTo 0
End Sub

' Opens the file.
'
' @param FilePath File path.
' @param AsWrite [Optional] Whether to open for writing. The default is False, which opens the file for reading.
' @param AsAppend [Optional] Whether to open for appending. Setting AsAppend to True fails when AsWrite is False.
' @param GetReadLock [Optional] Whether to acquire a read lock on the file.
' @param GetWriteLock [Optional] Whether to acquire a write lock on the file.
'
' @details
' Opens the file.
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call pCheckInit
    If pIsOpen Then Err.Raise vbObjectError, "Class TextFileEntity.OpenFile", "The file is already open. (" & pFilePath & ")"
    If AsAppend And Not AsWrite Then Err.Raise vbObjectError, "Class TextFileEntity.OpenFile", "AsAppend cannot be True when AsWrite is False."

    If Dir(pFilePath, vbNormal Or vbHidden Or vbSystem Or vbDirectory) <> "" And AsWrite And Not AsAppend Then
        If Not Force Then Err.Raise vbObjectError + 1, "Class TextFileEntity.OpenFile", "The file already exists. (" & pFilePath & ")"
    End If

    pGetReadLock = GetReadLock
    pGetWriteLock = GetWriteLock
    pAsRead = Not AsWrite
    pAsWrite = AsWrite
    pAsAppend = AsAppend

    pFileDesc = FreeFile()

    If pAsWrite Then
        If pAsAppend Then
            If pGetReadLock Then
                Open pFilePath For Append Lock Read Write As #pFileDesc
            ElseIf pGetWriteLock Then
                Open pFilePath For Append Lock Write As #pFileDesc
            Else
                Open pFilePath For Append Shared As #pFileDesc
            End If
        Else
            If pGetReadLock Then
                Open pFilePath For Output Lock Read Write As #pFileDesc
            ElseIf pGetWriteLock Then
                Open pFilePath For Output Lock Write As #pFileDesc
            Else
                Open pFilePath For Output Shared As #pFileDesc
            End If
        End If
    Else
        If pGetReadLock Then
            Open pFilePath For Input Lock Read Write As #pFileDesc
        ElseIf pGetWriteLock Then
            Open pFilePath For Input Lock Write As #pFileDesc
        Else
            Open pFilePath For Input Shared As #pFileDesc
        End If
    End If

    pIsOpen = True
End Sub
Private Sub ITextFileEntity_OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)

    Call OpenFile(AsWrite:=AsWrite, AsAppend:=AsAppend, GetReadLock:=GetReadLock, GetWriteLock:=GetWriteLock, Force:=Force)
End Sub

' Reads one line from the file.
'
' @return The string that was read.
'
' @details
' Reads one line of text from the file.
Public Function ReadLine() As String
    Call pCheckOpened
    If pAsWrite Then Err.Raise vbObjectError, "Class TextFileEntity.ReadLine", "The file is open for writing. (" & pFilePath & ")"

    Dim result_value As String
    Line Input #pFileDesc, result_value
    ReadLine = result_value
End Function
Private Function ITextFileEntity_ReadLine() As String
    ITextFileEntity_ReadLine = ReadLine()
End Function

' Writes a string to the file.
'
' @param Expression The string to write.
'
' @details
' Writes a string to the file. A line break is automatically appended to the end.
Public Sub WriteLine(ByVal Expression As String)
    Call pCheckOpened
    If pAsRead Then Err.Raise vbObjectError, "Class TextFileEntity.WriteLine", "The file is open for reading. (" & pFilePath & ")"

    Print #pFileDesc, Expression
End Sub
Private Sub ITextFileEntity_WriteLine(ByVal Expression As String)
    Call WriteLine(Expression)
End Sub

' Closes the file.
'
' @details
' Closes the file. Does nothing if the file is not open.
Public Sub CloseFile()
    Call pCheckInit
    If Not IsOpen Then Exit Sub

    Close #pFileDesc
    Call pResetOpenState
End Sub
Private Sub ITextFileEntity_CloseFile()
    Call CloseFile
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class TextFileEntity.pCheckInit", "The object is not initialized."
End Sub

Private Sub pCheckOpened()
    If Not pIsOpen Then Err.Raise vbObjectError + 1, "Class TextFileEntity.pCheckOpened", "The file is not open."
End Sub

Private Sub pResetOpenState()
    pFileDesc = 0
    pIsOpen = False
    pAsRead = False
    pAsWrite = False
    pAsAppend = False
    pGetReadLock = False
    pGetWriteLock = False
End Sub

Private Sub pRequireFileSystemService(ByVal MethodName As String)
    If FsSrv Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class TextFileEntity" & "." & MethodName, Description:="InitializeCommonService has not been called."
    End If
End Sub
