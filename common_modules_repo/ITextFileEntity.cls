VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITextFileEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#Interface

' #############################################################################
'!
'! @brief
'! Interface for opening, closing, and writing text files.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface ITextFileEntity", Description:="This class is intended for interfaces."
End Sub

'* File path.
Public Property Get FilePath() As String
End Property

'* Whether the file is opened for reading. (False when not open.)
Public Property Get AsRead() As Boolean
End Property

'* Whether the file is opened for writing. (False when not open.)
Public Property Get AsWrite() As Boolean
End Property

'* Whether the file is opened for appending. (False when not open.)
Public Property Get AsAppend() As Boolean
End Property

'* Whether a read lock has been acquired. (False when not open.)
Public Property Get GetReadLock() As Boolean
End Property

'* Whether a write lock has been acquired. (False when not open.)
Public Property Get GetWriteLock() As Boolean
End Property

'* Whether the file is currently open.
Public Property Get IsOpen() As Boolean
End Property

'* Whether the file is at end of file. (True when not open or when in write/append mode.)
Public Property Get IsEndOfFile() As Boolean
End Property

'* Opens a file.
'*
'* @param FilePath File path.
'* @param AsWrite [Optional] Whether to open for writing. The default is False, which opens for reading.
'* @param AsAppend [Optional] Whether to open for appending. Setting AsAppend to True when AsWrite is False fails.
'* @param GetReadLock [Optional] Whether to acquire a read lock on the file.
'* @param GetWriteLock [Optional] Whether to acquire a write lock on the file.
'*
'* @details
'* Opens a file.
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)
End Sub

'* Reads one line from the file.
'*
'* @return String that was read.
'*
'* @details
'* Reads one line of text from the file.
Public Function ReadLine() As String
End Function

'* Writes a string to the file.
'*
'* @param Expression String to write.
'*
'* @details
'* Writes a string to the file. A newline character is automatically appended to the end.
Public Sub WriteLine(ByVal Expression As String)
End Sub

'* Closes the file.
'*
'* @details
'* Closes the file. Does nothing if the file is not open.
Public Sub CloseFile()
End Sub
