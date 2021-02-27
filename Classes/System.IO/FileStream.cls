VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class FileStream - a VB6 FileStream
Implements Stream
'Implements TextReader
'Implements TextWriter
'Klasse verbindet ein PathFileName mit der Möglichkeit
'eine Datei zu öffnen, zu schreiben und zu lesen.
'In dotNET gibt es dazu drei verschiedene Klassen:
'FileStream, StreamReader und StreamWriter
'also diese Klasse am besten in dotNET implementieren,
'mit diesen dreien als Private Member
'man könnte das ganze auch mit API machen, is bissl umständlicher
'und wenn sowieso nur ASCII-Textdateien gelesen und geschrieben werden
'dann sind die VB-Funktionen schon ganz OK
'
'reizevoll wäre es aber schon sowas mit der API zu programmieren
'
'Private Declare Function CreateFileA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

'Public Enum FileAttributes
'  ReadOnly = 1             '= VBA.VbFileAttribute.vbReadOnly = 1
'  Hidden = 2               '= VBA.VbFileAttribute.vbHidden = 2
'  System = 4               '= VBA.VbFileAttribute.vbSystem = 4
'  '                        '= VBA.VbFileAttribute.vbVolume = 8
'  Directory = 16           '= VBA.VbFileAttribute.vbDirectory = 16
'  Archive = 32             '= VBA.VbFileAttribute.vbArchive = 32
'  Device = 64              '= VBA.VbFileAttribute.vbAlias = 64
'  Normal = 128             'aber VBA.VbFileAttribute.vbNormal = 0
'  Temporary = 256          '
'  SparseFile = 512         '
'  ReparsePoint = 1024      '
'  Compressed = 2048        '
'  Offline = 4096           '
'  NotContentIndexed = 8192 '
'  Encrypted = 16384        '
'End Enum

'
'Public Enum SeekOrigin
'  SeekOrigin_Begin = 0
'  SeekOrigin_current = 1
'  SeekOrigin_End = 2
'End Enum
'
'Public Enum SearchOption
'  TopDirectoryOnly = 0
'  AllDirectories = 1
'End Enum

'Public Enum FileMode
'   FileMode_CreateNew = 1
'   FileMode_Create = 2
'   FileMode_Open = 3
'   FileMode_OpenOrCreate = 4
'   FileMode_Truncate = 5
'   FileMode_Append = 6
'End Enum
Public Enum FileMode
  FileMode_Append
  FileMode_Binary
  FileMode_Input
  FileMode_Output
  FileMode_Random
End Enum
Public Enum FileAccess
  FileAccess_None
  FileAccess_Read = 1
  FileAccess_Write = 2
  FileAccess_ReadWrite = 3 'Read Or Write 'ist die Standardeinstellung
End Enum
'Public Enum FileShare
'   FileShare_None = 0
'   FileShare_Read = 1
'   FileShare_Write = 2
'   FileShare_ReadWrite = 3
'   FileShare_Inheritable = 16
'End Enum
Public Enum FileShare
  FileShare_None = 0
  FileShare_LockRead = 1
  FileShare_LockWrite = 2
  FileShare_LockReadWrite = 3
  FileShare_Shared = 16
End Enum
Private Enum ReadKind
  ReadInput      'Read one thing
  ReadLineInput  'Read one line
  ReadInputFkt   'Read the whole stuff to the end
End Enum
'Oh Shit, 5 * 3 * 4 = 60 verschiedene Möglichkeiten 'oder manche sind siher nicht möglich
Private mPath As String 'PathFileName
Private mFNr As Integer
Private mFMode As FileMode
Private mFAccess As FileAccess
Private mFShare As FileShare

Private Sub Class_Initialize()
  mFAccess = FileAccess_ReadWrite
End Sub

Private Sub Class_Terminate()
  Call cClose
End Sub

Friend Sub New_(aPath As String, Optional FMode As FileMode = FileMode_Random, Optional FAccess As FileAccess = FileAccess_ReadWrite, Optional FShare As FileShare = FileShare_None)
  'Call mPFN.NewC(aPath, aFilename, aExt)
  mPath = aPath
  mFMode = FMode
  mFAccess = FAccess
  mFShare = FShare
  Call cClose
End Sub

Private Function OOpen(ByVal aFMode As FileMode, ByVal aFAccess As FileAccess, ByVal aFShare As FileShare) As Boolean
TryE: On Error GoTo CatchE
  'aFMode = mFMode
  'aFAccess = mFAccess
  'aFShare = mFShare
  mFNr = FreeFile
  Select Case aFMode
  Case FileMode_Append
    Select Case aFAccess
    Case FileAccess_None
      Select Case aFShare
      Case FileShare_None:          Open mPath For Append As #mFNr
      Case FileShare_LockRead:      Open mPath For Append Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Append Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Append Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Append Shared As #mFNr
      End Select
'    Case FileAccess_Read 'Not Possible
'      Select Case aFShare
'      Case FileShare_None:          Open mPath For Append Access Read As #mFNr
'      Case FileShare_LockRead:      Open mPath For Append Access Read Lock Read As #mFNr
'      Case FileShare_LockWrite:     Open mPath For Append Access Read Lock Write As #mFNr
'      Case FileShare_LockReadWrite: Open mPath For Append Access Read Lock Read Write As #mFNr
'      Case FileShare_Shared:        Open mPath For Append Access Read Shared As #mFNr
'      End Select
    Case FileAccess_Write
      Select Case aFShare
      Case FileShare_None:          Open mPath For Append Access Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Append Access Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Append Access Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Append Access Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Append Access Write Shared As #mFNr
      End Select
    Case FileAccess_ReadWrite
      Select Case aFShare
      Case FileShare_None:          Open mPath For Append Access Read Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Append Access Read Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Append Access Read Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Append Access Read Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Append Access Read Write Shared As #mFNr
      End Select
    End Select
  Case FileMode_Binary
    Select Case aFAccess
    Case FileAccess_None
      Select Case aFShare
      Case FileShare_None:          Open mPath For Binary As #mFNr
      Case FileShare_LockRead:      Open mPath For Binary Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Binary Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Binary Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Binary Shared As #mFNr
      End Select
    Case FileAccess_Read
      Select Case aFShare
      Case FileShare_None:          Open mPath For Binary Access Read As #mFNr
      Case FileShare_LockRead:      Open mPath For Binary Access Read Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Binary Access Read Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Binary Access Read Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Binary Access Read Shared As #mFNr
      End Select
    Case FileAccess_Write
      Select Case aFShare
      Case FileShare_None:          Open mPath For Binary Access Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Binary Access Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Binary Access Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Binary Access Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Binary Access Write Shared As #mFNr
      End Select
    Case FileAccess_ReadWrite
      Select Case aFShare
      Case FileShare_None:          Open mPath For Binary Access Read Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Binary Access Read Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Binary Access Read Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Binary Access Read Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Binary Access Read Write Shared As #mFNr
      End Select
    End Select
  Case FileMode_Input
    Select Case aFAccess
    Case FileAccess_None
      Select Case aFShare
      Case FileShare_None:          Open mPath For Input As #mFNr
      Case FileShare_LockRead:      Open mPath For Input Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Input Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Input Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Input Shared As #mFNr
      End Select
    Case FileAccess_Read
      Select Case aFShare
      Case FileShare_None:          Open mPath For Input Access Read As #mFNr
      Case FileShare_LockRead:      Open mPath For Input Access Read Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Input Access Read Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Input Access Read Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Input Access Read Shared As #mFNr
      End Select
'    Case FileAccess_Write 'Not Possible
'      Select Case aFShare
'      Case FileShare_None:          Open mPath For Input Access Write As #mFNr
'      Case FileShare_LockRead:      Open mPath For Input Access Write Lock Read As #mFNr
'      Case FileShare_LockWrite:     Open mPath For Input Access Write Lock Write As #mFNr
'      Case FileShare_LockReadWrite: Open mPath For Input Access Write Lock Read Write As #mFNr
'      Case FileShare_Shared:        Open mPath For Input Access Write Shared As #mFNr
'      End Select
'    Case FileAccess_ReadWrite 'Not Possible
'      Select Case aFShare
'      Case FileShare_None:          Open mPath For Input Access Read Write As #mFNr
'      Case FileShare_LockRead:      Open mPath For Input Access Read Write Lock Read As #mFNr
'      Case FileShare_LockWrite:     Open mPath For Input Access Read Write Lock Write As #mFNr
'      Case FileShare_LockReadWrite: Open mPath For Input Access Read Write Lock Read Write As #mFNr
'      Case FileShare_Shared:        Open mPath For Input Access Read Write Shared As #mFNr
'      End Select
    End Select
  Case FileMode_Output
    Select Case aFAccess
    Case FileAccess_None
      Select Case aFShare
      Case FileShare_None:          Open mPath For Output As #mFNr
      Case FileShare_LockRead:      Open mPath For Output Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Output Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Output Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Output Shared As #mFNr
      End Select
'    Case FileAccess_Read
'      Select Case aFShare 'Not Possible
'      Case FileShare_None:          Open mPath For Output Access Read As #mFNr
'      Case FileShare_LockRead:      Open mPath For Output Access Read Lock Read As #mFNr
'      Case FileShare_LockWrite:     Open mPath For Output Access Read Lock Write As #mFNr
'      Case FileShare_LockReadWrite: Open mPath For Output Access Read Lock Read Write As #mFNr
'      Case FileShare_Shared:        Open mPath For Output Access Read Shared As #mFNr
'      End Select
    Case FileAccess_Write
      Select Case aFShare
      Case FileShare_None:          Open mPath For Output Access Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Output Access Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Output Access Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Output Access Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Output Access Write Shared As #mFNr
      End Select
'    Case FileAccess_ReadWrite
'      Select Case aFShare 'Not Possible
'      Case FileShare_None:          Open mPath For Output Access Read Write As #mFNr
'      Case FileShare_LockRead:      Open mPath For Output Access Read Write Lock Read As #mFNr
'      Case FileShare_LockWrite:     Open mPath For Output Access Read Write Lock Write As #mFNr
'      Case FileShare_LockReadWrite: Open mPath For Output Access Read Write Lock Read Write As #mFNr
'      Case FileShare_Shared:        Open mPath For Output Access Read Write Shared As #mFNr
'      End Select
    End Select
  Case FileMode_Random
    Select Case aFAccess
    Case FileAccess_None
      Select Case aFShare
      Case FileShare_None:          Open mPath For Random As #mFNr
      Case FileShare_LockRead:      Open mPath For Random Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Random Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Random Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Random Shared As #mFNr
      End Select
    Case FileAccess_Read
      Select Case aFShare
      Case FileShare_None:          Open mPath For Random Access Read As #mFNr
      Case FileShare_LockRead:      Open mPath For Random Access Read Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Random Access Read Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Random Access Read Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Random Access Read Shared As #mFNr
      End Select
    Case FileAccess_Write
      Select Case aFShare
      Case FileShare_None:          Open mPath For Random Access Write As #mFNr
      Case FileShare_LockRead:      Open mPath For Random Access Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Random Access Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Random Access Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Random Access Write Shared As #mFNr
      End Select
    Case FileAccess_ReadWrite
      Select Case aFShare
      Case FileShare_None:          Open mPath For Random Access Read As #mFNr
      Case FileShare_LockRead:      Open mPath For Random Access Read Write Lock Read As #mFNr
      Case FileShare_LockWrite:     Open mPath For Random Access Read Write Lock Write As #mFNr
      Case FileShare_LockReadWrite: Open mPath For Random Access Read Write Lock Read Write As #mFNr
      Case FileShare_Shared:        Open mPath For Random Access Read Write Shared As #mFNr
      End Select
    End Select
  End Select
  OOpen = True
  Exit Function
CatchE:
  MsgBox "Error: " & CStr(Err.Number) & " in PFNStream::OOpen: " & Err.Description
End Function

Public Sub cClose()
  If mFNr <> 0 Then
    Close #mFNr
    mFNr = 0
  End If
End Sub

Public Function Length() As Long
'Ruft die Länge des Streams in Bytes ab.
  Length = FileLen(mPath)
End Function
'Public Overrides Function Read(ByVal array() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
Public Function Read(ByteArray As Variant, ByVal offset As Long, ByVal Count As Long) As Long
Dim BA() As Byte
Dim afm As FileMode
Dim afa As FileAccess
Dim afs As FileShare
  Call cClose
  afm = FileMode_Binary
  afa = FileAccess_Read
  afs = FileShare_None
  If OOpen(afm, afa, afs) Then
    If offset = 0 Then offset = 1
    ReDim BA(0 To Count - 1)
    Get #mFNr, offset, BA
    ByteArray = BA()
  End If
End Function
Public Function ReadLine() As String
  ReadLine = InternalReadString(ReadLineInput)
End Function
Public Function ReadToEnd() As String
  'Read the whole file as one String with the Input()-Funktion
  ReadToEnd = InternalReadString(ReadInputFkt)
End Function

Private Function InternalReadString(ReadHow As ReadKind) As String
Dim s As String '* 1
TryE: On Error GoTo CatchE
  If IsEOF Then Exit Function
  Select Case ReadHow
  Case ReadInput
    Input #mFNr, s
    'StreamReader = s
  Case ReadLineInput
    Line Input #mFNr, s 'StreamReader
  Case ReadInputFkt
    s = Input(FileLen(mPath), #mFNr)
  End Select
  InternalReadString = s
  Exit Function
CatchE:
  MsgBox "Error: " & CStr(Err.Number) & " in PFNStream::StreamReader: " & Err.Description
End Function
Public Sub Flush()
  'alles speichern
End Sub
Public Function IsEOF() As Boolean
'Achtung bei verwendung von IsEOF wird die Datei geöffnet
  If mFNr = 0 Then If Not OOpen(mFMode, mFAccess, mFShare) Then IsEOF = True
  IsEOF = EOF(mFNr)
End Function

Public Sub wWrite(StrVal As String)
  Call InternalWriteString(False, StrVal)
End Sub
Public Sub WriteLine(StrVal As String)
  Call InternalWriteString(True, StrVal)
End Sub

Private Sub InternalWriteString(BolLine As Boolean, StrVal As String)
TryE: On Error GoTo CatchE
  If mFNr = 0 Then If Not OOpen(mFMode, mFAccess, mFShare) Then Exit Sub
  If BolLine Then
    Print #mFNr, StrVal
  Else
    Print #mFNr, StrVal;
  End If
  Exit Sub
CatchE:
  MsgBox "Error: " & CStr(Err.Number) & " in PFNStream::StreamWriter: " & Err.Description
End Sub


'######################' v   Implements Stream   v '######################'
Private Property Get Stream_CanRead() As Boolean
  Stream_CanRead = CanRead
End Property
Public Property Get CanRead() As Boolean
  CanRead = (mFAccess = FileAccess_Read) Or (mFAccess = FileAccess_ReadWrite)
End Property
Private Property Get Stream_CanSeek() As Boolean
  Stream_CanSeek = CanSeek
End Property
Public Property Get CanSeek() As Boolean
  CanSeek = True
End Property
Private Property Get Stream_CanWrite() As Boolean
  Stream_CanWrite = CanWrite
End Property
Public Property Get CanWrite() As Boolean
  CanWrite = (mFAccess = FileAccess_Write) Or (mFAccess = FileAccess_ReadWrite)
End Property
Private Sub Stream_cClose()
  Call cClose
End Sub
Private Sub Stream_Flush()
  Call Flush
End Sub
Private Property Get Stream_Length() As Long
  Stream_Length = Length
End Property
Private Property Let Stream_Position(RHS As Long)
  
End Property
Private Property Get Stream_Position() As Long
  
End Property
Private Function Stream_Read(buffer As Variant, ByVal offset As Long, ByVal Count As Long) As Long
  Stream_Read = Read(buffer, offset, Count)
End Function
Private Function Stream_ReadByte() As Long
  '
End Function
Private Sub Stream_SetLength(ByVal Value As Long)
  '
End Sub
Private Sub Stream_WriteByte(ByVal Value As Byte)
  '
End Sub
Private Sub Stream_wWrite(ByVal buffer As Variant, ByVal offset As Long, ByVal Count As Long)
  Call wWrite(CStr(buffer))
End Sub
'######################' ^   Implements Stream   ^ '######################'

'    aFS.BeginRead()
'    aFS.BeginWrite()
'    aFS.CanRead()
'    aFS.CanSeek()
'    aFS.CanWrite()
'    aFS.Close()
'    aFS.CreateObjRef()
'    aFS.EndRead()
'    aFS.EndWrite()
'    aFS.Flush()
'    aFS.GetLifetimeService()
'    aFS.GetType()
'    aFS.Handle()
'    aFS.InitializeLifetimeService()
'    aFS.IsAsync()
'    aFS.Length()
'    aFS.Lock()
'    aFS.Name()
'    aFS.Null()
'    aFS.Position()
'    aFS.Read()
'    aFS.ReadByte()
'    aFS.Seek()
'    aFS.SetLength()
'    aFS.Unlock()
'    aFS.Write()
'    aFS.WriteByte()

