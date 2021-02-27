VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public MustInherit Class Stream
'          Inherits System.MarshalByRefObject
'     Member von: System.IO
Option Explicit
'Klasse ist abstrakt gibt also nur das Verhalten aller Streams vor
'Public Overridable Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult
'Public Function BeginRead(ByVal buffer As variant, ByVal offset As long, ByVal count As long, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult



'Public Overridable Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult
'Public Function BeginWrite(ByVal buffer As variant, ByVal offset As long, ByVal count As long, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult

'Public Overridable MustOverride ReadOnly Property CanRead() As Boolean
Public Property Get CanRead() As Boolean
End Property

'Public Overridable MustOverride ReadOnly Property CanSeek() As Boolean
Public Property Get CanSeek() As Boolean
End Property

'Public Overridable MustOverride ReadOnly Property CanWrite() As Boolean
Public Property Get CanWrite() As Boolean
End Property

'Public Overridable Sub Close()
Public Sub cClose()
End Sub

'Public Overridable Function EndRead(ByVal asyncResult As System.IAsyncResult) As Integer
'Public Function EndRead(ByVal asyncResult As System.IAsyncResult) As Long
'
'End Function

'Public Overridable Sub EndWrite(ByVal asyncResult As System.IAsyncResult)
'Public Sub EndWrite(ByVal asyncResult As System.IAsyncResult)
'
'End Sub

'Public Overridable MustOverride Sub Flush()
Public Sub Flush()
End Sub

'Public Overridable MustOverride ReadOnly Property Length() As Long
Public Property Get Length() As Long
End Property

'Public ReadOnly Shared Dim Null As System.IO.Stream

'Public Overridable MustOverride Property Position() As Long
Public Property Get Position() As Long
End Property
Public Property Let Position(PosVal As Long)
End Property

'Public Overridable MustOverride Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
Public Function Read(buffer As Variant, ByVal offset As Long, ByVal Count As Long) As Long
End Function

'Public Overridable Function ReadByte() As Integer
Public Function ReadByte() As Long
End Function

'Public Overridable MustOverride Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
'Public Function Seek(ByVal offset As Long, ByVal origin As SeekOrigin) As Long

'Public Overridable MustOverride Sub SetLength(ByVal value As Long)
Public Sub SetLength(ByVal value As Long)
End Sub

'Public Overridable MustOverride Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
Public Sub wWrite(ByVal buffer As Variant, ByVal offset As Long, ByVal Count As Long)
End Sub

'Public Overridable Sub WriteByte(ByVal value As Byte)
Public Sub WriteByte(ByVal value As Byte)
End Sub

'Protected Overridable Function CreateWaitHandle() As System.Threading.WaitHandle
'Function CreateWaitHandle() As WaitHandle
'
'End Function


