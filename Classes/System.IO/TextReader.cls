VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public MustInherit Class TextReader
'          Inherits System.MarshalByRefObject
'     Member von: System.IO
Option Explicit

Public Property Get nNull() As TextReader
End Property

'Public Overridable Sub Close()
Public Sub cClose()
End Sub

'Public Overridable Function Peek() As Integer
Public Function Peek() As Long
End Function

'Public Overridable Function Read() As Integer
'Public Overridable Function Read(   <In(), Out()> ByVal buffer As Char(), ByVal index As Integer, ByVal count As Integer) As Integer
Public Function Read(Optional ByVal buffer As Variant, Optional ByVal index As Long = 0, Optional ByVal Count As Long = 1) As Long
End Function

'Public Overridable Function ReadBlock(ByVal buffer() As Char, ByVal index As Integer, ByVal count As Integer) As Integer
Public Function ReadBlock(ByVal buffer As Variant, ByVal index As Long, ByVal Count As Long) As Long
End Function

'Public Overridable Function ReadLine() As String
Public Function ReadLine() As String
End Function

'Public Overridable Function ReadToEnd() As String
Public Function ReadToEnd() As String
End Function

'Public Shared Function Synchronized(ByVal reader As TextReader) As TextReader
'Public Static Function Synchronized(ByVal reader As TextReader) As TextReader
'End Function


