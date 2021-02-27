VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StreamReader
'          Inherits System.IO.TextReader
'     Member von: System.IO
Option Explicit
Implements TextReader
Private Const DefaultBufferSize     As Long = 1024
Private Const DefaultFileBufferSize As Long = 4096
Private Const MinimumBufferSize     As Long = 128
'    //
'    // The input buffer
Private mInputBuffer() As Byte 'input_buffer;
'    //
'    // The decoded buffer from the above input buffer
Private mDecodedBuffer '() 'As Byte '    char [] decoded_buffer;
'    //
'    // Decoded bytes in decoded_buffer.
Private mDecodedCount As Long '    int 'decoded_count
'    //
'    // Current position in the decoded_buffer
Private mPos As Long 'pos;
'    //
'    // The buffer size that we are using
Private mBufferSize As Long 'buffer_size;
Private mDoChecks As Long   'do_checks;
'Private mEncoding As Encoding 'Encoding encoding;
'Private mDecoder As Decoder ''    Decoder decoder;
Private mMayBlock As Boolean
Private mLineBuilder As StringBuilder 'line_builder;
Private mBaseStream As Stream 'base_stream;
Private miEncoding As Long '1 = Byte, 2 = Integer

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  '
End Sub

'Public Sub New(ByVal path As String)
'Public Sub New(ByVal path As String, ByVal detectEncodingFromByteOrderMarks As Boolean)
'Public Sub New(ByVal path As String, ByVal encoding As System.Text.Encoding)
'Public Sub New(ByVal path As String, ByVal encoding As System.Text.Encoding, ByVal detectEncodingFromByteOrderMarks As Boolean)
'Public Sub New(ByVal path As String, ByVal encoding As System.Text.Encoding, ByVal detectEncodingFromByteOrderMarks As Boolean, ByVal bufferSize As Integer)
Public Sub NewP(ByVal aPath As String, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long)
    Set mBaseStream = MNew.FileStream(aPath)
    Call InternalReadBaseStream
End Sub

'Public Sub New(ByVal stream As System.IO.Stream)
'Public Sub New(ByVal stream As System.IO.Stream, ByVal detectEncodingFromByteOrderMarks As Boolean)
'Public Sub New(ByVal stream As System.IO.Stream, ByVal encoding As System.Text.Encoding)
'Public Sub New(ByVal stream As System.IO.Stream, ByVal encoding As System.Text.Encoding, ByVal detectEncodingFromByteOrderMarks As Boolean)
'Public Sub New(ByVal stream As System.IO.Stream, ByVal encoding As System.Text.Encoding, ByVal detectEncodingFromByteOrderMarks As Boolean, ByVal bufferSize As Integer)
Friend Sub New_(ByVal aStream As Stream, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long)
    Set mBaseStream = aStream
    'Without decoding
    Call InternalReadBaseStream
End Sub

Private Sub InternalReadBaseStream()
    'Dim iEnc As Long
    
    'Without decoding
    'Dim mDB '() As Variant
    ReDim mInputBuffer(0 To 1)
    Call mBaseStream.Read(mInputBuffer(), 1, 2)
    If (mInputBuffer(0) = &HFF) And (mInputBuffer(1) = &HFE) Then
        'dann ist es wahrscheinlich utf-16, also Integer
        mDecodedCount = mBaseStream.Length \ 2
        ReDim mDecodedBuffer(0 To mDecodedCount) As Integer
    Else
        mDecodedCount = mBaseStream.Length
        ReDim mDecodedBuffer(0 To mDecodedCount) As Byte
    End If
    Call mBaseStream.Read(mDecodedBuffer, 1, mDecodedCount)
    'mDecodedBuffer() = mDB
End Sub
'Public Overridable ReadOnly Property BaseStream() As System.IO.Stream
Public Property Get BaseStream() As Stream
  Set BaseStream = mBaseStream
End Property

'Public Overridable Sub Close()
Private Sub TextReader_cClose()
  Call cClose
End Sub
'Public Overrides Sub Close()
Public Sub cClose()
  mBaseStream.cClose
End Sub

'Public Overridable ReadOnly Property CurrentEncoding() As System.Text.Encoding
'Public Property Get CurrentEncoding() 'As System_Text.Encoding
'
'End Property
'
''Public Sub DiscardBufferedData()
'Public Sub DiscardBufferedData()
'
'End Sub

'Public ReadOnly Shared Dim Null As System.IO.StreamReader
Private Property Get TextReader_nNull() As TextReader
  Set TextReader_nNull = nNull
End Property
Public Property Get nNull() As StreamReader
  'set nnull =
End Property

'Public Overridable Function Peek() As Integer
Private Function TextReader_Peek() As Long
  TextReader_Peek = Peek
End Function
'Public Overrides Function Peek() As Integer
Public Function Peek() As Long
'Gibt, ohne es zu verarbeiten, das nächste verfügbare Zeichen zurück, oder -1, wenn
'keine weiteren Zeichen verfügbar sind oder der Stream keine Suchvorgänge unterstützt.
  If mPos < mDecodedCount Then
    Peek = CLng(mDecodedBuffer(mPos))
  Else
    Peek = -1
  End If
End Function

'Public Overridable Function Read() As Integer
'Public Overridable Function Read(ByVal buffer() As Char, ByVal index As Integer, ByVal count As Integer) As Integer
Private Function TextReader_Read(Optional ByVal buffer As Variant, Optional ByVal index As Long = 0, Optional ByVal Count As Long = 1) As Long
  TextReader_Read = Read(buffer, index, Count)
End Function
'Public Overrides Function Read() As Integer
'Public Overrides Function Read(ByVal buffer() As Char, ByVal index As Integer, ByVal count As Integer) As Integer
Public Function Read(Optional ByVal buffer As Variant, Optional ByVal index As Long = 0, Optional ByVal Count As Long = 1) As Long
'wir lesen momentan nur eins
  If mPos < mDecodedCount Then
    Read = CLng(mDecodedBuffer(mPos))
    mPos = mPos + 1
  Else
    Read = -1
  End If
End Function

'Public Overridable Function ReadBlock(ByVal buffer() As Char, ByVal index As Integer, ByVal count As Integer) As Integer
Private Function TextReader_ReadBlock(ByVal buffer As Variant, ByVal index As Long, ByVal Count As Long) As Long
  TextReader_ReadBlock = ReadBlock(buffer, index, Count)
End Function
Public Function ReadBlock(ByVal buffer As Variant, ByVal index As Long, ByVal Count As Long) As Long
  '
End Function
'Public Overridable Function ReadLine() As String
Private Function TextReader_ReadLine() As String
  TextReader_ReadLine = ReadLine
End Function
'Public Overrides Function ReadLine() As String
Public Function ReadLine() As String
  '
End Function

'Public Overridable Function ReadToEnd() As String
Private Function TextReader_ReadToEnd() As String
  TextReader_ReadToEnd = ReadToEnd
End Function
'Public Overrides Function ReadToEnd() As String
Public Function ReadToEnd() As String
  '
End Function

'Public Shared Function Synchronized(ByVal reader As System.IO.TextReader) As System.IO.TextReader
'Public Function TextReader_Synchronized(ByRef reader As TextReader) As TextReader
'
'End Function

'Protected Overrides Sub Dispose(ByVal disposing As Boolean)
Public Sub Dispose(ByVal disposing As Boolean)
  Call cClose
  Erase mDecodedBuffer
  Erase mInputBuffer
  mPos = 0
End Sub
