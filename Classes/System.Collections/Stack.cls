VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class Stack
'          Inherits System.Object
'     Member von: System.Collections
Option Explicit
Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

'Public Sub New(ByVal initialCapacity As Integer)
'Public Sub New(ByVal col As System.Collections.ICollection)
Friend Sub New_(ByVal initialCapacity As Long)
  'nö
End Sub

Public Sub Clear()
  Set mCol = New Collection
End Sub

'Public Overridable Function Clone() As Object
Public Function Clone() As Object
  Set Clone = Me
End Function

'Public Overridable Function Contains(ByVal obj As Object) As Boolean
Public Function Contains(ByVal obj As Object) As Boolean
Dim aObj 'As Object
  'Set Contains = New cBoolean
  For Each aObj In mCol
    If aObj Is obj Then
      Contains = True
      Exit Function
    End If
  Next
End Function

'Public Overridable Sub CopyTo(ByVal array As System.Array, ByVal index As Integer)
'Public Sub CopyTo(ByVal sarray As SystemArray, ByVal index As Long)
'
'End Sub

'Public Overridable ReadOnly Property Count() As Integer
Public Property Get Count() As Long
  Count = mCol.Count
End Property

'Public Overridable Function GetEnumerator() As System.Collections.IEnumerator
'Public Function GetEnumerator() As IUnknown 'IEnumerator
'  Set GetEnumerator = mCol.[_NewEnum]
'End Function

'Public Overridable ReadOnly Property IsSynchronized() As Boolean
Public Property Get IsSynchronized() As Boolean
  'Set IsSynchronized = New cBoolean
End Property

'Public Overridable Function Peek() As Object
Public Function Peek() As Variant 'Object
'Gibt das oberste Objekt von System.Collections.Stack zurück, ohne es zu entfernen.
  If TypeOf mCol.Item(mCol.Count) Is Object  Then
    Set Peek = mCol.Item(mCol.Count)
  Else
    Peek = mCol.Item(mCol.Count)
  End If
End Function

'Public Overridable Function Pop() As Object
Public Function Pop() As Variant 'Object
'Entfernt das oberste Objekt aus System.Collections.Stack und gibt es zurück.
  If TypeOf mCol.Item(mCol.Count) Is Object  Then
    Set Pop = mCol.Item(mCol.Count)
  Else
    Pop = mCol.Item(mCol.Count)
  End If
  Call mCol.Remove(mCol.Count)
End Function

'Public Overridable Sub Push(ByVal obj As Object)
Public Sub Push(ByVal obj As Variant)
'Fügt ein Objekt am Anfang von System.Collections.Stack ein.
  If mCol.Count = 0 Then
    mCol.Add obj
  Else
    mCol.Add obj ', , 1
  End If
End Sub

'Public Shared Function Synchronized(ByVal stack As System.Collections.Stack) As System.Collections.Stack
Public Static Function Synchronized(ByVal aStack As Stack) As Stack
  '
End Function

'Public Overridable ReadOnly Property SyncRoot() As Object
Public Property Get SyncRoot() As Object
  '
End Property

'Public Overridable Function ToArray() As Object()
Public Function ToArray() As Object '()
  'Set ToArray = New SystemArray
  
End Function

Public Function ToStr() As String
Dim i As Long
  ToStr = "{"
  For i = 1 To mCol.Count
    ToStr = ToStr & mCol.Item(i)
    If i < mCol.Count Then ToStr = ToStr & ", "
  Next
  ToStr = ToStr & "}"
End Function
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
