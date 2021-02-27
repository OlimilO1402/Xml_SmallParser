VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SimpleArrayList Collection
Private mCol As Collection

Private Sub Class_Initialize()
  Call Clear
End Sub
'hier könnte es auch Get Item as String heißen
Public Property Get Item(Index_Name As Variant) As Variant
Attribute Item.VB_UserMemId = 0
  If TypeOf mCol.Item(Index_Name) Is Object  Then
    Set Item = mCol.Item(Index_Name)
  Else
    Item = mCol.Item(Index_Name)
  End If
End Property

Public Sub Clear()
  Set mCol = New Collection
End Sub

'Public Function GetEnumerator() As IUnknown
'  GetEnumerator = mCol.[_NewEnum]
'End Function

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Function ToArray() As Collection 'Variant
'ReDim AArr(mCol.Count)
'Dim i As Long
  Set ToArray = mCol 'Array(mCol)
End Function

Public Function ToStr() As String
'Dim sb As New StringBuilder
'  Call ToSB(sb)
'  ToStr = sb.ToStr
End Function

'Public Sub ToSB(aSB As StringBuilder)
'Dim i As Long
'  For i = 1 To mCol.Count
'    Call aSB.Append(mCol.Item(i))
'    If (i < mCol.Count) And (mCol.Count > 1) Then Call aSB.Append(", ")
'  Next
'  ToStr = aSB.ToStr
'End Sub

Public Sub Add(StrValue As String)
  mCol.Add StrValue
End Sub
