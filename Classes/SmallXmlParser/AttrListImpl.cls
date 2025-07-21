VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AttrListImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAttrList
Private attrNames As ArrayList '= New ArrayList
Private attrValues As ArrayList ' = New ArrayList

Private Sub Class_Initialize()
    Call Clear
    'attrNames.Clear
    'attrValues.Clear
End Sub
Private Sub Class_Terminate()
    Set attrNames = Nothing
    Set attrValues = Nothing
End Sub

Private Property Get IAttrList_Length() As Long
    IAttrList_Length = Length
End Property
Public Property Get Length() As Long
    Length = attrNames.Count
End Property

Private Property Get IAttrList_IsEmpty() As Boolean
    IAttrList_IsEmpty = IsEmpty
End Property
Public Property Get IsEmpty() As Boolean
    IsEmpty = (attrNames.Count = 0)
End Property

Private Function IAttrList_GetName(ByVal i As Long) As String
    IAttrList_GetName = GetName(i)
End Function
Public Function GetName(ByVal i As Long) As String
    GetName = attrNames.Item(i)
End Function

Private Function IAttrList_GetValue(ByVal i_name As Variant) As String
    IAttrList_GetValue = GetValue(i_name)
End Function
Public Function GetValue(ByVal i_name As Variant) As String
    GetValue = attrValues.Item(i_name) ', String)
End Function

'Public Function GetValueName(ByVal name As String) As String
'Dim i As Long
'  i = 0
'  While i < attrNames.Count
'    If CType(attrNames(i), String) = name Then
'          Return CType(attrValues(i), String)
'        End If
'    System.Math.Min(System.Threading.Interlocked.Increment(i),i-1)
'  End While
'  'Return Nothing
'End Function

Private Property Get IAttrList_Names() As Variant
    IAttrList_Names = Names
End Property
Public Property Get Names() As Variant 'String()
    Set Names = attrNames.ToArray '(GetType(String)), String())
End Property

Private Function IAttrList_ToStr() As String
    IAttrList_ToStr = ToStr
End Function
Public Function ToStr() As String
    Dim sb As New StringBuilder
    Call ToSB(sb)
    If sb.Length > 0 Then ToStr = sb.ToStr
End Function

Private Sub IAttrList_ToSB(aSB As StringBuilder)
    Call ToSB(aSB)
End Sub
Public Sub ToSB(aSB As StringBuilder)
    Dim i As Long
    For i = 1 To attrNames.Count
        aSB.Append (attrNames.Item(i))
        If i <= attrValues.Count Then
            aSB.Append("=").Append (attrValues.Item(i))
        End If
        If i < attrNames.Count Then Call aSB.Append(" ")
    Next
End Sub

Private Property Get IAttrList_Values() As Variant
    IAttrList_Values = Values
End Property
Public Property Get Values() As Variant 'String()
    Set Values = attrValues.ToArray '(GetType(String)), String())
End Property

Friend Sub Clear()
    Set attrNames = New ArrayList '.Clear
    Set attrValues = New ArrayList '.Clear
End Sub

Friend Sub Add(name As String, Value As String)
    Call attrNames.Add(name)
    Call attrValues.Add(Value)
End Sub

