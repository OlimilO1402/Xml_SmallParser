VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAttrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Length() As Long
End Property

Public Property Get IsEmpty() As Boolean
End Property

Public Function GetName(ByVal i As Long) As String
End Function

Public Function GetValue(ByVal i_name As Variant) As String
End Function

Public Property Get Names() As Variant 'String()
End Property

Public Property Get Values() As Variant 'String()
End Property

Public Function ToStr() As String
End Function

Public Sub ToSB(aSB As StringBuilder)
End Sub
