VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmallXmlParserException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inherits SystemException
Private mMess As String
Private mLine As Long
Private mColumn As Long
'Private MyBase As SystemException

Private Sub Class_Initialize()
  '
End Sub
Friend Sub New_(ByVal aMsg As String, ByVal ALine As Long, ByVal AColumn As Long)
  'MyBase.New(String.Format("{0}. At ({1},{2})", msg, line, column))
  mLine = ALine
  mColumn = AColumn
  mMess = aMsg & ". At (" & CStr(mLine) & "," & CStr(mColumn) & ")"
End Sub

Public Property Get Line() As Long
  Line = mLine
End Property

Public Property Get Column() As Long
  Column = mColumn
End Property

'Public Sub Throw()
''Dim e As New ErrObject
'  'e.Description = mMess
'   'Call Err.Raise(9, , mMess)
'End Sub
