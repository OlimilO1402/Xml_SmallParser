VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IContentHandler
Private mSB As StringBuilder

Private Sub Class_Initialize()
    Set mSB = New StringBuilder
End Sub
Private Sub Class_Terminate()
    Set mSB = Nothing
End Sub

Public Sub OnStartParsing(parser As SmallXmlParser)
    'MsgB "Parser starts"
    'Call mSB.Append("Parser starts").Append(vbCrLf)
End Sub

Public Sub OnEndParsing(parser As SmallXmlParser)
    'MsgB "Parser ends"
    'Call mSB.Append("Parser ends").Append(vbCrLf)
End Sub

Public Sub OnStartElement(ByVal name As String, attrs As IAttrList)
    'MsgB "name=" & name & vbCrLf & attrs.ToStr
    Call mSB.Append("<").Append(name).Append(" ")
    Call attrs.ToSB(mSB)
    Call mSB.Append(">")
End Sub

Public Sub OnEndElement(ByVal name As String)
    'MsgB "Name=" & name
    Call mSB.Append("</").Append(name).Append(">")
End Sub

Public Sub OnChars(ByVal s As String)
    'MsgB "Chars='" & s & "'"
    Call mSB.Append(s)
End Sub

Public Sub OnIgnorableWhitespace(ByVal s As String)
    'MsgB s
    Call mSB.Append(s)
End Sub

Public Sub OnProcessingInstruction(ByVal name As String, ByVal text As String)
    'Call MsgB("name=" & name & "; text=" & text)
    Call mSB.Append("<").Append("?").Append(name).Append(" ").Append(text).Append(">")
End Sub

Public Function ToStr() As String
    ToStr = mSB.ToStr
End Function
Private Sub MsgB(str As String)
    Dim mr As VbMsgBoxResult
    mr = MsgBox(str, vbOKCancel)
    If mr = vbOK Then
        Exit Sub
    Else
        End 'Exit Sub
    End If
End Sub

'########################' v Implements IContentHandler v '########################'
Private Sub IContentHandler_OnStartParsing(parser As SmallXmlParser)
    Call OnStartParsing(parser)
End Sub
Private Sub IContentHandler_OnEndParsing(parser As SmallXmlParser)
    Call OnEndParsing(parser)
End Sub
Private Sub IContentHandler_OnStartElement(ByVal name As String, attrs As IAttrList)
    Call OnStartElement(name, attrs)
End Sub
Private Sub IContentHandler_OnEndElement(ByVal name As String)
    Call OnEndElement(name)
End Sub
Private Sub IContentHandler_OnChars(ByVal s As String)
    Call OnChars(s)
End Sub
Private Sub IContentHandler_OnIgnorableWhitespace(ByVal s As String)
    Call OnIgnorableWhitespace(s)
End Sub
Private Sub IContentHandler_OnProcessingInstruction(ByVal name As String, ByVal text As String)
    Call OnProcessingInstruction(name, text)
End Sub
