VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub OnStartParsing(parser As SmallXmlParser)
End Sub

Public Sub OnEndParsing(parser As SmallXmlParser)
End Sub

Public Sub OnStartElement(ByVal name As String, attrs As IAttrList)
End Sub

Public Sub OnEndElement(ByVal name As String)
End Sub

Public Sub OnChars(ByVal s As String)
End Sub

Public Sub OnIgnorableWhitespace(ByVal s As String)
End Sub

Public Sub OnProcessingInstruction(ByVal name As String, ByVal text As String)
End Sub
