VERSION 5.00
Begin VB.Form FValidate 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Top             =   0
      Width           =   6975
   End
   Begin VB.Frame FraTest81 
      Caption         =   "8.1 Validate XML Document Against XML Schema"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton BtnTest814 
         Caption         =   "8.1.4 Validating with an Inline XSD Schema.            "
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4935
      End
      Begin VB.CommandButton BtnTest813 
         Caption         =   "8.1.3 Validating with XMLSchemaCache.                 "
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CommandButton BtnTest812 
         Caption         =   "8.1.2 Validating with schemaLocation.                     "
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton BtnTest811 
         Caption         =   "8.1.1 Validating with noNamespaceSchemaLocation."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.CommandButton BtnTest82 
      Caption         =   "8.2 Validate an XML Document or Fragment              "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   5175
   End
End
Attribute VB_Name = "FValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim L As Single: L = Text1.Left
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub BtnTest811_Click()
    '
End Sub

Private Sub BtnTest812_Click()
    '
End Sub

Private Sub BtnTest813_Click()
    '
End Sub

Private Sub BtnTest814_Click()
    '
End Sub

Private Sub BtnTest82_Click()
    ' Output string:
    Dim strout As String
    
    ' Load an XML document into a DOM instance.
    Dim oXMLDoc As DOMDocument60: Set oXMLDoc = MNew.DOMDocument60F(App.path + "\xml\books.xml")
    If oXMLDoc Is Nothing Then Exit Sub
    
    ' Load the schema for the xml document.
    Dim oXSDDoc As DOMDocument60: Set oXSDDoc = MNew.DOMDocument60F(App.path + "\xml\books.xsd")
    If oXSDDoc Is Nothing Then Exit Sub
    
    ' Create a schema cache instance.
    Dim oSCache As New XMLSchemaCache60
    
    ' Add the just-loaded schema definition to the schema collection
    oSCache.Add "urn:books", oXSDDoc
    
    ' Assign the schema to the XML document's schema collection.
    Set oXMLDoc.schemas = oSCache
    
    ' Validate the entire DOM.
    strout = strout & "Validating DOM..." & vbNewLine
    Dim oError As IXMLDOMParseError: Set oError = oXMLDoc.Validate
    If oError.errorCode <> 0 Then
        strout = strout & vbTab & "XMLDoc is not valid because " & vbNewLine & oError.reason & vbNewLine
    Else
        strout = strout & vbTab & "XMLDoc is validated:" & vbNewLine & oXMLDoc.xml & vbNewLine
    End If
    
    Dim oNodes As IXMLDOMNodeList
    ' Validate all "//books" nodes, node by node.
    strout = strout & "Validating all book nodes, '//book\', " & "one by one ..." & vbNewLine
    Set oNodes = oXMLDoc.selectNodes("//book")
    strout = strout & ValidateNodes(oXMLDoc, oNodes)
    
    ' Validate all children of //books nodes, node by node.
    strout = strout & "Validating all children of all book nodes, //book/*, " & "one by one ..." & vbNewLine
    Set oNodes = oXMLDoc.selectNodes("//book/*")
    strout = strout & ValidateNodes(oXMLDoc, oNodes)
    
    Text1.Text = strout
End Sub

Private Function ValidateNodes(oXMLDoc As DOMDocument60, oNodes As IXMLDOMNodeList) As String
    If oXMLDoc Is Nothing Then
        ValidateNodes = "Error in ValidateNodes(): Invalid oXMLDoc"
        Exit Function
    End If
    
    If oNodes Is Nothing Then
        ValidateNodes = "Error in ValidateNodes(): Invalid oNodes"
        Exit Function
    End If
    
    Dim oNode As IXMLDOMNode
    Dim oError As IXMLDOMParseError
    Dim strout As String
    
    Dim i As Long
    For i = 0 To oNodes.length - 1
        Set oNode = oNodes.nextNode
        If Not (oNode Is Nothing) Then
           Set oError = oXMLDoc.validateNode(oNode)
           If oError.errorCode = 0 Then
               strout = strout & vbTab & "<" & oNode.nodeName & "> (" & CStr(i) & ") is a valid node " & vbNewLine
           Else
               strout = strout & vbTab & "<" & oNode.nodeName & "> (" & CStr(i) & ") " & "is not valid because" & vbNewLine & oError.reason & vbNewLine
           End If
        End If
    Next
    ValidateNodes = strout
End Function

'Output 8.2
'Validating DOM...
'        XMLDoc is not valid because
'Element 'review' is unexpected according to content model of parent element 'book'.
'Expecting: pub_date.
'
'Validating all book nodes, '//book', one by one ...
'        <book (0) is a valid node
'        <book> (1) is not valid because
'Element 'review' is unexpected according to content model of parent element 'book'.
'Expecting: pub_date.
'
'Validating all children of all book nodes, //book/*, one by one...
'        <author> (0) is a valid node
'        <title> (1) is a valid node
'        <genre> (2) is a valid node
'        <price> (3) is a valid node
'        <pub_date> (4) is a valid node
'        <review> (5) is a valid node
'        <author> (6) is a valid node
'        <title> (7) is a valid node
'        <genre> (8) is a valid node
'        <price> (9) is a valid node
'        <review> (10) is not valid beacause
'Element 'review' is unexpected according to content model of parent element 'book'.
'Expecting: pub_date.
