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
'8. Validate XML
'---------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms756009(v=vs.85)
'Demonstrates how to validate XML documents against an XML schema as well as how to validate document node fragments.
'
'8.1 Validate an XML Document Against an XML Schema
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757071(v=vs.85)
'Demonstrates how to validate an XML document and/or fragment against an XML schema using Visual Basic.
'
'8.1.1
'Example 1: Validating with noNamespaceSchemaLocation
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757051(v=vs.85)
'You can use the xsi:noNamespaceSchemaLocationattribute to reference the XSD schema file from within the XML document.
'
'8.1.2
'Example 2: Validating with schemaLocation
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757072(v=vs.85)
'
'8.1.3
'Example 3: Validating with XMLSchemaCache
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762616(v=vs.85)
'
'8.1.4
'Example 4: Validating with an Inline XSD Schema
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms754572(v=vs.85)
'
'8.2 Validate an XML Document or Fragment
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762693(v=vs.85)
'Demonstrates how to validate an XML document and/or fragment against an XML schema using Visual Basic.

Private Sub Form_Resize()
    Dim L As Single: L = Text1.Left
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

' ############################## '        ' ############################## '
Private Sub BtnTest811_Click()
    Dim sPath As String: sPath = App.path & "\Xml\"
    Dim sOutput As String
    sOutput = ValidateFile_nn(sPath & "nn-valid.xml")
    sOutput = sOutput & ValidateFile_nn(sPath & "nn-notValid.xml")
    Text1.Text = sOutput
End Sub
Function ValidateFile_nn(sXmlPathFileName As String) As String
    Dim s As String
    ' Create an XML DOMDocument object and set first-level DOM properties.
    Dim x As DOMDocument60: Set x = MNew.DOMDocument60(False, , True)
    
    ' Load and validate the specified file into the DOM.
    x.Load sXmlPathFileName
    ' Return validation results in message to the user.
    If x.parseError.errorCode <> 0 Then
        s = "Validation failed on: " & sXmlPathFileName & vbCrLf & _
            "===================== " & vbCrLf & _
            "Reason: " & x.parseError.Reason & vbCrLf & _
            "Source: " & x.parseError.srcText & vbCrLf & _
            "Line  : " & x.parseError.Line & vbCrLf
    Else
        s = "Validation succeeded for: " & sXmlPathFileName & vbCrLf & _
            "========================= " & vbCrLf & _
            x.xml & vbCrLf
    End If
    ValidateFile_nn = s
End Function
'Output for 8.1.1:
'Validation succeeded for nn-valid.xml
'=====================================
'<?xml version="1.0"?>
'<catalog>
'    <book xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'
'xsi:          noNamespaceSchemaLocation = "nn.xsd"
'          id="bk101">
'        <author>Gambardella, Matthew</author>
'        <title>XML Developer's Guide</title>
'        <genre>Computer</genre>
'        <price>44.95</price>
'        <publish_date>2000-10-01</publish_date>
'        <description>An in-depth look at creating applications
'         with XML.</description>
'    </book>
'</catalog>
'
'Validation failed on nn-notValid.xml
'====================================
'Reason: Element 'cost' is unexpected according to content model of parent element 'book'.
'
'Expecting: price.
'
'Source:       <cost>44.95</cost>
'Line: 8
'''Explanation: This is because the correct and valid name for the element in use at this location in the XML documents is <price/>, not <cost/>.

' ############################## '        ' ############################## '
Private Sub BtnTest812_Click()
    Dim sPath As String: sPath = App.path & "\Xml\"
    Dim sOutput As String
    sOutput = ValidateFile_sl(sPath & "sl-valid.xml")
    sOutput = sOutput & ValidateFile_sl(sPath & "sl-notValid.xml")
    Text1.Text = sOutput
End Sub
Function ValidateFile_sl(sXmlPathFileName As String) As String
    Dim s As String
    ' Create an XML DOMDocument object and set first-level DOM properties.
    Dim x As DOMDocument60: Set x = MNew.DOMDocument60(False, , True)
    
    ' Configure DOM properties for namespace selection.
    x.setProperty "SelectionLanguage", "XPath"
    Dim ns As String: ns = "xmlns:x='urn:book'"
    x.setProperty "SelectionNamespaces", ns
    ' Load and validate the specified file into the DOM.
    x.Load sXmlPathFileName
    ' Return validation results in message to the user.
    If x.parseError.errorCode <> 0 Then
        s = "Validation failed on: " & sXmlPathFileName & vbCrLf & _
            "===================== " & vbCrLf & _
            "Reason: " & x.parseError.Reason & vbCrLf & _
            "Source: " & x.parseError.srcText & vbCrLf & _
            "Line  : " & x.parseError.Line & vbCrLf
    Else
        s = "Validation succeeded for: " & sXmlPathFileName & vbCrLf & _
             "======================== " & vbCrLf & _
             x.xml & vbCrLf
    End If
    ValidateFile_sl = s
End Function
'Output for 8.1.2:
'Validation succeeded for sl-valid.xml
'=====================================
'<?xml version="1.0"?>
'<catalog xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'
'         xsi:schemaLocation='urn:book sl.xsd'>
'   <x:book xmlns:x='urn:book' id="bk101">
'      <x:author>Gambardella, Matthew</x:author>
'      <x:title>XML Developer's Guide</x:title>
'      <x:genre>Computer</x:genre>
'      <x:price>44.95</x:price>
'      <x:publish_date>2000-10-01</x:publish_date>
'      <x:description>An in-depth look at creating applications with
'      XML.</x:description>
'   </x:book>
'</catalog>
'
'Validation failed on sl-notValid.xml
'====================================
'Reason: Element '{urn:book}cost' is unexpected according to content model of parent element '{urn:book}book'.
'Expecting: {urn:book}price.
'
'Source:       <x:cost>44.95</x:cost>
'Line: 10

' ############################## '        ' ############################## '
Private Sub BtnTest813_Click()
    Dim sPath As String: sPath = App.path & "\Xml\"
    Dim s As String: s = ""
    s = s & ValidateFile_sc(sPath & "sc-valid.xml", "urn:books", sPath & "sc.xsd")
    s = s & ValidateFile_sc(sPath & "sc-notValid.xml", "urn:books", sPath & "sc.xsd")
    Text1.Text = s
End Sub
Function ValidateFile_sc(sXmlPathFileName As String, sUrn As String, sXsdPathFileName As String) As String
    Dim s As String
    ' Create a schema cache and add books.xsd to it.
    Dim xs As New MSXML2.XMLSchemaCache60: xs.Add sUrn, sXsdPathFileName
    
    ' Create an XML DOMDocument object.
    Dim xd As New MSXML2.DOMDocument60
    ' Assign the schema cache to the DOM document.
    ' schemas collection.
    Set xd.schemas = xs
    ' Load books.xml as the DOM document.
    xd.async = False
    xd.Load sXmlPathFileName
    
    ' Return validation results in message to the user.
    If xd.parseError.errorCode <> 0 Then
        s = "Validation failed on: " & sXmlPathFileName & vbCrLf & _
            "===================== " & vbCrLf & _
            "Reason: " & xd.parseError.Reason & vbCrLf & _
            "Source: " & xd.parseError.srcText & vbCrLf & _
            "Line  : " & xd.parseError.Line & vbCrLf
    Else
         s = "Validation succeeded for: " & sXmlPathFileName & vbCrLf & _
             "========================= " & vbCrLf & _
             xd.xml & vbCrLf
    End If
    ValidateFile_sc = s
End Function
'Output:
'Validation succeeded for sc-valid.xml
'======================
'<?xml version="1.0"?>
'<x:catalog xmlns:x="urn:books">
'        <book id="bk101">
'                <author>Gambardella, Matthew</author>
'                <title>XML Developer's Guide</title>
'                <genre>Computer</genre>
'                <price>44.95</price>
'                <publish_date>2000-10-01</publish_date>
'                <description>An in-depth look at creating applications
'      with XML.</description>
'        </book>
'</x:catalog>
'
'Validation failed on sc-notValid.xml
'=====================
'Reason: Element 'cost' is unexpected according to content model of parent elemen
'T 'book'.
'Expecting: price.
'
'Source:       <cost>44.95</cost>
'Line: 7

' ############################## '        ' ############################## '
Private Sub BtnTest814_Click()
    Dim sPath As String: sPath = App.path & "\Xml\"
    Dim sOutput As String
    sOutput = ValidateFile_il(sPath & "il-valid.xml")
    sOutput = sOutput & ValidateFile_il(sPath & "il-notValid.xml")
    Text1.Text = sOutput
End Sub
Function ValidateFile_il(sXmlPathFileName As String) As String
    Dim s As String
    ' Create an XML DOMDocument object and set first-level DOM properties.
    Dim x As DOMDocument60: Set x = MNew.DOMDocument60(False, , True)
    ' Load and validate the specified file into the DOM.
    x.setProperty "UseInlineSchema", True
    x.Load sXmlPathFileName
    
    ' Return validation results in message to the user.
    If x.parseError.errorCode <> 0 Then
       s = "Validation failed on " & sXmlPathFileName & vbCrLf & _
           "=====================" & vbCrLf & _
           "Reason: " & x.parseError.Reason & vbCrLf & _
           "Source: " & x.parseError.srcText & vbCrLf & _
           "Line  : " & x.parseError.Line & vbCrLf
    Else
       s = "Validation succeeded for " & sXmlPathFileName & vbCrLf & _
           "======================" & vbCrLf & _
           x.xml & vbCrLf
    End If
    ValidateFile_il = s
End Function

' ############################## '        ' ############################## '
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
        strout = strout & vbTab & "XMLDoc is not valid because " & vbNewLine & oError.Reason & vbNewLine
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
               strout = strout & vbTab & "<" & oNode.nodeName & "> (" & CStr(i) & ") " & "is not valid because" & vbNewLine & oError.Reason & vbNewLine
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
