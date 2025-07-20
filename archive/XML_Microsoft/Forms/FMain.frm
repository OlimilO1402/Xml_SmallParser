VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "MS XML-Tutorial"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
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
   ScaleHeight     =   4815
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTest8 
      Caption         =   "8. Validate XML >>"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest7 
      Caption         =   "7. Make XML Requests Over HTTP >>"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest6 
      Caption         =   "6: Query XML DOM Nodes"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest5 
      Caption         =   "5: Create an XML DOM Object Dynamically"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest4 
      Caption         =   "4. Perform XSL Transformations"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest3 
      Caption         =   "3. Serialize an XML DOM Object to a File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton BtnTest2 
      Caption         =   "2. Load an XML File into a DOM Object"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
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
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   0
      Width           =   6975
   End
   Begin VB.CommandButton BtnTest1 
      Caption         =   "1. Instantiate an XML DOM Object"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This VB-Project is dedicated to the XML-Tutorial by Microsoft:
'
'Program with DOM in Visual Basic
'================================
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms766564(v=vs.85)
'
'Set Up My Visual Basic Project:
'-------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms763689(v=vs.85)
'Discusses the requirements for using MSXML, and describes how to install the MSXML components.
'Microsoft XML, v6.0: C:\Windows\SysWOW64\msxml6.dll
'
'1. Instantiate an XML DOM Object
'--------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms759156(v=vs.85)
'Demonstrates two ways to instantiate an XML DOM object in Visual Basic.
'see "Sub Test1"
'
'2. Load an XML File into a DOM Object
'-------------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms765424(v=vs.85)
'Demonstrates how to create an XML DOM instance and load its content from an external XML data file.
'see "Sub Test2"
'
'3. Serialize an XML DOM Object to a File
'----------------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms767687(v=vs.85)
'Demonstrates how to serialize an XML DOM object in a text file.
'see "Sub Test3"
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762318(v=vs.85)
'
'4. Perform XSL Transformations
'------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms765520(v=vs.85)
'Demonstrates how to perform XSL Transformations (XSLT).
'see "Sub Test4"
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762211(v=vs.85)
'
'5. Create an XML DOM Object Dynamically
'---------------------------------------
'Demonstrates how to create an XML DOM object programmatically, including processing
'instructions, comments, elements, attributes, CDATA sections and text nodes.
'see "Sub Test5"
'
'6. Query XML DOM Nodes
'----------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms765471(v=vs.85)
'Demonstrates how to query a DOM node or node-set using XPath expressions.
'
'7. Make XML Requests Over HTTP
'------------------------------
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757026(v=vs.85)
'Demonstrates how to perform client requests of XML data from a web server.
'
Private Sub Form_Resize()
    Dim L As Single: L = Text1.Left
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

'
'Source: InstantiateDOM.frm
Private Sub BtnTest1_Click()
    ' Instantiate an object at run time.
    Dim dom1
    Set dom1 = New DOMDocument60
    dom1.async = False
    dom1.resolveExternals = False
    dom1.loadXML "<a>A</a>"
    
    ' Instantiate an object at compile time.
    Dim dom2 As New DOMDocument60
    dom2.async = False
    dom2.resolveExternals = False
    dom2.loadXML "<b>B</b>"
    
    ' Display the content of both objects.
    Text1.Text = "Test1:" & vbCrLf & "dom1: " & dom1.xml & vbNewLine & "dom2: " & dom2.xml
End Sub

'Source: LoadXMLFile.frm
Private Sub BtnTest2_Click()
    Dim doc As New DOMDocument60
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    
    Dim path As String
    path = App.path & "\Xml\test2.xml"
    doc.Load path
    Text1.Text = "Test2:" & vbCrLf & "doc: " & vbNewLine & doc.xml
End Sub

'Source: SaveDOMToFile.frm
Private Sub BtnTest3_Click()
    Dim doc As New DOMDocument60
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    
    Dim sxml As String: sxml = _
        "<?xml version='1.0'?>" & vbNewLine & _
        "<doc title='test'>" & vbNewLine & _
        "   <page num='1'>" & vbNewLine & _
        "      <para title='Saved at last'>" & vbNewLine & _
        "          This XML data is finally saved." & vbNewLine & _
        "      </para>" & vbNewLine & _
        "   </page>" & vbNewLine & _
        "   <page num='2'>" & vbNewLine & _
        "      <para>" & vbNewLine & _
        "          This page is intentionally left blank." & vbNewLine & _
        "      </para>" & vbNewLine & _
        "   </page>" & vbNewLine & _
        "</doc>"
        
    doc.loadXML sxml
    Dim path As String
    path = App.path & "\Xml\saved.xml"
    doc.save path
    Text1.Text = "Test3:" & vbCrLf & doc.xml
End Sub

'Source: XSLT.frm
Private Sub BtnTest4_Click()
    Dim s As String
    Dim doc As DOMDocument60: Set doc = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False) ', doResolveExternals:=False, doPreserveWhitespace:=False
    doc.Load App.path & "\Xml\test4.xml"
    
    Dim xsl As DOMDocument60: Set xsl = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False) ', doResolveExternals:=False, doPreserveWhitespace:=False
    xsl.Load App.path & "\Xml\test4.xsl"
    
    Dim str As String: str = doc.transformNode(xsl)
    s = "Test4.1:" & vbCrLf & "doc.transformNode: " & vbNewLine & str
    
    Dim out As DOMDocument60: Set out = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False) ', doResolveExternals:=False, doPreserveWhitespace:=False
    doc.transformNodeToObject xsl, out
    Text1.Text = s & vbNewLine & "Test4.2:" & vbCrLf & "doc.transformNodeToObject:" & vbNewLine & out.xml
End Sub

'Source: dynamDOM.frm
Private Sub BtnTest5_Click()
    Dim dom  As MSXML2.DOMDocument60
    Dim Node As MSXML2.IXMLDOMNode
    Dim attr As MSXML2.IXMLDOMAttribute

Try: On Error GoTo Catch

    Set dom = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False, doPreserveWhitespace:=True)  ', doResolveExternals:=False, doPreserveWhitespace:=False
    
    ' Create a processing instruction targeted for xml.
    Set Node = dom.createProcessingInstruction("xml", "version='1.0'")
    dom.appendChild Node
    
    ' Create a processing instruction targeted for xml-stylesheet.
    Set Node = dom.createProcessingInstruction("xml-stylesheet", "type='text/xml' href='test.xsl'")
    dom.appendChild Node
    
    ' Create a comment for the document.
    Set Node = dom.createComment("sample xml file created using XML DOM object.")
    dom.appendChild Node
    
    ' Create the root element.
    Dim root As IXMLDOMElement
    Set root = dom.createElement("root")
    
    ' Create a "created" attribute for the root element and
    ' assign the "using dom" character data as the attribute value.
    Set attr = dom.createAttribute("created")
    attr.Value = "using dom"
    root.setAttributeNode attr
    
    ' Add the root element to the DOM instance.
    dom.appendChild root
    ' Insert a newline + tab.
    root.appendChild dom.createTextNode(vbNewLine & vbTab)
    ' Create and add more nodes to the root element just created.
    ' Create a text element.
    Set Node = dom.createElement("node1")
    Node.Text = "some character data"
    ' Add text node to the root element.
    root.appendChild Node
      ' Add a newline plus tab.
    root.appendChild dom.createTextNode(vbNewLine & vbTab)
    
    ' Create an element to hold a CDATA section.
    Set Node = dom.createElement("node2")
    Dim cd As IXMLDOMCDATASection
    Set cd = dom.createCDATASection("<some mark-up text>")
    Node.appendChild cd
    dom.documentElement.appendChild Node
      ' Add a newline plus tab.
    root.appendChild dom.createTextNode(vbNewLine & vbTab)
    
    ' Create an element to hold three empty subelements.
    Set Node = dom.createElement("node3")
    
    ' Create a document fragment to be added to node3.
    Dim frag As IXMLDOMDocumentFragment
    Set frag = dom.createDocumentFragment
        ' Add a newline + tab + tab.
    frag.appendChild dom.createTextNode(vbNewLine & vbTab & vbTab)
    frag.appendChild dom.createElement("subNode1")
       ' Add a newline + tab + tab.
    frag.appendChild dom.createTextNode(vbNewLine & vbTab & vbTab)
    frag.appendChild dom.createElement("subNode2")
       ' Add a newline + tab + tab.
    frag.appendChild dom.createTextNode(vbNewLine & vbTab & vbTab)
    frag.appendChild dom.createElement("subNode3")
       ' Add a newline + tab.
    frag.appendChild dom.createTextNode(vbNewLine & vbTab)
    Node.appendChild frag
    root.appendChild Node
    
        ' Add a newline.
    root.appendChild dom.createTextNode(vbNewLine)
    
    ' Save the XML document to a file.
    dom.save App.path & "\xml\dynamDom.xml"
    Text1.Text = "Test5:" & vbCrLf & dom.xml
    
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub

'Source: queryNodes.frm
Private Sub BtnTest6_Click()
    ' Output string:
    Dim strout As String
    
    ' Load an xml document into a DOM instance.
    Dim oXMLDom As DOMDocument60: Set oXMLDom = MNew.DOMDocument60(False, False, , True)
    If oXMLDom.Load(App.path & "\Xml\stocks.xml") = False Then
        MsgBox "Failed to load xml data from file."
        Exit Sub
    End If
    
    ' Query a single node.
    Dim oNode As IXMLDOMNode: Set oNode = oXMLDom.selectSingleNode("//stock[1]/*")
    If oNode Is Nothing Then GoTo MoreNodes
    
    strout = "Result from selectSingleNode" & vbNewLine & _
             "Node, <" & oNode.nodeName & ">: " & vbNewLine & vbTab & _
             oNode.xml & vbNewLine & vbNewLine
    
MoreNodes:
    ' Query a node-set.
    Dim oNodes As IXMLDOMNodeList
    Set oNodes = oXMLDom.selectNodes("//stock[1]/*")
    
    strout = strout & "Results from selectNodes:" & vbNewLine
    Dim sName As String
    Dim sData As String
    Dim i As Long
    For i = 0 To oNodes.length - 1
        Set oNode = oNodes.nextNode
        If Not (oNode Is Nothing) Then
            sName = oNode.nodeName
            sData = oNode.xml
            strout = strout _
               & "Node (" & CStr(i) & "), <" & sName & ">:" _
               & vbNewLine & vbTab & sData & vbNewLine
        End If
    Next
    Text1.Text = "Test6:" & vbCrLf & strout
End Sub

Private Sub BtnTest7_Click()
    FXMLOverHTTP.Show
End Sub

Private Sub BtnTest8_Click()
    FValidate.Show
End Sub

