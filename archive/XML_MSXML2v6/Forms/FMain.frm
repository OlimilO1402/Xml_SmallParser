VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
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
   ScaleHeight     =   8085
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7605
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton BtnReadFile 
      Caption         =   "Read xml file"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1455
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
      Height          =   7605
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton BtnOpenFile 
      Caption         =   "Open xml file"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://activevb.de/tutorials/tut_msxml/msxml.html
'Microsoft XML, v6.0
'C:\Windows\SysWOW64\msxml6.dll
Private m_xPFN As PathFileName
Private m_xdoc As MSXML2.DOMDocument60
Private m_SB   As StringBuilder
Private m_IndStack As Integer

Private Sub Form_Load()
    BtnReadFile.Enabled = False
End Sub

Private Sub Form_Resize()
    Dim l As Single
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth / 2
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move l, t, W, H
    l = W
    If W > 0 And H > 0 Then Text2.Move l, t, W, H
End Sub

Private Sub BtnOpenFile_Click()
    Set m_xPFN = MNew.PathFileName(GetXmlFile)
    If Not m_xPFN.Exists Then
        Set m_xdoc = Nothing
        BtnReadFile.Enabled = False
        Exit Sub
    End If
    Set m_xdoc = New MSXML2.DOMDocument60
    m_xdoc.validateOnParse = True
    m_xdoc.setProperty "SelectionLanguage", "XPath"
    If Not m_xdoc.Load(m_xPFN.Value) Then Exit Sub
    m_xPFN.Encoding = ETextEncoding.Text_UTF8Encoding
    Dim s As String: s = m_xPFN.ReadAllText
    s = Replace(s, vbLf, vbCrLf)
    Text1.Text = s
    BtnReadFile.Enabled = True
End Sub

Private Function GetXmlFile() As String
    With New OpenFileDialog
        .InitialDirectory = App.Path '"C:\Windows\System32\"
        .Filter = "Komoot-gps-Dateien [*.gpx]|*.gpx|" & _
                  "Xml-Dateien [*.xml]|*.xml|"
        If .ShowDialog = VbMsgBoxResult.vbCancel Then Exit Function
        GetXmlFile = .FileName
    End With
End Function

Private Sub BtnReadFile_Click()
    If m_xdoc Is Nothing Then Exit Sub
    Set m_SB = New StringBuilder
    Dim xnods As MSXML2.IXMLDOMNodeList: Set xnods = m_xdoc.childNodes
    If xnods.Length Then
        PrintNodes xnods
        Text2.Text = m_SB.ToStr
    End If
End Sub

Private Sub PrintNodes(nods As IXMLDOMNodeList)
    Dim xnod As MSXML2.IXMLDOMNode
    Dim i As Long
    For i = 0 To nods.Length - 1
        Set xnod = nods.Item(i)
        PrintNode xnod
    Next
End Sub

Private Function NodeType_ToStr(E As tagDOMNodeType) As String
    Dim s As String
    Select Case E
    Case tagDOMNodeType.NODE_INVALID:                s = "NODE_INVALID"                ' 0
    Case tagDOMNodeType.NODE_ELEMENT:                s = "NODE_ELEMENT"                ' 1
    Case tagDOMNodeType.NODE_ATTRIBUTE:              s = "NODE_ATTRIBUTE"              ' 2
    Case tagDOMNodeType.NODE_TEXT:                   s = "NODE_TEXT"                   ' 3
    Case tagDOMNodeType.NODE_CDATA_SECTION:          s = "NODE_CDATA_SECTION"          ' 4
    Case tagDOMNodeType.NODE_ENTITY_REFERENCE:       s = "NODE_ENTITY_REFERENCE"       ' 5
    Case tagDOMNodeType.NODE_ENTITY:                 s = "NODE_ENTITY"                 ' 6
    Case tagDOMNodeType.NODE_PROCESSING_INSTRUCTION: s = "NODE_PROCESSING_INSTRUCTION" ' 7
    Case tagDOMNodeType.NODE_COMMENT:                s = "NODE_COMMENT"                ' 8
    Case tagDOMNodeType.NODE_DOCUMENT:               s = "NODE_DOCUMENT"               ' 9
    Case tagDOMNodeType.NODE_DOCUMENT_TYPE:          s = "NODE_DOCUMENT_TYPE"          '10
    Case tagDOMNodeType.NODE_DOCUMENT_FRAGMENT:      s = "NODE_DOCUMENT_FRAGMENT"      '11
    Case tagDOMNodeType.NODE_NOTATION:               s = "NODE_NOTATION"               '12
    End Select
    NodeType_ToStr = s
End Function

Private Sub PrintNode(nod As IXMLDOMNode)
    Dim nt As tagDOMNodeType: nt = nod.nodeType
    'Debug.Print NodeType_ToStr(nt) & ": " & nod.nodeName
    Select Case nt
    Case tagDOMNodeType.NODE_INVALID:                PrintNodeInvalid nod
    Case tagDOMNodeType.NODE_ELEMENT:                PrintNodeElement nod
    Case tagDOMNodeType.NODE_ATTRIBUTE:              PrintNodeAttributes nod
    Case tagDOMNodeType.NODE_TEXT:                   'PrintNodeText nod
    Case tagDOMNodeType.NODE_CDATA_SECTION:          PrintNodeCData nod
    Case tagDOMNodeType.NODE_ENTITY_REFERENCE:       PrintNodEntRef nod
    Case tagDOMNodeType.NODE_ENTITY:                 PrintNodeEntity nod
    Case tagDOMNodeType.NODE_PROCESSING_INSTRUCTION: PrintNodeProcInstr nod
    Case tagDOMNodeType.NODE_COMMENT:                PrintNodeComment nod
    Case tagDOMNodeType.NODE_DOCUMENT:               PrintNodeDocument nod
    Case tagDOMNodeType.NODE_DOCUMENT_TYPE:          PrintNodeDocType nod
    Case tagDOMNodeType.NODE_DOCUMENT_FRAGMENT:      PrintNodeDocFrgmt nod
    Case tagDOMNodeType.NODE_NOTATION:               PrintNodeNotation nod
    Case Else: 'do nothing, what else to do?
    End Select
'    m_SB.Append "<"
'    Dim isProcInstr As Boolean: isProcInstr = nod.nodeType = tagDOMNodeType.NODE_PROCESSING_INSTRUCTION
'    Dim qm As String
'    If isProcInstr Then m_SB.Append ("?")
'    qm = IIf(isProcInstr, "'", """")
'    m_SB.Append nod.nodeName
'    Dim attrs As IXMLDOMNamedNodeMap: Set attrs = nod.Attributes
'    If Not attrs Is Nothing Then
'        If attrs.Length Then
'            m_SB.Append " "
'            PrintAttributes attrs, qm
'        End If
'    End If
'    Dim txt: txt = nod.nodeValue
'    Dim doEndtag As Boolean
'    doEndtag = nod.hasChildNodes Or Len(txt) <> 0
'    If Not doEndtag Then
'        m_SB.Append "/"
'    End If
'    m_SB.AppendLine ">"
'    'If Len(txt) Then Debug.Print txt
'    If Len(txt) Then m_SB.Append CStr(txt)
'    If nod.hasChildNodes Then
'        PrintNodes nod.childNodes
'    End If
'    If doEndtag Then
'        m_SB.Append("</").Append(nod.nodeName).AppendLine ">"
'    End If
End Sub

'Private Sub PrintNodeAttrib(nod As IXMLDOMNode)
'    'm_SB.Append attr.Name & "=" & quotemark & attr.Value & quotemark
'End Sub

Private Sub PrintNodeCData(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeComment(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeDocument(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeDocFrgmt(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeDocType(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeElement(nod As IXMLDOMNode)
    m_SB.AppendLine ""
    m_SB.Append IndStackPeek
    m_SB.Append("<").Append nod.nodeName
    Dim attrs As IXMLDOMNamedNodeMap: Set attrs = nod.Attributes
    If Not attrs Is Nothing Then
        If attrs.Length Then
            PrintNodeAttributes nod.Attributes
        End If
    End If
    If nod.hasChildNodes Then
        'Halt hier gehört nur ein neue-zeile zeichen wenn Childnodes
        m_SB.Append ">"
        IndStackPush
        Dim nodTxt As IXMLDOMText: Set nodTxt = nod.nodeValue
        PrintNodeText nodTxt
        PrintNodes nod.childNodes
        IndStackPop
        'm_SB.AppendLine ""
        m_SB.Append("</").Append(nod.nodeName).Append ">"
    Else
        m_SB.Append "/>"
    End If
End Sub

Private Sub PrintNodeEntity(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodEntRef(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeInvalid(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeNotation(nod As IXMLDOMNode)
    '
End Sub

Private Sub PrintNodeProcInstr(nod As IXMLDOMNode)
    m_SB.Append("<?").Append (nod.nodeName)
    Dim attrs As IXMLDOMNamedNodeMap: Set attrs = nod.Attributes
    If Not attrs Is Nothing Then
        If attrs.Length Then
            PrintNodeAttributes nod.Attributes, "'"
        End If
    End If
    m_SB.Append "?>"
End Sub

Private Sub PrintNodeText(nod As IXMLDOMNode)
    Dim txt: txt = nod.nodeValue
    m_SB.Append CStr(txt)
End Sub

Private Sub PrintNodeAttributes(attrs As IXMLDOMNamedNodeMap, Optional ByVal quotemark As String = """")
    Dim attr As IXMLDOMAttribute
    Dim i As Long, u As Long: u = attrs.Length - 1
    For i = 0 To u
        Set attr = attrs.Item(i)
        PrintNodeAttribute attr, quotemark
    Next
End Sub

Private Sub PrintNodeAttribute(attr As IXMLDOMAttribute, Optional ByVal quotemark As String = """")
    m_SB.Append " " & attr.Name & "=" & quotemark & attr.Value & quotemark
End Sub

Private Sub IndStackPush()
    m_IndStack = m_IndStack + 1
End Sub
Private Sub IndStackPop()
    m_IndStack = m_IndStack - 1
End Sub

Private Function IndStackPeek() As String
    IndStackPeek = Space(m_IndStack * 2)
End Function
