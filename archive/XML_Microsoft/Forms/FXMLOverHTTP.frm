VERSION 5.00
Begin VB.Form FXMLOverHTTP 
   Caption         =   "XMLOverHTTP"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4005
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text3 
      Height          =   2775
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton BtnGet 
      Caption         =   "Get"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Service"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FXMLOverHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757026(v=vs.85)

'Source: XMLOverHTTP.frm
Public HttpReq As MSXML2.XMLHTTP60
Public XMLDoc  As MSXML2.DOMDocument60

Private Sub Form_Load()
    Text1.Text = "https://localhost/sxh/contact.asp"
    Text2.Text = "John Doe"
    'Text3.Text = ""
End Sub

Private Sub BtnGet_Click()
    MakeRequest True
End Sub

Private Sub MakeRequest(ByVal isAsync As Boolean)
    Set HttpReq = New XMLHTTP60
    Dim xhrHandler As HttpRequestHandler

    If isAsync = True Then
        Set xhrHandler = New HttpRequestHandler
        
        ' Set a readyStateChange handler.
        HttpReq.OnReadyStateChange = xhrHandler
    End If
    
    ' Construct the URL from user input.
    Dim url As String
    url = Text1.Text
    If Text2.Text <> "" Then
        url = url & "?SearchID=" & Text2.Text
    End If
    ' Clear the display.
    Text3.Text = ""
    
    ' Open a connection and set up a request to the server.
    HttpReq.open "GET", url, isAsync
        
    ' Send the request to the server.
    HttpReq.send
    
    ' In a synchronous call, we must call ProcessResponse. In an
    ' asynchronous call, the OnReadyStateChange handler calls
    ' ProcessResponse.
    If isAsync = False Then
        ProcessResponse
    End If
End Sub

Public Sub ProcessResponse()
    ' Receive the response from the server.
    Set XMLDoc = HttpReq.responseXML
    Dim Node As MSXML2.IXMLDOMNode
    ' Display the server response to the user.
    Set Node = XMLDoc.selectSingleNode("//phone")
    If Node Is Nothing Then
        Text3.Text = "Requested information not found."
    Else
        Text3.Text = Node.Text
    End If
End Sub
