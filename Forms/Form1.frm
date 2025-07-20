VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
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
   Begin VB.CommandButton BtnOpenFile 
      Caption         =   "Open xml file"
      Height          =   375
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   480
      Width           =   12735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SXP As SmallXmlParser

Private Sub Form_Initialize()
    'Call System.InitSystem(App.path, App.EXEName, App.HelpFile, Command$)
End Sub

Private Sub Form_Load()
    Set SXP = New SmallXmlParser
End Sub

Private Sub BtnOpenFile_Click()
    Dim aPath As String
    Dim FS As FileStream
    Dim SR As StreamReader
    Dim DH As New DefaultHandler
    aPath = GetXmlFile
    If Len(aPath) > 0 Then
        Set FS = MNew.FileStream(aPath) ', FileMode_Input, FileAccess_Read, FileShare_None)
        Set SR = MNew.StreamReader(FS)
        DoEvents
        Call SXP.Parse(SR, DH)
        Text1.text = DH.ToStr
    End If
End Sub

Private Function GetXmlFile() As String
    With New OpenFileDialog
        .InitialDirectory = "C:\Windows\System32\"
        .Filter = "Xml-Dateien [*.xml]|*.xml"
        If .ShowDialog = VbMsgBoxResult.vbOK Then
            GetXmlFile = .FileName
        End If
    End With
End Function

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    T = Text1.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SXP = Nothing
End Sub
