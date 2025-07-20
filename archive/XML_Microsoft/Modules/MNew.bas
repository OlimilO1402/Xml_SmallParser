Attribute VB_Name = "MNew"
Option Explicit

Public Function DOMDocument60(Optional bAsync As Boolean = True, Optional doValidateOnParse As Boolean = True, Optional doResolveExternals As Boolean = False, Optional doPreserveWhitespace As Boolean = False) As DOMDocument60
    Set DOMDocument60 = New MSXML2.DOMDocument60
    With DOMDocument60
        .async = bAsync                            'by default true
        .validateOnParse = doValidateOnParse       'by default true
        .resolveExternals = doResolveExternals     'by default false anyway
        .preserveWhiteSpace = doPreserveWhitespace 'by default false
    End With
End Function

Public Function DOMDocument60F(ByVal aPathFileName As String) As DOMDocument60
    If Len(aPathFileName) = 0 Then Exit Function
    Set DOMDocument60F = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False, doResolveExternals:=False, doPreserveWhitespace:=True)
    If Not DOMDocument60F.Load(aPathFileName) Then
        MsgBox "Can't create DOM from " & aPathFileName
        Set DOMDocument60F = Nothing
        Exit Function
    End If
End Function

