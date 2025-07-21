VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmallXmlParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private handler       As IContentHandler
Private reader        As TextReader
Private elementNames  As Stack
Private xmlSpaces     As Stack
Private xmlSpace      As String
Private buffer        As StringBuilder
Private nameBuffer    As String ''Char() = New Char(30) {}
Private mIsWhitespace As Boolean
Private attributes    As AttrListImpl
Private Line          As Long
Private Column        As Long
Private resetColumn   As Boolean

Private Sub Class_Initialize()
    Set elementNames = New Stack
    Set xmlSpaces = New Stack
    Set buffer = New StringBuilder ''MNew.StringBuilder(, , , 200)
    Set attributes = New AttrListImpl
    Line = 1
End Sub
Private Sub Class_Terminate()
    Call Cleanup
End Sub

Private Function Error(ByVal msg As String) As SmallXmlParserException 'Exception
    Set Error = New SmallXmlParserException:  Error.New_ msg, Line, Column
End Function

Private Function UnexpectedEndError() As SmallXmlParserException 'Exception
    ReDim arr(elementNames.Count) As String
    'elementNames.CopyTo(arr, 0)
    Set UnexpectedEndError = Error("Unexpected end of stream. Element stack content is " & Join(arr))
End Function

Private Function IsNameChar(c As Long, start As Boolean) As Boolean
    'Dim bc As Long
    'bc = AscW(c)
    Select Case c
    Case 58, 95 'AscW(":"), AscW("_")
        IsNameChar = True: Exit Function
    Case 45, 46 'AscW("-"), AscW(".")
        IsNameChar = Not start: Exit Function
    End Select
    If c > &H100 Then
        Select Case c
        Case &H559, &H6E5, &H6E6   ''\u0559'"?", '\u06E5'"?", '\u06E6'"?"
            IsNameChar = True: Exit Function
        End Select
        If (&H2BB <= c) And (c <= &H2C1) Then 'AscW("?")=63; AscW("i")=105
            IsNameChar = True: Exit Function
        End If
    End If
    Select Case c
    Case 97 To 122, 223, 228, 246, 252 'a..z,ß,ä,ö,ü LowercaseLetter
        IsNameChar = True: Exit Function
    Case 65 To 90, 196, 214, 220       'A..Z,  Ä,Ö,Ü UppercaseLetter
        IsNameChar = True: Exit Function
    Case 48 To 57                      '0...9,       DecimalDigitNumber
        IsNameChar = Not start: Exit Function
    Case 40, 41, 91, 93, 123, 125  ' (, ), [, ], {, } EnclosingMark
        IsNameChar = Not start: Exit Function
    Case Else
        IsNameChar = False
    End Select
    '  Select Case Char.GetUnicodeCategory(c)
    '    Case UnicodeCategory.LowercaseLetter, UnicodeCategory.UppercaseLetter, UnicodeCategory.OtherLetter, UnicodeCategory.TitlecaseLetter, UnicodeCategory.LetterNumber
    '      IsNameChar = True
    '    Case UnicodeCategory.SpacingCombiningMark, UnicodeCategory.EnclosingMark, UnicodeCategory.NonSpacingMark, UnicodeCategory.ModifierLetter, UnicodeCategory.DecimalDigitNumber
    '      IsNameChar = Not start
    '    Case Else
    '      IsNameChar = False
    '  End Select
End Function

Private Function IsWhitespace(c As Long) As Boolean
    Select Case c 'AscW(" "), AscW(vbCr), AscW(vbLF), AscW(vbTab)
    Case 32, 13, 10, 9: IsWhitespace = True
    Case Else:          IsWhitespace = False
    End Select
End Function

Private Sub HandleWhitespaces()
    While IsWhitespace(Peek)
        Call buffer.Append(ChrW$(Read))
    Wend 'While
    If (Peek <> 60) And (Peek >= 0) Then '"<"
        mIsWhitespace = False
    End If
End Sub

Public Sub SkipWhitespaces(Optional expected As Boolean = False)
    While True
        Select Case Peek
        Case 9, 10, 13, 32 'Chr$(9), Chr$(10), Chr$(13), " "
            Read
            If expected Then
                expected = False
            End If
            ' continue
        End Select
        If expected Then
            'Throw Error("Whitespace is expected.")
            MsgBox "Error: Whitespace is expected."
            Debug.Assert (False)
        End If
        Exit Sub 'Return
    Wend 'While
End Sub

Private Function Peek() As Long
    Peek = reader.Peek
End Function

Private Function Read() As Long
    Dim i As Long: i = reader.Read
    If i = 10 Then 'Chr(10)
        resetColumn = True
    End If
    If resetColumn Then
        'Min(System.Threading.Interlocked.Increment(line),line-1)
        Line = Line + 1
        resetColumn = False
        Column = 1
    Else
        'Min(System.Threading.Interlocked.Increment(column),column-1)
        Column = Column + 1
    End If
    Read = i
End Function

Public Sub Expect(c As Long)
    Dim p As Long: p = Read 'ChrW$(62) = ">"
    If p < 0 Then
        'Throw UnexpectedEndError
        MsgBox "Error: UnexpectedEndError"
        Debug.Assert (False)
    Else
        If Not (p = c) Then
            'Throw Error("Expected '" & Chr$(c) & "' but got " & Chr$(p))
            MsgBox "Error: Expected '" & ChrW$(c) & "' but got " & ChrW$(p)
            Debug.Assert (False)
        End If
    End If
End Sub

Private Function ReadUntil(cuntil As Long, handleReferences As Boolean) As String
    Dim c As Long
    'Dim ret As String
    Do While True
        If Peek < 0 Then
            'Throw UnexpectedEndError
            MsgBox "Error: UnexpectedEndError"
            Debug.Assert (False)
        End If
        c = Read ', Char)
        If c = cuntil Then
            Exit Do ' break
        ElseIf (handleReferences) And (c = 38) Then '"&"
            ReadReference
        Else
            Call buffer.Append(ChrW$(c))
        End If
    Loop 'Wend 'End While
    'ret = buffer.ToStr
    ReadUntil = buffer.ToStr
    buffer.Length = 0
    'ReadUntil = ret
End Function

Public Function ReadName() As String
Dim idx As Long ': idx = 0
Dim tmp As String
Dim i As Long
Dim c As Long
Dim aSB As StringBuilder: Set aSB = MNew.StringBuilder(, , , 200)
  'nameBuffer = vbNullString
  '
  If (Peek < 0) Or (Not IsNameChar(Peek, True)) Then
    'Throw Error("XML name start character is expected.")
    MsgBox "Error: XML name start character is expected."
    Debug.Assert (False)
  End If
  i = Peek
  Do While i >= 0
    c = i '
    If Not IsNameChar(c, False) Then
      Exit Do ' break
    End If
    'If idx = Len(nameBuffer) Then 'nameBuffer.Length
    '  ReDim tmp(idx * 2) As Byte 'String * 1 'Char
    '  '(idx*2)
    '  'Copy(sourceArray, destinationArray, length)
    '  'Array.Copy(nameBuffer, tmp, idx)
    '  'nameBuffer = tmp
    '  nameBuffer = tmp
    'End If
    'nameBuffer[idx++] = c;
    'nameBuffer(System.Math.Min(System.Threading.Interlocked.Increment(idx), idx - 1)) = c
    idx = idx + 1
    'nameBuffer = nameBuffer & ChrW$(c)
    Call aSB.Append(ChrW$(c))
    Read
    i = Peek
  Loop 'Wend 'While
  If idx = 0 Then
    'Throw Error("Valid XML name is expected.")
    MsgBox "Error: Valid XML name is expected."
    Debug.Assert (False)
  End If
  'nameBuffer = mSB.ToStr
  ReadName = aSB.ToStr 'nameBuffer
End Function

Public Sub Parse(trinput As TextReader, ahandler As IContentHandler)
  Set reader = trinput
  Set handler = ahandler
  Call handler.OnStartParsing(Me)
  While Peek >= 0
    ReadContent
  Wend 'While
  HandleBufferedContent
  If elementNames.Count > 0 Then
    'Throw Error("Insufficient close tag: " & CStr(elementNames.Peek))
    MsgBox "Error: Insufficient close tag: " & ChrW$(elementNames.Peek)
    Debug.Assert (False)
  End If
  Call handler.OnEndParsing(Me)
  Cleanup
End Sub

Public Sub ReadContent()
Dim name As String
Dim text As String
Dim expected As String
  If IsWhitespace(Peek) Then
    If buffer.Length = 0 Then
      mIsWhitespace = True
    End If
    HandleWhitespaces
  End If
  If Peek = 60 Then '"<"
    Read
    Select Case Peek
    Case 33 '"!"' // declarations
      Read
      If Peek = 91 Then '"["
        Read
        If Not (ReadName = "CDATA") Then
          'Throw Error("Invalid declaration markup")
          MsgBox "Error: Invalid declaration markup"
          Debug.Assert (False)
        End If
        Call Expect(91) '"["
        ReadCDATASection
        Exit Sub 'Return
      Else
        If Peek = 45 Then '"-"
          ReadComment
          Exit Sub 'Return
        Else
          If Not (ReadName = "DOCTYPE") Then
            'Throw Error("Invalid declaration markup.")
            MsgBox "Error: Invalid declaration markup."
            Debug.Assert (False)
          Else
            'Throw Error("This parser does not support document type.")
            MsgBox "Error: This parser does not support document type."
            Debug.Assert (False)
          End If
        End If
      End If
    Case 63 '"?"' // PIs
      HandleBufferedContent
      Read
      name = ReadName
      SkipWhitespaces
      '= String.Empty
      If Not (Peek = 63) Then '"?"
        Do While True
          text = text & ReadUntil(63, False) '"?"
          If Peek = 62 Then '">"
            Exit Do ' break
          End If
          text = text & "?"
        Loop 'Wend 'End While
      End If
      Call handler.OnProcessingInstruction(name, text)
      Call Expect(62) '">"
      Exit Sub 'Return
    Case 47 '"/"  '// end tags
      HandleBufferedContent
      If elementNames.Count = 0 Then
        'Throw UnexpectedEndError
        MsgBox "Error: UnexpectedEndError"
        Debug.Assert (False)
      End If
      Read
      name = ReadName
      SkipWhitespaces
      expected = elementNames.Pop ', String)
      xmlSpaces.Pop
      If xmlSpaces.Count > 0 Then
        xmlSpace = xmlSpaces.Peek ', String)
      Else
        xmlSpace = vbNullString 'Nothing
      End If
      If Not (name = expected) Then
        'Throw Error("End tag mismatch: expected " & expected & " but found " & name)
        MsgBox "Error: End tag mismatch: expected " & expected & " but found " & name
        Debug.Assert (False)
      End If
      Call handler.OnEndElement(name)
      Call Expect(62)  '">"
      Exit Sub 'Return
    Case Else '// start tags (including empty tags)
      HandleBufferedContent
      name = ReadName
      While ((Peek <> 62) And (Peek <> 47)) 'AscW(">")=62 ; AscW("/")=47
        Call ReadAttribute(attributes)
      Wend 'End While
      Call handler.OnStartElement(name, attributes)
      attributes.Clear
      SkipWhitespaces
      If Peek = 47 Then '"/"
        Read
        Call handler.OnEndElement(name)
      Else
        Call elementNames.Push(name)
        Call xmlSpaces.Push(xmlSpace)
      End If
      Call Expect(62) '">"
      Exit Sub 'Return
    End Select
  Else
    ReadCharacters
  End If
End Sub

Private Sub HandleBufferedContent()
  If buffer.Length = 0 Then
    Exit Sub 'Return
  End If
  If mIsWhitespace Then
    Call handler.OnIgnorableWhitespace(buffer.ToStr)
  Else
    Call handler.OnChars(buffer.ToStr)
  End If
  buffer.Length = 0
  mIsWhitespace = False
End Sub

Private Sub ReadCharacters()
Dim i As Long
  mIsWhitespace = False
  While True
    i = Peek
    Select Case i
    Case -1
      Exit Sub 'Return
    Case 60 '"<"
      Exit Sub 'Return
    Case 38 '"&"
      Read
      ReadReference
      ' continue
    Case Else
      Call buffer.Append(ChrW$(Read))
      ' continue
    End Select
  Wend 'While
End Sub

Private Sub ReadReference()
Dim name As String
  If Peek = 35 Then '"#"
    Read
    ReadCharacterReference
  Else
    name = ReadName
    Call Expect(59) '";"
    Select Case name
    Case "amp"
      Call buffer.Append("&")
      ' break
    Case "quot"
      Call buffer.Append("""")
      ' break
    Case "apos"
      Call buffer.Append("'")
      ' break
    Case "lt"
      Call buffer.Append("<")
      ' break
    Case "gt"
      Call buffer.Append(">")
      ' break
    Case Else
      'Throw Error("General non-predefined entity reference is not supported in this parser.")
      MsgBox "Error: General non-predefined entity reference is not supported in this parser."
      Debug.Assert (False)
    End Select
  End If
End Sub

Private Function ReadCharacterReference() As Long
Dim n As Long: n = 0
Dim i As Long
  'e.g: &#x20; for a space " "
  If Peek = 120 Then 'AscW("x")=120
    Read
    i = Peek
    'eine Ansi-HEX-Zahl
    Do While i >= 0
      If (48 <= i) And (i <= 57) Then '"0"=48; "9"=57
        'n = n << 4 + i - 48 '"0"
        n = MMath.ShL(n, 4) + i - 48
      Else
        If (65 <= i) And (i <= 70) Then 'AscW("A")=65; AscW("F")=70;
          'n = n << 4 + i - 65 + 10'"A"
          n = MMath.ShL(n, 4) + i - 65 + 10
        Else
          If (97 <= i) And (i <= 102) Then 'AscW("a")=97; AscW("f")=102;
            'n = n << 4 + i - 97 + 10'"a"
            n = MMath.ShL(n, 4) + i - 97 + 10
          Else
            Exit Do ' break
          End If
        End If
      End If
      Read
      i = Peek
    Loop 'Wend 'End While
  Else
    i = Peek
    Do While i >= 0
      If (48 <= i) And (i <= 57) Then  'AscW("0")=48; AscW("9")=57;
        'n = n << 4 + i - 48'"0"
        n = MMath.ShL(n, 4) + i - 48
      Else
        Exit Do ' break
      End If
      Read
      i = Peek
    Loop 'Wend 'End While
  End If
  ReadCharacterReference = n
End Function

Private Sub ReadAttribute(a As AttrListImpl)
  Call SkipWhitespaces(True)
  If (Peek = 47) Or (Peek = 62) Then 'AscW("/")=47; AscW(">")=62
    Exit Sub 'Return
  End If
  Dim name As String: name = ReadName
  Dim value As String
  SkipWhitespaces
  Call Expect(61) '"="
  SkipWhitespaces
  Select Case Read
  Case 39 ' AscW("'")
    value = ReadUntil(39, True) '"'"
    ' break
  Case 34 '""""
    value = ReadUntil(34, True) '""""
    ' break
  Case Else
    'Throw Error("Invalid attribute value markup.")
    MsgBox "Error: Invalid attribute value markup."
    Debug.Assert (False)
  End Select
  If name = "xml:space" Then
    xmlSpace = value
  End If
  Call a.Add(name, value)
End Sub

Private Sub ReadCDATASection()
Dim nBracket As Long
Dim c As Long 'Char
Dim i As Long
  nBracket = 0
  Do While True
    If Peek < 0 Then
      'Throw UnexpectedEndError
      MsgBox "Error: UnexpectedEndError"
      Debug.Assert (False)
    End If
    c = Read ', Char)
    If c = 93 Then '"]"
      nBracket = nBracket + 1 'Call Min(Increment(nBracket), nBracket - 1)
    Else
      If (c = 62) And (nBracket > 1) Then '">"
        i = nBracket
        While i > 2
          Call buffer.Append("]")
          i = i - 1 'Call Max(Decrement(i), i + 1)
        Wend 'While
        Exit Do ' break
      Else
        i = 0
        While i < nBracket
          Call buffer.Append("]")
          i = i + 1 'Min(Increment(i),i-1)
        Wend 'While
        nBracket = 0
        Call buffer.Append(ChrW$(c))
      End If
    End If
  Loop 'Wend 'While
End Sub

Private Sub ReadComment()
  Call Expect(45) '"-"
  Call Expect(45) '"-"
  Do While True
    If Not (Read = 45) Then '"-"
      ' continue
    End If
    If Not (Read = 45) Then '"-"
      ' continue
    End If
    If Not (Read = 62) Then '">"
      'Throw Error("'--' is not allowed inside comment markup.")
      MsgBox "Error: '--' is not allowed inside comment markup."
      Debug.Assert (False)
    End If
    Exit Do ' break
  Loop 'Wend 'While
End Sub

Private Sub Cleanup()
  Line = 1
  Column = 0
  Set handler = Nothing
  Set reader = Nothing
  elementNames.Clear
  xmlSpaces.Clear
  attributes.Clear
  buffer.Length = 0
  'Set xmlSpace = Nothing
  xmlSpace = vbNullString
  mIsWhitespace = False
End Sub

