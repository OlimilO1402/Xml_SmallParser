Attribute VB_Name = "MSmallXMLParser"
'Dies ist der von ICSharpCode.net von C# nach VB.NET umgewandelte Code
'Class DefaultHandler
'Inherits SmallXmlParser.IContentHandler
'
'  Public Sub OnStartParsing(ByVal parser As SmallXmlParser)
'  End Sub
'
'  Public Sub OnEndParsing(ByVal parser As SmallXmlParser)
'  End Sub
'
'  Public Sub OnStartElement(ByVal name As String, ByVal attrs As SmallXmlParser.IAttrList)
'  End Sub
'
'  Public Sub OnEndElement(ByVal name As String)
'  End Sub
'
'  Public Sub OnChars(ByVal s As String)
'  End Sub
'
'  Public Sub OnIgnorableWhitespace(ByVal s As String)
'  End Sub
'
'  Public Sub OnProcessingInstruction(ByVal name As String, ByVal text As String)
'  End Sub
'End Class
'
'Public Class SmallXmlParser
'
'  Public Interface IContentHandler
'
'    Sub OnStartParsing(ByVal parser As SmallXmlParser)
'
'    Sub OnEndParsing(ByVal parser As SmallXmlParser)
'
'    Sub OnStartElement(ByVal name As String, ByVal attrs As IAttrList)
'
'    Sub OnEndElement(ByVal name As String)
'
'    Sub OnProcessingInstruction(ByVal name As String, ByVal text As String)
'
'    Sub OnChars(ByVal text As String)
'
'    Sub OnIgnorableWhitespace(ByVal text As String)
'  End Interface
'
'  Public Interface IAttrList
'
'    ReadOnly Property Length() As Integer
'
'    ReadOnly Property IsEmpty() As Boolean
'
'    Function GetName(ByVal i As Integer) As String
'
'    Function GetValue(ByVal i As Integer) As String
'
'    Function GetValue(ByVal name As String) As String
'
'    ReadOnly Property Names() As String()
'
'    ReadOnly Property Values() As String()
'  End Interface
'
'  Class AttrListImpl
'  Implements IAttrList
'
'    Public ReadOnly Property Length() As Integer
'      Get
'        Return attrNames.Count
'      End Get
'    End Property
'
'    Public ReadOnly Property IsEmpty() As Boolean
'      Get
'        Return attrNames.Count = 0
'      End Get
'    End Property
'
'    Public Function GetName(ByVal i As Integer) As String
'      Return CType(attrNames(i), String)
'    End Function
'
'    Public Function GetValue(ByVal i As Integer) As String
'      Return CType(attrValues(i), String)
'    End Function
'
'    Public Function GetValue(ByVal name As String) As String
'      Dim i As Integer = 0
'      While i < attrNames.Count
'        If CType(attrNames(i), String) = name Then
'          Return CType(attrValues(i), String)
'        End If
'        System.Math.Min(System.Threading.Interlocked.Increment(i),i-1)
'      End While
'      Return Nothing
'    End Function
'
'    Public ReadOnly Property Names() As String()
'      Get
'        Return CType(attrNames.ToArray(GetType(String)), String())
'      End Get
'    End Property
'
'    Public ReadOnly Property Values() As String()
'      Get
'        Return CType(attrValues.ToArray(GetType(String)), String())
'      End Get
'    End Property
'    Private attrNames As ArrayList = New ArrayList
'    Private attrValues As ArrayList = New ArrayList
'
'    Friend Sub Clear()
'      attrNames.Clear
'      attrValues.Clear
'    End Sub
'
'    Friend Sub Add(ByVal name As String, ByVal value As String)
'      attrNames.Add (name)
'      attrValues.Add (value)
'    End Sub
'  End Class
'  Private handler As IContentHandler
'  Private reader As TextReader
'  Private elementNames As Stack = New Stack
'  Private xmlSpaces As Stack = New Stack
'  Private xmlSpace As String
'  Private buffer As StringBuilder = New StringBuilder(200)
'  Private nameBuffer As Char() = New Char(30) {}
'  Private IsWhitespace As Boolean
'  Private attributes As AttrListImpl = New AttrListImpl
'  Private line As Integer = 1
'  Private column As Integer
'  Private resetColumn As Boolean
'
'  Public Sub New()
'  End Sub
'
'  Private Function Error(ByVal msg As String) As Exception
'    Return New SmallXmlParserException(msg, line, column)
'  End Function
'
'  Private Function UnexpectedEndError() As Exception
'    Dim arr(elementNames.Count) As String
'    elementNames.CopyTo(arr, 0)
'    Return Error(String.Format("Unexpected end of stream. Element stack content is {0}", String.Join(",", arr)))
'  End Function
'
'  Private Function IsNameChar(ByVal c As cChar, ByVal start As Boolean) As Boolean
  Private Function IsNameChar(ByVal c As Integer, ByVal start As Boolean) As Boolean
'    Select c
'    Case ":"C, "_"C
'      Return True
'    Case "-"C, "."C
'      Return Not start
'    End Select
'    If c > 256 Then
'      Select c
'      Case "?"C, "?"C, "?"C
'        Return True
'      End Select
'      If "?"C <= c AndAlso c <= "i"C Then
'        Return True
'      End If
'    End If
'    Select Char.GetUnicodeCategory(c)
'    Case UnicodeCategory.LowercaseLetter, UnicodeCategory.UppercaseLetter, UnicodeCategory.OtherLetter, UnicodeCategory.TitlecaseLetter, UnicodeCategory.LetterNumber
'      Return True
'    Case UnicodeCategory.SpacingCombiningMark, UnicodeCategory.EnclosingMark, UnicodeCategory.NonSpacingMark, UnicodeCategory.ModifierLetter, UnicodeCategory.DecimalDigitNumber
'      Return Not start
'    Case Else
'      Return False
'    End Select
  End Function
'
  Private Function IsWhitespace(ByVal c As Integer) As Boolean
'    Select c
'    Case " "C, Microsoft.VisualBasic.Chr(13), Microsoft.VisualBasic.Chr(9), Microsoft.VisualBasic.Chr(10)
'      Return True
'    Case Else
'      Return False
'    End Select
  End Function
'
'  Public Sub SkipWhitespaces()
'    SkipWhitespaces (False)
'  End Sub
'
  Private Sub HandleWhitespaces()
'    While IsWhitespace(Peek)
'      buffer.Append (CType(Read, Char))
'    End While
'    If Not (Peek = "<"C) AndAlso Peek >= 0 Then
'      IsWhitespace = False
'    End If
  End Sub
'
  Public Sub SkipWhitespaces(ByVal expected As Boolean)
'    While True
'      Select Peek
'      Case " "C, Microsoft.VisualBasic.Chr(13), Microsoft.VisualBasic.Chr(9), Microsoft.VisualBasic.Chr(10)
'        Read
'        If expected Then
'          expected = False
'        End If
'        ' continue
'      End Select
'      If expected Then
'        Throw Error("Whitespace is expected.")
'      End If
'      Return
'    End While
  End Sub
'
  Private Function Peek() As Integer
'    Return reader.Peek
  End Function
'
  Private Function Read() As Integer
'    Dim i As Integer = reader.Read
'    If i = Microsoft.VisualBasic.Chr(10) Then
'      resetColumn = True
'    End If
'    If resetColumn Then
'      line = line + 1 'System.Math.Min(System.Threading.Interlocked.Increment(line),line-1)
'      resetColumn = False
'      column = 1
'    Else
'      Column = Column +1 'System.Math.Min(System.Threading.Interlocked.Increment(column),column-1)
'    End If
'    Return i
  End Function
'
  Public Sub Expect(ByVal c As Integer)
'    Dim p As Integer = Read
'    If p < 0 Then
'      Throw UnexpectedEndError
'    Else
'      If Not (p = c) Then
'        Throw Error(String.Format("Expected '{0}' but got {1}", CType(c, Char), CType(p, Char)))
'      End If
'    End If
  End Sub
'
'  Private Function ReadUntil(ByVal cuntil As cChar, ByVal handleReferences As Boolean) As String
  Private Function ReadUntil(ByVal cuntil As Integer, ByVal handleReferences As Boolean) As String
'    While True
'      If Peek < 0 Then
'        Throw UnexpectedEndError
'      End If
'      Dim c As Char = CType(Read, Char)
'      If c = until Then
'        ' break
'      Else
'        If handleReferences AndAlso c = "&"C Then
'          ReadReference
'        Else
'          buffer.Append (c)
'        End If
'      End If
'    End While
'    Dim ret As String = buffer.ToString
'    buffer.length = 0
'    Return ret
  End Function
'
  Public Function ReadName() As String
'    Dim idx As Integer = 0
'    If Peek < 0 OrElse Not IsNameChar(CType(Peek, Char), True) Then
'      Throw Error("XML name start character is expected.")
'    End If
'    Dim i As Integer = Peek
'    While i >= 0
'      Dim c As Char = CType(i, Char)
'      If Not IsNameChar(c, False) Then
'        ' break
'      End If
'      If idx = nameBuffer.length Then
'        Dim tmp(idx * 2) As Char
'        Array.Copy(nameBuffer, tmp, idx)
'        nameBuffer = tmp
'      End If
'      nameBuffer(System.Math.Min(System.Threading.Interlocked.Increment(idx), idx - 1)) = c
'      Read
'      i = Peek
'    End While
'    If idx = 0 Then
'      Throw Error("Valid XML name is expected.")
'    End If
'    Return New String(nameBuffer, 0, idx)
  End Function
'
  Public Sub Parse(ByVal trinput As TextReader, ByVal handler As IContentHandler)
'    Me.reader = input
'    Me.handler = handler
'    handler.OnStartParsing (Me)
'    While Peek >= 0
'      ReadContent
'    End While
'    HandleBufferedContent
'    If elementNames.Count > 0 Then
'      Throw Error(String.Format("Insufficient close tag: {0}", elementNames.Peek))
'    End If
'    handler.OnEndParsing (Me)
'    Cleanup
  End Sub
'
  Private Sub Cleanup()
'    line = 1
'    column = 0
'    handler = Nothing
'    reader = Nothing
'    elementNames.Clear
'    xmlSpaces.Clear
'    attributes.Clear
'    buffer.length = 0
'    xmlSpace = Nothing
'    IsWhitespace = False
  End Sub
'
  Public Sub ReadContent()
'    Dim name As String
'    If IsWhitespace(Peek) Then
'      If buffer.length = 0 Then
'        IsWhitespace = True
'      End If
'      HandleWhitespaces
'    End If
'    If Peek = "<"C Then
'      Read
'      Select Peek
'      Case "!"C
'        Read
'        If Peek = "["C Then
'          Read
'          If Not (ReadName = "CDATA") Then
'            Throw Error("Invalid declaration markup")
'          End If
'          Expect("["C)
'          ReadCDATASection
'          Return
'        Else
'          If Peek = "-"C Then
'            ReadComment
'            Return
'          Else
'            If Not (ReadName = "DOCTYPE") Then
'              Throw Error("Invalid declaration markup.")
'            Else
'              Throw Error("This parser does not support document type.")
'            End If
'          End If
'        End If
'      Case "?"C
'        HandleBufferedContent
'        Read
'        name = ReadName
'        SkipWhitespaces
'        Dim text As String = String.Empty
'        If Not (Peek = "?"C) Then
'          While True
'            text += ReadUntil("?"C, False)
'            If Peek = ">"C Then
'              ' break
'            End If
'            text += "?"
'          End While
'        End If
'        handler.OnProcessingInstruction(name, text)
'        Expect(">"C)
'        Return
'      Case "/"C
'        HandleBufferedContent
'        If elementNames.Count = 0 Then
'          Throw UnexpectedEndError
'        End If
'        Read
'        name = ReadName
'        SkipWhitespaces
'        Dim expected As String = CType(elementNames.Pop, String)
'        xmlSpaces.Pop
'        If xmlSpaces.Count > 0 Then
'          xmlSpace = CType(xmlSpaces.Peek, String)
'        Else
'          xmlSpace = Nothing
'        End If
'        If Not (name = expected) Then
'          Throw Error(String.Format("End tag mismatch: expected {0} but found {1}", expected, name))
'        End If
'        handler.OnEndElement (name)
'        Expect(">"C)
'        Return
'      Case Else
'        HandleBufferedContent
'        name = ReadName
'        While Not (Peek = ">"C) AndAlso Not (Peek = "/"C)
'          ReadAttribute (attributes)
'        End While
'        handler.OnStartElement(name, attributes)
'        attributes.Clear
'        SkipWhitespaces
'        If Peek = "/"C Then
'          Read
'          handler.OnEndElement (name)
'        Else
'          elementNames.Push (name)
'          xmlSpaces.Push (xmlSpace)
'        End If
'        Expect(">"C)
'        Return
'      End Select
'    Else
'      ReadCharacters
'    End If
  End Sub
'
  Private Sub HandleBufferedContent()
'    If buffer.length = 0 Then
'      Return
'    End If
'    If IsWhitespace Then
'      handler.OnIgnorableWhitespace (buffer.ToString)
'    Else
'      handler.OnChars (buffer.ToString)
'    End If
'    buffer.length = 0
'    IsWhitespace = False
  End Sub
'
  Private Sub ReadCharacters()
'    IsWhitespace = False
'    While True
'      Dim i As Integer = Peek
'      Select i
'      Case -1
'        Return
'      Case "<"C
'        Return
'      Case "&"C
'        Read
'        ReadReference
'        ' continue
'      Case Else
'        buffer.Append (CType(Read, Char))
'        ' continue
'      End Select
'    End While
  End Sub
'
  Private Sub ReadReference()
'    If Peek = "#"C Then
'      Read
'      ReadCharacterReference
'    Else
'      Dim name As String = ReadName
'      Expect(";"C)
'      Select name
'      Case "amp"
'        buffer.Append("&"C)
'        ' break
'      Case "quot"
'        buffer.Append(""""C)
'        ' break
'      Case "apos"
'        buffer.Append("'"C)
'        ' break
'      Case "lt"
'        buffer.Append("<"C)
'        ' break
'      Case "gt"
'        buffer.Append(">"C)
'        ' break
'      Case Else
'        Throw Error("General non-predefined entity reference is not supported in this parser.")
'      End Select
'    End If
  End Sub
'
  Private Function ReadCharacterReference() As Integer
'    Dim n As Integer = 0
'    If Peek = "x"C Then
'      Read
'      Dim i As Integer = Peek
'      While i >= 0
'        If "0"C <= i AndAlso i <= "9"C Then
'          n = n << 4 + i - "0"C
'        Else
'          If "A"C <= i AndAlso i <= "F"C Then
'            n = n << 4 + i - "A"C + 10
'          Else
'            If "a"C <= i AndAlso i <= "f"C Then
'              n = n << 4 + i - "a"C + 10
'            Else
'              ' break
'            End If
'          End If
'        End If
'        Read
'        i = Peek
'      End While
'    Else
'      Dim i As Integer = Peek
'      While i >= 0
'        If "0"C <= i AndAlso i <= "9"C Then
'          n = n << 4 + i - "0"C
'        Else
'          ' break
'        End If
'        Read
'        i = Peek
'      End While
'    End If
'    Return n
  End Function
'
  Private Sub ReadAttribute(ByVal a As AttrListImpl)
'    SkipWhitespaces (True)
'    If Peek = "/"C OrElse Peek = ">"C Then
'      Return
'    End If
'    Dim name As String = ReadName
'    Dim value As String
'    SkipWhitespaces
'    Expect("="C)
'    SkipWhitespaces
'    Select Read
'    Case "'"C
'      value = ReadUntil("'"C, True)
'      ' break
'    Case """"C
'      value = ReadUntil(""""C, True)
'      ' break
'    Case Else
'      Throw Error("Invalid attribute value markup.")
'    End Select
'    If name = "xml:space" Then
'      xmlSpace = value
'    End If
'    a.Add(name, value)
  End Sub
'
  Private Sub ReadCDATASection()
'    Dim nBracket As Integer = 0
'    While True
'      If Peek < 0 Then
'        Throw UnexpectedEndError
'      End If
'      Dim c As Char = CType(Read, Char)
'      If c = "]"C Then
'        System.Math.Min(System.Threading.Interlocked.Increment(nBracket),nBracket-1)
'      Else
'        If c = ">"C AndAlso nBracket > 1 Then
'          Dim i As Integer = nBracket
'          While i > 2
'            buffer.Append("]"C)
'            System.Math.Max(System.Threading.Interlocked.Decrement(i),i+1)
'          End While
'          ' break
'        Else
'          Dim i As Integer = 0
'          While i < nBracket
'            buffer.Append("]"C)
'            System.Math.Min(System.Threading.Interlocked.Increment(i),i-1)
'          End While
'          nBracket = 0
'          buffer.Append (c)
'        End If
'      End If
'    End While
  End Sub
'
  Private Sub ReadComment()
'    Expect("-"C)
'    Expect("-"C)
'    While True
'      If Not (Read = "-"C) Then
'        ' continue
'      End If
'      If Not (Read = "-"C) Then
'        ' continue
'      End If
'      If Not (Read = ">"C) Then
'        Throw Error("'--' is not allowed inside comment markup.")
'      End If
'      ' break
'    End While
  End Sub
'End Class
'
'Public Class SmallXmlParserException
'Inherits SystemException
'  Private line As Integer
'  Private column As Integer
'
'  Public Sub New(ByVal msg As String, ByVal line As Integer, ByVal column As Integer)
'    MyBase.New(String.Format("{0}. At ({1},{2})", msg, line, column))
'    Me.line = line
'    Me.column = column
'  End Sub
'
'  Public ReadOnly Property Line() As Integer
'    Get
'      Return line
'    End Get
'  End Property
'
'  Public ReadOnly Property Column() As Integer
'    Get
'      Return column
'    End Get
'  End Property
'End Class
'
