VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class StringBuilder
'          Inherits System.Object
'     Member von: System.Text
'Zusammenfassung:
' Stellt eine veränderbare Zeichenfolge dar. Diese Klasse kann nicht geerbt werden.
'
'Comments:
'dieser Klasse muss zuerst mit New ein String hinzugefügt werden,
'bevor etwas damit gemacht werden kann. deshalb auch nicht GlobalMultiUse sondern nur MultiUse
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private mStrValue As String  'm_sString
Private mCachedStr As String
Private mLength As Long   'm_iPos 'in Byte!!!!
'Private mPos As Long
Private mCapacity As Long 'm_iLen
Private mMaxCapacity As Long '= 2147483647 'braucht man das überhaupt?
Private Const constDefaultCapacity = 8192 '16384 '16 'm_iChunkSize
'
'Many Thanks To:
'Steve McMahon (steve@vbaccelerator.com)
'Mono .NET, all the guys around it and Novell
'especially the Authors of System.String.cs, and System.StringBuilder.cs
'Matthew Curland for "Advanced Visual Basic 6"
'
Private Sub Class_Initialize()
  
   ' The default allocation: 8192 characters.
   'm_iChunkSize = constDefaultCapacity '16384
   'mCapacity = constDefaultCapacity
   mMaxCapacity = 2147483647
   mCapacity = constDefaultCapacity
   mStrValue = InternalAllocateStr(mCapacity)
End Sub

'Public Sub New()
'Public Sub New(ByVal capacity As Integer)
'Public Sub New(ByVal capacity As Integer, ByVal maxCapacity As Integer)
'Public Sub New(ByVal value As String)
'Public Sub New(ByVal value As String, ByVal capacity As Integer)
'Public Sub New(ByVal value As String, ByVal startIndex As Integer, ByVal length As Integer, ByVal capacity As Integer)
'    public StringBuilder(string value, int startIndex, int length, int capacity)
'    {
'      // first, check the parameters and throw appropriate exceptions if needed
'      if (null == value)
'        value = "";
'
'      // make sure startIndex is zero or positive
'      if (startIndex < 0)
'        throw new System.ArgumentOutOfRangeException ("startIndex", startIndex, "StartIndex cannot be less than zero.");
'
'      // make sure length is zero or positive
'      if(length < 0)
'        throw new System.ArgumentOutOfRangeException ("length", length, "Length cannot be less than zero.");
'
'      if (capacity < 0)
'        throw new System.ArgumentOutOfRangeException ("capacity", capacity, "capacity must be greater than zero.");
'
'      // make sure startIndex and length give a valid substring of value
'      // re-ordered to avoid possible integer overflow
'      if (startIndex > value.Length - length)
'        throw new System.ArgumentOutOfRangeException ("startIndex", startIndex, "StartIndex and length must refer to a location within the string.");
'
'      if (capacity == 0)
'        capacity = constDefaultCapacity;
'
'      _str = String.InternalAllocateStr ((length > capacity) ? length : capacity);
'      if (length > 0)
'        String.InternalStrcpy(_str, 0, value, startIndex, length);
'
'      _length = length;
'    }
Friend Sub New_(Optional ByVal sValue As String, Optional ByVal startIndex As Long, Optional ByVal lLength As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long)
    Dim n As Long, IAS As Long
    'If sValue = vbNullString Then sValue = ""
    'If startIndex < 0 Then Call Throw_New_ArgumentOutOfRangeException("startIndex", CStr(startIndex), "StartIndex cannot be less than zero.")
    'If length < 0 Then Call Throw_New_ArgumentOutOfRangeException("length", CStr(length), "length cannot be less than zero.")
    'If capacity < 0 Then Call Throw_New_ArgumentOutOfRangeException("capacity", CStr(capacity), "capacity cannot be less than zero.")
    'Make sure Startindex and length give a valid substring of value
    're-ordered to avoid possible integer overflow
    'If startIndex > n - length Then Call Throw_New_ArgumentOutOfRangeException("startIndex", CStr(startIndex), "StartIndex and length must refer to a location within the string.")
    n = Len(sValue)
    If Capacity = 0 Then
        mCapacity = constDefaultCapacity
    Else
        mCapacity = Capacity
    End If
    If maxCapacity > 0 Then mMaxCapacity = maxCapacity
    If n > mCapacity Then IAS = n Else IAS = mCapacity
    mStrValue = InternalAllocateStr(IAS)
    If lLength = 0 Then lLength = n
    If lLength > 0 Then
        Call InternalStrcpy(mStrValue, 0, sValue, startIndex, lLength)
    End If
    mLength = lLength * 2
End Sub
Public Function StringBuilder(Optional ByVal value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
    Set StringBuilder = New StringBuilder: StringBuilder.New_ value, startIndex, Length, Capacity, maxCapacity
End Function



'Public Function Append(ByVal value As Boolean) As System.Text.StringBuilder
'Public Function Append(ByVal value As Byte) As System.Text.StringBuilder
'Public Function Append(ByVal value() As Char) As System.Text.StringBuilder
'Public Function Append(ByVal value() As Char, ByVal startIndex As Integer, ByVal charCount As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Char) As System.Text.StringBuilder
'Public Function Append(ByVal value As Char, ByVal repeatCount As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Decimal) As System.Text.StringBuilder
'Public Function Append(ByVal value As Double) As System.Text.StringBuilder
'Public Function Append(ByVal value As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Long) As System.Text.StringBuilder
'Public Function Append(ByVal value As Object) As System.Text.StringBuilder
'Public Function Append(ByVal value As Short) As System.Text.StringBuilder
'Public Function Append(ByVal value As Single) As System.Text.StringBuilder
'Public Function Append(ByVal value As String) As System.Text.StringBuilder
'Public Function Append(ByVal value As String, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.SByte) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt16) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt32) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt64) As System.Text.StringBuilder
'Public Function Append(ByVal varValue As Variant) As StringBuilder
'  mValue = mValue & CStr(varValue)
'  Set Append = Me
'End Function
Public Function Append(ByRef StrValue As String) As StringBuilder
Dim lLen As Long, n As Long, rLen As Long
   'm_iPos = mLength
   'm_iLen = mCapacity
   ' Append an item to the string:
   n = Len(StrValue)
   lLen = 2 * n
   rLen = mLength + lLen 'die resultierende Länge
   'gut ist, wenn mCapacity immer größer ist!
   If rLen > mCapacity Then
     'constDefaultCapacity sollte natürlich so groß sein,
     'daß diese Schleife so wenig wie möglich ausgeführt wird
     Do
       mStrValue = mStrValue & Space$(mCapacity)
       mCapacity = mCapacity + mCapacity
     Loop While mCapacity < rLen
   End If
   'CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
   CopyMemory ByVal UnsignedAdd(StrPtr(mStrValue), mLength), ByVal StrPtr(StrValue), lLen
   mLength = mLength + lLen
   Set Append = Me
End Function
'Public Sub AppendSub(ByRef StrValue As String) 'As StringBuilder
'Dim lLen As Long, BLen As Long, n As Long, rLen As Long
'   'm_iPos = mLength
'   'm_iLen = mCapacity
'   ' Append an item to the string:
'   n = Len(StrValue)
'   lLen = 2 * n
'   BLen = 2 * mLength
'   rLen = BLen + lLen 'die resultierende Länge
'   'gut ist, wenn mCapacity immer größer ist!
'   If rLen > mCapacity Then
'     'constDefaultCapacity sollte natürlich so groß sein,
'     'daß diese Schleife so wenig wie möglich ausgeführt wird
'     Do
'       mStrValue = mStrValue & Space$(mCapacity)
'       mCapacity = mCapacity + mCapacity
'       'mStrValue = mStrValue & Space$(constDefaultCapacity)
'       'mCapacity = mCapacity + constDefaultCapacity
'     Loop While mCapacity < rLen
'   End If
'   'CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
'   CopyMemory ByVal UnsignedAdd(StrPtr(mStrValue), BLen), ByVal StrPtr(StrValue), lLen
'   mLength = mLength + n
'   'Set Append = Me
'End Sub

'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal ParamArray args() As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal provider As System.IFormatProvider, ByVal format As String, ByVal ParamArray args() As Object) As System.Text.StringBuilder
Public Function AppendFormat(ByVal sFormat As String, ParamArray args()) As StringBuilder
  Set AppendFormat = Me
End Function

Public Property Get Capacity() As Long
  Capacity = mCapacity
End Property
Public Property Let Capacity(LngVal As Long)
  mCapacity = LngVal
End Property

'Public Function EnsureCapacity(ByVal capacity As Integer) As Integer
Public Function EnsureCapacity(ByVal Capacity As Long) As Long 'cInteger
Dim n As Long
  n = Len(mStrValue)
  'Set EnsureCapacity = New cInteger
  If n >= Capacity Then
    EnsureCapacity = n
  Else
    mStrValue = mStrValue & Space$(Capacity - n)
    EnsureCapacity = Capacity
  End If
End Function

'Public Function Equals(ByVal sb As System.Text.StringBuilder) As Boolean
Public Function Equals(ByVal sb As StringBuilder) As Boolean
  'Set Equals = New cBoolean
  If sb.ToStr = mStrValue Then Equals = True
End Function

'Public Function Insert(ByVal index As Integer, ByVal value As Boolean) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Byte) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Char) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Decimal) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Double) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Integer) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Long) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Object) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Short) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Single) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.SByte) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt16) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt32) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt64) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As String) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As String, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value() As Char) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value() As Char, ByVal startIndex As Integer, ByVal charCount As Integer) As System.Text.StringBuilder
Public Function Insert(ByVal index As Long, ByVal value As Variant, Optional ByVal startIndex As Long = 1, Optional ByVal charCount As Long = 1) As StringBuilder
'Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
Dim lLen As Long
Dim lPos As Long
Dim lSize As Long
Dim sValue As String
  ' is iIndex within bounds?
  If (index * 2 > mLength) Then
    Err.Raise 9
  Else
    sValue = CStr(value)
    lLen = 2 * Len(sValue)
    If (mLength + lLen) > mCapacity Then
      mStrValue = mStrValue & Space$(mCapacity \ 2)
      mCapacity = mCapacity + mCapacity
    End If
      
    ' Move existing characters from current position
    lPos = UnsignedAdd(StrPtr(mStrValue), index * 2)
    lSize = mLength - index * 2
    
    ' moving from iIndex to iIndex + lLen
    CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
    
    ' Insert new characters:
    CopyMemory ByVal lPos, ByVal StrPtr(sValue), lLen
    
    mLength = mLength + lLen
  End If
  Set Insert = Me
End Function

'Public Function Remove(ByVal startIndex As Integer, ByVal length As Integer) As System.Text.StringBuilder
Public Function Remove(ByVal startIndex As Long, ByVal Length As Long) As StringBuilder
'Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
Dim lSrc As Long
Dim lDst As Long
Dim lSize As Long

   ' is iIndex within bounds?
   If (startIndex * 2 > mLength) Then
      'Err.Raise 9
   Else
      ' is there sufficient length?
      If ((startIndex + Length) * 2 > mLength) Then
         'Err.Raise 9
      Else
         ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
         lSrc = UnsignedAdd(StrPtr(mStrValue), (startIndex + Length) * 2)
         lDst = UnsignedAdd(StrPtr(mStrValue), startIndex * 2)
         lSize = (mLength - (startIndex + Length) * 2)
         CopyMemory ByVal lDst, ByVal lSrc, lSize
         mLength = mLength - Length * 2
      End If
   End If
   Set Remove = Me
'End Sub
End Function

Private Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
Dim lInstr As Long
  'Set Find = New cInteger
  If (lStartIndex > 0) Then
    lInstr = InStr(lStartIndex, mStrValue, sToFind, compare)
  Else
    lInstr = InStr(mStrValue, sToFind, compare)
  End If
  If (lInstr < mLength \ 2) Then
    Find = lInstr
  End If
End Function

'Public Function Replace(ByVal oldChar As Char, ByVal newChar As Char) As System.Text.StringBuilder
'Public Function Replace(ByVal oldChar As Char, ByVal newChar As Char, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Replace(ByVal oldValue As String, ByVal newValue As String) As System.Text.StringBuilder
'Public Function Replace(ByVal oldValue As String, ByVal newValue As String, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
Public Function Replace(ByVal oldValue As String, ByVal newValue As String, Optional ByVal startIndex As Long, Optional ByVal Count As Long) As StringBuilder
Dim PartStrVal As String, n As Long, StrPos As Long
Dim LStrVal As String, RStrVal As String
  n = Len(mStrValue)
  If startIndex > 0 And Count > 0 Then
    If startIndex < n Then
      If startIndex + Count > n Then Count = n - startIndex
      PartStrVal = Mid(mStrValue, startIndex, Count)
    End If
  Else
    PartStrVal = mStrValue
  End If
  StrPos = InStr(1, PartStrVal, oldValue, vbBinaryCompare)
  Do While StrPos > 0
    StrPos = InStr(StrPos, PartStrVal, oldValue, vbBinaryCompare)
  Loop
  Set Replace = Me
End Function


'Public Overrides Function ToStr() As String
'Public Function ToStr(ByVal startIndex As Integer, ByVal length As Integer) As String
Public Function ToStr(Optional ByVal startIndex As Long, Optional ByVal Length As Long) As String
  'Set ToStr = New cString
     ' The internal string:
  If Len(mStrValue) > 0 Then
    If startIndex > 1 Then
      If Length > 0 Then
        ToStr = VBA.Mid$(mStrValue, startIndex, Length)
      Else 'die Länge wurde nicht angegeben, also die gesamte verbleibende Länge
        ToStr = VBA.Mid$(mStrValue, startIndex, Length - startIndex)
      End If
    Else
      ToStr = VBA.Left$(mStrValue, mLength \ 2) '\2 durch 2, da mLength in Byte
    End If
  End If
End Function

'wird in .NET-Klasse nicht öffentlich verwendet
'Public Sub HeapMinimize()
'Dim iLen As Long
'
'   ' Reduce the string size so only the minimal chunks
'   ' are allocated:
'   If (m_iLen - m_iPos) > m_iChunkSize Then
'      iLen = m_iLen
'      Do While (iLen - m_iPos) > m_iChunkSize
'         iLen = iLen - m_iChunkSize
'      Loop
'      m_sString = Left$(m_sString, iLen \ 2)
'      m_iLen = iLen
'   End If
'
'End Sub
Private Function UnsignedAdd(start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If start And &H80000000 Then 'Start < 0
      UnsignedAdd = start + Incr
   ElseIf (start Or &H80000000) < -Incr Then
      UnsignedAdd = start + Incr
   Else
      UnsignedAdd = (start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

'InternalEnsureCapacity (capacity);
'InternalEnsureCapacity(value);
'InternalEnsureCapacity (replace.Length + (_length - count))
'InternalEnsureCapacity (needed_cap);

'_str.InternalSetChar (index, value);
'_str.InternalSetChar (replaceIterate, newChar);
'_str.InternalSetChar(_length, value);

'_str.InternalSetLength(_length);

'String.InternalStrcpy (_str, startIndex, _str, startIndex + length, _length - (startIndex + length))
'String.InternalStrcpy (_str, startIndex, replace);
'String.InternalStrcpy (_str, startIndex + replace.Length, end);
'String.InternalStrcpy (_str, _length, value);
'String.InternalStrcpy (_str, _length, value);
'String.InternalStrcpy (_str, _length, value, startIndex, charCount);
'internal extern static void
'InternalStrcpy (String dest, int destPos, String src);
'InternalStrcpy (String dest, int destPos, char[] chars);
'InternalStrcpy (String dest, int destPos, String src, int sPos, int count);
'InternalStrcpy (String dest, int destPos, char[] chars, int sPos, int count);
'
Private Sub InternalStrcpy(dest As String, destPos As Long, src As String, Optional sPos As Long, Optional Count As Long)
'kopiert den String src an die Stelle destPos im string dest
'kopiert den Teilstring src = mid$(src, spos, count) an die Stelle destPos im String dest
Dim lLen As Long
Dim PartStr As String
  If Count = 0 Then Count = Len(src)
  If sPos = 0 Then sPos = 1
  PartStr = Mid$(src, sPos, Count)
  lLen = 2 * Len(PartStr)
' CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
  
  CopyMemory ByVal UnsignedAdd(StrPtr(dest), destPos), ByVal StrPtr(PartStr), lLen
End Sub
'string tmp = String.InternalAllocateStr (capacity);
Private Function InternalAllocateStr(mCap As Long) As String
  'sollte man besser Space$ oder String$(Length, vbNullString) verwenden??
  InternalAllocateStr = Space$(mCap)
End Function

Private Sub InternalEnsureCapacity(ByVal size As Long)
Dim tmp As String
  If (size > mLength Or mCachedStr = mStrValue) Then
    mCapacity = mLength

        '// Try double buffer, if that doesn't work, set the length as capacity
    If (size > mCapacity) Then
          
          '// The first time a string is appended, we just set _cached_str
          '// and _str to it. This allows us to do some optimizations.
          '// Below, we take this into account.
          If (mCachedStr = mStrValue) And (mCapacity < constDefaultCapacity) Then
            mCapacity = constDefaultCapacity
          End If
          
          'capacity = capacity << 1; ShiftLeft????
          If (size > mCapacity) Then mCapacity = size
          If (mCapacity >= 2147483647 Or mCapacity < 0) Then
            mCapacity = 2147483647 'Int32.MaxValue;
          End If

          If (mCapacity > mMaxCapacity And size <= mMaxCapacity) Then
            mCapacity = mMaxCapacity
          End If
          
          If (mCapacity > mMaxCapacity) Then
          '  throw new ArgumentOutOfRangeException ("size", "capacity was less than the current size.");
          End If
    End If

    tmp = InternalAllocateStr(mCapacity)
    If (mLength > 0) Then
      Call InternalStrcpy(tmp, 0, mStrValue, 0, mLength)
    End If
    mStrValue = tmp
  End If

  mCachedStr = vbNullString
End Sub

Public Property Let Length(LngVal As Long)
    mLength = LngVal
    If mLength = 0 Then
        New_
    End If
End Property
Public Property Get Length() As Long
    Length = mLength
End Property

Private Sub Class_Terminate()
    Length = 0
End Sub
