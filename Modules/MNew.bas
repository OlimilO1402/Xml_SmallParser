Attribute VB_Name = "MNew"
Option Explicit

Public Function StringBuilder(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
  Set StringBuilder = New StringBuilder: StringBuilder.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

Public Function FileStream(aPath As String, Optional FMode As FileMode = FileMode_Random, Optional FAccess As FileAccess = FileAccess_ReadWrite, Optional FShare As FileShare = FileShare_None) As FileStream
  Set FileStream = New FileStream: FileStream.New_ aPath, FMode, FAccess, FShare
End Function

Public Function StreamReader(ByVal aStream As Stream, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long) As StreamReader
  Set StreamReader = New StreamReader: StreamReader.New_ aStream, Encoding, detectEncodingFromByteOrderMarks, bufferSize
End Function

Public Function StreamReaderP(ByVal aPath As String, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long) As StreamReader
  Set StreamReaderP = New StreamReader: StreamReaderP.NewP aPath, Encoding, detectEncodingFromByteOrderMarks, bufferSize
End Function

