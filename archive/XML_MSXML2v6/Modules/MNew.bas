Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function


