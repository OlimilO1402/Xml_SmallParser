Attribute VB_Name = "MShift"
Option Explicit
'#####################  v  for Bit Shifting v   ####################
' by Paul - wpsjr1@syix.com
' http://www.syix.com/wpsjr1/index.html

' Author's comments:  use ShiftLeft04 or ShiftRightZ05 without the wrappers if you need more speed,
' they're 25% faster than these.

' NOTE: YOU *MUST* CALL InitFunctionsShift() BEFORE USING THESE FUNCTIONS

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const SHLCode As String = "8A4C240833C0F6C1E075068B442404D3E0C20800"  ' shl eax, cl = D3 E0
Private Const SHRCode As String = "8A4C240833C0F6C1E075068B442404D3E8C20800"  ' shr eax, cl = D3 E8
Private Const SARCode As String = "8A4C240833C0F6C1E075068B442404D3F8C20800"  ' sar eax, cl = D3 F8
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Dim bHoldSHL() As Byte
Dim bHoldSHR() As Byte
Dim bHoldSAR() As Byte
Dim lCompiled As Long
'#####################  ^  for Bit Shifting ^   ####################

'#####################  v  for Bit Shifting v   ####################
Public Sub InitFunctionsShift() ' call this in your Sub Main or Form_Load
    If Compiled Then
        SubstituteCode bHoldSHL, SHLCode, AddressOf ShiftLeft
        SubstituteCode bHoldSHR, SHRCode, AddressOf ShiftRightZ
        SubstituteCode bHoldSAR, SARCode, AddressOf ShiftRight
    End If
End Sub

' this is the Murphy McCauley method which I modified slightly, http://www.fullspectrum.com/deeth/
Private Sub SubstituteCode(StoreHere() As Byte, CodeString As String, ByVal AddressOfFunctionToReplace As Long)
    Dim OldProtection As Long
    Dim s As String
    Dim i As Long
      
    ReDim StoreHere(Len(CodeString) \ 2 - 1)
    
    For i = 0 To Len(CodeString) \ 2 - 1
        StoreHere(i) = Val("&H" & Mid$(CodeString, i * 2 + 1, 2))
    Next
    
    VirtualProtect ByVal AddressOfFunctionToReplace, 21, PAGE_EXECUTE_READWRITE, OldProtection
    RtlMoveMemory ByVal AddressOfFunctionToReplace, &H90, 1 ' nop to insure our first line is not concated with the previous instruction
    RtlMoveMemory ByVal AddressOfFunctionToReplace + 1, StoreHere(0), 20 ' shr/shl code substitution
    VirtualProtect ByVal AddressOfFunctionToReplace, 21, OldProtection, OldProtection
    
    ' alternately, if the code is much longer use this instead:
    
    ' VirtualProtect ByVal AddressOfFunctionToReplace, 7, PAGE_EXECUTE_READWRITE, OldProtection
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace, &HB8, 1  ' mov eax, PointerToCode
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace + 1, Varptr(StoreHere(0)),4
    ' RtlMoveMemory ByVal AddressOfFunctionToReplace + 5, &HE0FF&, 2 ' jmp eax
    ' VirtualProtect ByVal AddressOfFunctionToReplace, 7, OldProtection, OldProtection
End Sub

' Leave these placeholder functions, and their code
Public Function ShiftLeft(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Dim mask As Long
    
    Select Case ShiftCount
    Case 1 To 31
        ' mask out bits that are pushed over the edge anyway
        mask = Pow2(31 - ShiftCount)
        ShiftLeft = value And (mask - 1)
        ' shift
        ShiftLeft = ShiftLeft * Pow2(ShiftCount)
        ' set sign bit
        If value And mask Then
          ShiftLeft = ShiftLeft Or &H80000000
        End If
    Case 0
        ' ret unchanged
        ShiftLeft = value
    End Select
End Function

Public Function ShiftRightZ(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Select Case ShiftCount
    Case 1 To 31
        If value And &H80000000 Then
            ShiftRightZ = (value And Not &H80000000) \ 2
            ShiftRightZ = ShiftRightZ Or &H40000000
            ShiftRightZ = ShiftRightZ \ Pow2(ShiftCount - 1)
        Else
            ShiftRightZ = value \ Pow2(ShiftCount)
        End If
    Case 0
        ' ret unchanged
        ShiftRightZ = value
    End Select
End Function

Public Static Function ShiftRight(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20011009
    Dim lPow2(0 To 30) As Long
    Dim i As Long
    
    Select Case ShiftCount
    Case 0:      ShiftRight = value
    Case 1 To 30
        If i = 0 Then
            lPow2(0) = 1
            For i = 1 To 30
                lPow2(i) = 2 * lPow2(i - 1)
            Next
        End If
        If value And &H80000000 Then
            ShiftRight = value \ lPow2(ShiftCount)
            If ShiftRight * lPow2(ShiftCount) <> value Then
                ShiftRight = ShiftRight - 1
            End If
        Else
            ShiftRight = value \ lPow2(ShiftCount)
        End If
    Case 31
        If value And &H80000000 Then
            ShiftRight = -1
        Else
            ShiftRight = 0
        End If
    End Select
End Function
 
Public Static Function Pow2(ByVal Exponent As Long) As Long
    ' by Donald, donald@xbeat.net, 20001217
    ' * Power205
    Dim alPow2(0 To 31) As Long
    Dim i As Long
    
    Select Case Exponent
    Case 0 To 31
        ' initialize lookup table
        If alPow2(0) = 0 Then
            alPow2(0) = 1
            For i = 1 To 30
                alPow2(i) = alPow2(i - 1) * 2
            Next
            alPow2(31) = &H80000000
        End If
        ' return
        Pow2 = alPow2(Exponent)
    End Select
End Function

Private Function Compiled() As Long
    On Error Resume Next
    Debug.Print 1 \ 0
    Compiled = (Err.Number = 0)
End Function


