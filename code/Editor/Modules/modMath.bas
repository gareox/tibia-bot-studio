Attribute VB_Name = "modMath"
'*************************************************
'*  modMath most functions from Tommy Lillehagen *
'*************************************************

Option Explicit

Function HiByte(ByVal iWord As Integer) As Byte
    HiByte = (iWord And &HFF00&) \ &H100
End Function

Function LoByte(ByVal iWord As Integer) As Byte
    LoByte = iWord And &HFF
End Function

Function HiWord(lDword As Long) As Integer
    HiWord = (lDword And &HFFFF0000) \ &H10000
End Function

Function LoWord(lDword As Long) As Integer
    If lDword And &H8000& Then
        LoWord = lDword Or &HFFFF0000
    Else
        LoWord = lDword And &HFFFF&
    End If
End Function

Function IsByte(Value As Long) As Boolean
    IsByte = (Value >= -128) And (Value <= 127)
End Function

Function IsWord(Value As Long) As Boolean
    IsWord = (Value >= -32768) And (Value <= 32767)
End Function

Function FixByte(Offset As Long, Value As Long)
    LinkCode(Offset) = HiByte(Value)
End Function

Function FixWord(Offset As Long, Value As Long)
    LinkCode(Offset) = LoByte(LoWord(Value))
    LinkCode(Offset + 1) = HiByte(HiWord(Value))
End Function

Function FixDWord(Offset As Long, Value As Long)
    LinkCode(Offset) = LoByte(LoWord(Value))
    LinkCode(Offset + 1) = HiByte(LoWord(Value))
    LinkCode(Offset + 2) = LoByte(HiWord(Value))
    LinkCode(Offset + 3) = HiByte(HiWord(Value))
End Function

