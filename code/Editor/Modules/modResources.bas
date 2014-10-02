Attribute VB_Name = "modResources"
Option Explicit

Public dUniqueID As Long
Public sUniqueID As Long
Public DataSection() As Byte

Sub InitData()
    dUniqueID = 0
    sUniqueID = 0
    eUniqueID = 0
    UniqueID = 0
    ReDim DataSection(0) As Byte
End Sub

Sub DeclareDataByte(Name As String, Value As Byte)
    AddSymbol Name, 512 + UBound(DataSection), Data, IS_BYTE
    AddDataByte Value
End Sub

Sub DeclareDataWord(Name As String, Value As Integer)
    AddSymbol Name, 512 + UBound(DataSection), Data, IS_WORD
    AddDataWord Value
    AddDataByte &H0
End Sub

Sub DeclareDataDWord(Name As String, Value As Long)
    AddSymbol Name, 512 + UBound(DataSection), Data, IS_DWORD
    AddDataDWord Value
End Sub

Sub ReserveSpace(BytesReserve As Long)
    Dim i As Integer
    If BytesReserve = 0 Then Exit Sub
    For i = 0 To BytesReserve
        AddDataByte &H0
    Next i
End Sub

Sub DeclareDataString(Name As String, Value As String, Optional Space As Long)
    Dim i As Integer

    AddSymbol Name, 512 + UBound(DataSection), Data, IS_STRING, Space
    
    For i = 1 To Len(Value)
        AddDataByte Asc(Mid(Value, i, 1))
    Next i
    
    If Space > 0 Then
        For i = Len(Value) To Space
            AddDataByte &H0
        Next i
    End If
    
    AddDataByte &H0
End Sub

Function GetVariableType(Name As String) As SYMBOL_TYPE
    Dim i As Integer
    For i = 0 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Name) Then
            GetVariableType = Symbols(i).SymType
        End If
    Next i
End Function

Sub AddDataDWord(Value As Long)
    AddDataWord LoWord(Value)
    AddDataWord HiWord(Value)
End Sub

Sub AddDataWord(Value As Integer)
    AddDataByte LoByte(Value)
    AddDataByte HiByte(Value)
End Sub

Sub AddDataByte(Value As Byte)
    ReDim Preserve DataSection(UBound(DataSection) + 1) As Byte
    DataSection(UBound(DataSection)) = Value
End Sub

