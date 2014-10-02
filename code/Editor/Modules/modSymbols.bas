Attribute VB_Name = "modSymbols"
Option Explicit

Enum SYMBOL_TYPE
    IS_FRAME = 0
    IS_BYTE = 1
    IS_WORD = 2
    IS_DWORD = 3
    IS_ALIAS = 4
    IS_STRING = 5
    IS_LABEL = 6
    IS_TYPE = 7
    IS_LOCAL_DWORD = 8
    IS_LOCAL_STRING = 9
End Enum

Type ITEM_SYMBOL
    Name As String
    Offset As Long
    Section As SECTION_SIZE
    SymType As SYMBOL_TYPE
    Space As Long
End Type

Public Symbols() As ITEM_SYMBOL

Sub InitSymbols()
    ReDim Symbols(0) As ITEM_SYMBOL
End Sub

Sub AddSymbol(Name As String, Offset As Long, Section As SECTION_SIZE, SymType As SYMBOL_TYPE, Optional Space As Long)
    If SymbolExists(Name) Then ErrMessage "symbol '" & Name & "' already exists.": Exit Sub
    ReDim Preserve Symbols(UBound(Symbols) + 1) As ITEM_SYMBOL
    Symbols(UBound(Symbols)).Name = Name
    Symbols(UBound(Symbols)).Offset = Offset
    Symbols(UBound(Symbols)).Section = Section
    Symbols(UBound(Symbols)).SymType = SymType
    Symbols(UBound(Symbols)).Space = Space
End Sub

Sub AddConstant(Name As String, Value As String)
    If IsNumeric(Value) Then
        DeclareDataDWord Name, CDbl(Value)
    Else
        DeclareDataString Name, Value, Len(Value)
    End If
End Sub

Function SymbolExists(Name As String) As Boolean
    Dim i As Integer
    If InStr(1, LCase(Name), "unique", vbTextCompare) <> 0 Then Exit Function
    
    For i = 1 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Name) Then
            SymbolExists = True
            Exit Function
        End If
    Next i
End Function

Function GetSymbolSpace(Ident As String) As Long
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Ident) Then
            GetSymbolSpace = Symbols(i).Space
            Exit Function
        End If
    Next i
End Function

Function GetSymbolType(Ident As String) As SYMBOL_TYPE
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Ident) Then
            GetSymbolType = Symbols(i).SymType
            Exit Function
        End If
    Next i
End Function
