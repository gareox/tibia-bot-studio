Attribute VB_Name = "modTypes"
Option Explicit

Type ITEM_TYPE
    Name As String
    TypeSourcePos As Long
End Type

Public Types() As ITEM_TYPE
Public TypeDec As String

Sub InitTypes()
    TypeDec = ""
    ReDim Types(0) As ITEM_TYPE
End Sub

Sub AddType(Name As String, TypeSourcePos As Long)
    ReDim Preserve Types(UBound(Types) + 1) As ITEM_TYPE
    Types(UBound(Types)).Name = Name
    Types(UBound(Types)).TypeSourcePos = TypeSourcePos
End Sub

Sub Assign_Type(Ident As String)
    Dim OldPos As Long
    
    TypeDec = Identifier
    AddSymbol TypeDec, 512 + UBound(DataSection), Data, IS_TYPE
    
    OldPos = SourcePos
    SourcePos = GetTypeSourcePos(Ident)
    VariableBlock Identifier
    Symbol "}"
    
    SourcePos = OldPos
    TypeDec = ""
    Terminator
    CodeBlock
End Sub

Function GetTypeSourcePos(Ident As String) As String
    Dim i As Integer
    For i = 1 To UBound(Types)
        If LCase(Types(i).Name) = LCase(Ident) Then
            GetTypeSourcePos = Types(i).TypeSourcePos
            Exit Function
        End If
    Next i
End Function

Function IsAssignedType(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Ident) Then
            If Symbols(i).SymType = IS_TYPE Then
                IsAssignedType = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsType(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Types)
        If LCase(Types(i).Name) = LCase(Ident) Then
            IsType = True
            Exit Function
        End If
    Next i
End Function
