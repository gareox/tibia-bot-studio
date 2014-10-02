Attribute VB_Name = "modFrames"
Option Explicit

Type ITEM_FRAME
    Name As String
    Declares As String
End Type

Public IsCallFrame As Boolean
Public Frames() As ITEM_FRAME
Public ArgCount As Long
Public CurrentFrame As String

Sub InitFrames()
    ReDim Frames(0) As ITEM_FRAME
End Sub

Sub AddFrameDeclare(VarName As String)
    Frames(UBound(Frames)).Declares = Frames(UBound(Frames)).Declares & VarName & ","
End Sub

Sub AddFrame(Name As String)
    ReDim Preserve Frames(UBound(Frames) + 1) As ITEM_FRAME
    Frames(UBound(Frames)).Name = Name
    CurrentFrame = Name
End Sub

Function IsLocalVariable(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(CurrentFrame & "." & Ident) Then
            IsLocalVariable = True
            Exit Function
        End If
    Next i
End Function

Function GetFrameIDByName(Name As String) As Long
    Dim i As Integer
    For i = 1 To UBound(Frames)
        If LCase(Frames(i).Name) = LCase(Name) Then
            GetFrameIDByName = i
            Exit Function
        End If
    Next i
End Function

Sub Call_Frame(Ident As String, Optional FromExpression As Boolean)
    Dim i As Integer
    Dim fID As Long
    Dim iLabel As Long
    Dim FrameDeclares As Variant
    
   IsCallFrame = True
    fID = GetFrameIDByName(Ident)
   
    FrameDeclares = Split(Frames(fID).Declares, ",")
    ReverseParams UBound(FrameDeclares)
    
    Symbol "("
    For i = UBound(FrameDeclares) - 1 To 0 Step -1
        If GetSymbolType(Frames(fID).Name & "." & FrameDeclares(i)) = IS_LOCAL_DWORD Then
            eUniqueID = eUniqueID + 1
            DeclareDataDWord "UniqueExpression" & eUniqueID, &H0
            Expression "UniqueExpression" & eUniqueID
            AddCodeByte &HA1
            AddCodeFixup "UniqueExpression" & eUniqueID
            PushEAX
        ElseIf GetSymbolType(Frames(fID).Name & "." & FrameDeclares(i)) = IS_LOCAL_STRING Then
            eUniqueID = eUniqueID + 1
            DeclareDataDWord "UniqueExpression" & eUniqueID, &H0
            Expression "UniqueExpression" & eUniqueID
            AddCodeByte &HB8
            AddCodeFixup "UniqueExpression" & eUniqueID
            PushEAX
        Else
            Expression Frames(fID).Name & "." & FrameDeclares(i)
        End If
        If IsSymbol(",") Then Skip
    Next i
    Symbol ")"
    If Not FromExpression Then Terminator
    Expr_Call Ident
    If Not FromExpression Then CodeBlock
    IsCallFrame = False
End Sub

Sub Statement_Return()
    Symbol "("
    Expression "Internal.Return"
    Expr_MovEAX "Internal.Return"
    PushEAX
    Expr_Jump CurrentFrame & ".end"
    Symbol ")"
    Terminator
    CodeBlock
End Sub

Sub Declare_Frame()
    Dim pCount As Long
    Dim Ident As String
    Dim Name As String
    Dim IdentII As String
    
    ArgCount = 0
    Ident = Identifier
    
    AddFrame Ident
    
    Symbol "("
    AddCodeDWord &H0
NextDeclare:
    IdentII = Identifier
    If IdentII <> "" Then
        VariableBlock IdentII, True
        ArgCount = ArgCount + 1
    End If
    
    If IsSymbol(",") Then
        Symbol (","): GoTo NextDeclare
    ElseIf IsSymbol(")") Then
        Symbol (")"): Terminator
    Else
        ErrMessage "unexpected '" & Mid(Source, SourcePos, 1) & "'"
    End If

    AddSymbol Ident, UBound(CodeSection), 0, IS_FRAME
    AddSymbol Ident & ".Address", &H1000 + 512 + UBound(CodeSection), Code, IS_DWORD
    StartFrame
    CodeBlock
    AddSymbol Ident & ".end", UBound(CodeSection), 0, IS_LABEL
    EndProc
    EndFrame ArgCount * 4
    CodeBlock
End Sub
