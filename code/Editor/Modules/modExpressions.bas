Attribute VB_Name = "modExpressions"
Option Explicit
Public Relation As String
Public eUniqueID As Long
Public IsStringCompare As Boolean

Sub ChooseRelation(iID As Long, LabelThen As String, LabelElse As String)
    If Relation = "=" Then
        Expr_JumpEqual LabelThen & iID
        Expr_JumpNotEqual LabelElse & iID
    ElseIf Relation = "!" Then
        Expr_JumpEqual LabelElse & iID
        Expr_JumpNotEqual LabelThen & iID
    ElseIf Relation = "<>" Then
        Expr_JumpEqual LabelElse & iID
        Expr_JumpNotEqual LabelThen & iID
    ElseIf Relation = "<" Then
        Expr_JumpBelow LabelThen & iID
        Expr_JumpAboveEqual LabelElse & iID
    ElseIf Relation = ">" Then
        Expr_JumpAbove LabelThen & iID
        Expr_JumpBelowEqual LabelElse & iID
    ElseIf Relation = ">=" Then
        Expr_JumpAboveEqual LabelThen & iID
        Expr_JumpBelow LabelElse & iID
    ElseIf Relation = "<=" Then
        Expr_JumpBelowEqual LabelThen & iID
        Expr_JumpAbove LabelElse & iID
    End If
End Sub

Sub Expression(StoreAt As String)

    Dim Term As String
    Dim Value As Long
    Dim Ident As String
    Dim VariableName As String
        
    While Mid(Source, SourcePos, 1) <> ")"
        SkipBlank
        If IsSymbol(",") Then Exit Sub
        If IsSymbol("+") Then
            Symbol "+": Term = "+"
        ElseIf IsSymbol("-") Then
            Symbol "-": Term = "-"
        ElseIf IsSymbol("*") Then
            Symbol "*": Term = "*"
        ElseIf IsSymbol("/") Then
            Symbol "/": Term = "/"
        ElseIf IsNumberExpression Then
            Value = NumberExpression
            If Term = "" Then Expr_Set StoreAt, Value
            If Term = "+" Then Expr_Add StoreAt, Value
            If Term = "-" Then Expr_Sub StoreAt, Value
            If Term = "*" Then Expr_Mul StoreAt, Value
            If Term = "/" Then Expr_Div StoreAt, Value
        ElseIf IsStringExpression Then
            dUniqueID = dUniqueID + 1
            Dim ImportID As Long: ImportID = GetImportIDByName("lstrcpyA")
            Dim StringValue As String: StringValue = StringExpression
            DeclareDataString "UniqueString" & dUniqueID, StringValue, Len(StringValue)
        
            If IsCallFrame = True Then GoTo IsFrameCall
            
            If GetSymbolType(StoreAt) = IS_DWORD Or GetSymbolType(StoreAt) = IS_WORD Or GetSymbolType(StoreAt) = IS_BYTE Then
                AddCodeByte &HC7, &H5
                AddCodeFixup StoreAt
                AddCodeFixup "UniqueString" & dUniqueID
            Else
IsFrameCall:
                sUniqueID = sUniqueID + 1
                DeclareDataDWord "AddressToString" & sUniqueID, &H0
                AddFixup "UniqueString" & dUniqueID, 512 + UBound(DataSection), &H400000, Data
                Push
                AddFixup "UniqueString" & dUniqueID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
                Push
                AddFixup StoreAt, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
                Invoke
                AddFixup "ImageImportByName" & ImportID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Import
            End If
        ElseIf IsSymbol("=") Then
            Symbol "=": Relation = "="
            Exit Sub
        ElseIf IsSymbol("!") Then
            Symbol "!": Relation = "!"
            Exit Sub
        ElseIf IsSymbol("<>") Then
            Symbol "<": Symbol ">": Relation = "<>"
            Exit Sub
        ElseIf IsSymbol("<") Then
            Symbol "<": Relation = "<"
            If IsSymbol("=") Then Symbol "=": Relation = Relation & "="
            Exit Sub
        ElseIf IsSymbol(">") Then
            Symbol ">": Relation = ">"
            If IsSymbol("=") Then Symbol "=": Relation = Relation & "="
            Exit Sub
        ElseIf IsSymbol(";") Then
            Exit Sub
        ElseIf IsSymbol(")") Then
            Exit Sub
        Else
            Ident = Identifier
            If IsAlias(Ident) Then
                Call_Alias Ident, True
                Expr_StoreEAX "Internal.Return"
                VariableName = "Internal.Return"
                GoTo IsVariableLabel
            ElseIf IsLocalVariable(Ident) Then
                Dim Label As String
                Dim iLabel As Long
                Label = CurrentFrame & "." & Ident
                AddCodeWord &H858D
                For iLabel = 1 To UBound(Symbols)
                    If Symbols(iLabel).Name = Label Then
                        AddCodeDWord Symbols(iLabel).Offset
                    End If
                Next iLabel
                AddCodeWord &H8B
                If Term = "" Then
                    Expr_StoreEAX StoreAt
                ElseIf Term = "+" Then
                    Expr_StoreEAX "Internal.Local"
                    Expr_Add_Var StoreAt, "Internal.Local"
                ElseIf Term = "-" Then
                    Expr_StoreEAX "Internal.Local"
                    Expr_Sub_Var StoreAt, "Internal.Local"
                ElseIf Term = "*" Then
                    Expr_StoreEAX "Internal.Local"
                    Expr_Mul_Var StoreAt, "Internal.Local"
                ElseIf Term = "/" Then
                    Expr_StoreEAX "Internal.Local"
                    Expr_Div_Var StoreAt, "Internal.Local"
                End If
            ElseIf IsVariable(Ident) Then
                VariableName = Ident
IsVariableLabel:
                If GetSymbolType(VariableName) = IS_STRING Then
                    sUniqueID = sUniqueID + 1
                    DeclareDataDWord "AddressToString" & sUniqueID, &H0
                    AddFixup VariableName, 512 + UBound(DataSection), &H400000, Data
                    IsStringCompare = True
                Else
                    If Term = "" Then Expr_Set_Var StoreAt, VariableName
                    If Term = "+" Then Expr_Add_Var StoreAt, VariableName
                    If Term = "-" Then Expr_Sub_Var StoreAt, VariableName
                    If Term = "*" Then Expr_Mul_Var StoreAt, VariableName
                    If Term = "/" Then Expr_Div_Var StoreAt, VariableName
                End If
            ElseIf IsAssignedType(Ident) Then
                If Term = "" Then Expr_Set_Var StoreAt, Ident
                If Term = "+" Then Expr_Add_Var StoreAt, Ident
                If Term = "-" Then Expr_Sub_Var StoreAt, Ident
                If Term = "*" Then Expr_Mul_Var StoreAt, Ident
                If Term = "/" Then Expr_Div_Var StoreAt, Ident
            ElseIf IsSymbol("@") Then
                Symbol "@"
                Expr_MovEAXAdress Identifier
                'mov [CopyTo], eax
                AddCodeByte &HA3
                AddCodeDWord &H0
                AddFixup StoreAt, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
            ElseIf IsFrame(Ident) Then
                Call_Frame Ident, True
                Expr_StoreEAX StoreAt
            ElseIf Ident = "" Then
                ErrMessage "unknown symbol '" & Mid(Source, SourcePos, 1) & "'"
            Else
                ErrMessage "unknown symbol '" & Mid(Source, SourcePos, 1) & "'"
                Exit Sub
            End If
        End If
        
    Wend
End Sub


