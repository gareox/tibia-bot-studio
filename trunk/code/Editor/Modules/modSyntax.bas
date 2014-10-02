Attribute VB_Name = "modSyntax"
Option Explicit
Public UniqueID As Long

Sub InitApplication()
    Dim Ident As String
    Dim ApplicationName As String
    Ident = Identifier
    If LCase(Ident) = "init" Then
        SkipBlank
        ApplicationName = StringExpression
        Symbol ","
            Ident = Identifier
            If LCase(Ident) = "gui" Then
                AppType = GUI
            ElseIf LCase(Ident) = "console" Then
                AppType = Console
            Else
                ErrMessage "expected 'GUI' or 'Console' but found '" & Ident & "'"
            End If
        Terminator
    Else
        ErrMessage "expected 'init' but found '" & Ident & "'"
    End If
End Sub

Sub Statement_Include()
    Dim Content As String
    Dim Ident As String
    Ident = StringExpression
    Terminator
    
    If Dir(App.Path & "\include\" & Ident) = "" Then ErrMessage "cannot include '" & Ident & "'. check your include folder.": Exit Sub
    Open App.Path & "\include\" & Ident For Binary As #1
        Content = Space(LOF(1))
        Get #1, , Content
    Close #1
    
    Dim Header As String
    Dim Footer As String
    Header = Mid(Source, 1, SourcePos)
    Footer = Mid(Source, SourcePos, Len(Source) - SourcePos)
    Source = Header & Content & Footer
    MsgBox Source
    CodeBlock
End Sub

Sub Statement_IA()
    Symbol "{"
    IABlock
    Symbol "}"
    CodeBlock
End Sub

Sub Statement_If()
    Dim iID As Long
    iID = iID + UniqueID: UniqueID = UniqueID + 1
    
    Symbol "("
    
    IsStringCompare = False
    
    Expression "Internal.Compare.I"
    Expression "Internal.Compare.II"
    
    If IsStringCompare Then
        Expr_StringCompare "AddressToString" & sUniqueID - 1, "AddressToString" & sUniqueID
    Else
        Expr_Compare "Internal.Compare.I", "Internal.Compare.II"
    End If
    
    ChooseRelation iID, "then", "else"
    Symbol ")"
    Symbol "{"
    AddSymbol "then" & iID, UBound(CodeSection), 0, IS_LABEL
    CodeBlock
    Expr_Jump "out" & iID
    Symbol "}"
    SkipBlank
    If IsIdent("else") Then
        Call Identifier
        Symbol "{"
        AddSymbol "else" & iID, UBound(CodeSection), 0, IS_LABEL
        CodeBlock
        Symbol "}"
    Else
        AddSymbol "else" & iID, UBound(CodeSection), 0, IS_LABEL
    End If
        AddSymbol "out" & iID, UBound(CodeSection), 0, IS_LABEL
    CodeBlock
End Sub

Sub Statement_While()
    Dim wID As Long
    wID = wID + UniqueID: UniqueID = UniqueID + 1
    Symbol "("
    
    AddSymbol "swhile" & wID, UBound(CodeSection), 0, IS_LABEL
    IsStringCompare = False
   
    Expression "Internal.Compare.I"
    Expression "Internal.Compare.II"
    
    If IsStringCompare Then
        Expr_StringCompare "AddressToString" & sUniqueID - 1, "AddressToString" & sUniqueID
    Else
        Expr_Compare "Internal.Compare.I", "Internal.Compare.II"
    End If
    
    ChooseRelation wID, "while", "endwhile"
    
    Symbol ")"
    Symbol "{"
    
    AddSymbol "while" & wID, UBound(CodeSection), 0, IS_LABEL
    CodeBlock
    Expr_Jump "swhile" & wID
    Symbol "}"
    SkipBlank
    
    AddSymbol "endwhile" & wID, UBound(CodeSection), 0, IS_LABEL
    CodeBlock
End Sub

Sub Statement_Call()
    Symbol "("
    Expr_Jump Identifier
    Symbol ")"
    Terminator
    CodeBlock
End Sub

Sub Declare_String(Optional TypeDec As String, Optional IsFrameDeclare As Boolean)
    Dim Space As Long
    Dim Ident As String
    Dim Value As String
    
    Ident = Identifier
    
    If IsSymbol("=") Then
        Symbol "="
        Value = StringExpression
    Else
        Value = ""
    End If
    
    If IsSymbol("[") Then
        Symbol "["
        Space = NumberExpression
        Symbol "]"
    Else
        Space = 0
    End If
    
    If IsFrameDeclare = False Then Terminator
    
    If IsFrameDeclare = False Then
        DeclareDataString Switch(TypeDec = "", Ident, TypeDec <> "", TypeDec & "." & Ident), Value, Space
    Else
        AddSymbol CurrentFrame & "." & Ident, 8 + (ArgCount * 4), 0, IS_LOCAL_STRING
        AddFrameDeclare Ident
    End If
    
    CodeBlock
End Sub

Sub Declare_Byte(Optional TypeDec As String, Optional IsFrameDeclare As Boolean)
    Dim Ident As String
    Dim Value As Byte
    
    If IsFrameDeclare Then Declare_DWord TypeDec, IsFrameDeclare: Exit Sub
    
    Ident = Identifier
    If IsSymbol("=") Then
        Symbol "="
        Value = NumberExpression
    Else
        Value = 0
    End If
    
    If IsFrameDeclare = False Then Terminator
    
    If IsByte(CLng(Value)) Then
        If IsFrameDeclare = False Then
            DeclareDataByte Switch(TypeDec = "", Ident, TypeDec <> "", TypeDec & "." & Ident), Value
        Else
            DeclareDataByte CurrentFrame & "." & Ident, Value
            AddFrameDeclare Ident
        End If
    Else
        ErrMessage Value & " is no byte size."
    End If
    
    CodeBlock
End Sub

Sub Declare_Integer(Optional TypeDec As String, Optional IsFrameDeclare As Boolean)
    Dim Ident As String
    Dim Value As Integer
    
    If IsFrameDeclare Then Declare_DWord TypeDec, IsFrameDeclare: Exit Sub
    
    Ident = Identifier
    If IsSymbol("=") Then
        Symbol "="
        Value = NumberExpression
    Else
        Value = 0
    End If
    If IsFrameDeclare = False Then Terminator
    
    If IsWord(CLng(Value)) Then
        If IsFrameDeclare = False Then
            DeclareDataWord Switch(TypeDec = "", Ident, TypeDec <> "", TypeDec & "." & Ident), Value
        Else
            DeclareDataWord CurrentFrame & "." & Ident, Value
            AddFrameDeclare Ident
        End If
    Else
        ErrMessage Value & " is no word size."
    End If
    
    CodeBlock
End Sub


Sub Declare_DWord(Optional TypeDec As String, Optional IsFrameDeclare As Boolean)
    Dim Ident As String
    Dim Value As Long
    
    Ident = Identifier
    If IsSymbol("=") Then
        Symbol "="
        Value = NumberExpression
    Else
        Value = 0
    End If
    If IsFrameDeclare = False Then Terminator
    
    If IsFrameDeclare = False Then
        DeclareDataDWord Switch(TypeDec = "", Ident, TypeDec <> "", TypeDec & "." & Ident), Value
    Else
        AddSymbol CurrentFrame & "." & Ident, 8 + (ArgCount * 4), 0, IS_LOCAL_DWORD
        AddFrameDeclare Ident
    End If
    CodeBlock
End Sub

Sub Declare_Local()
    Dim Ident As String
    Dim IdentII As String
    Dim Value As Variant
    Dim Space As Long
    Dim ArrayValue As Long
    
    Ident = Identifier
    IdentII = Identifier
    
    If LCase(Ident) = "byte" Or _
       LCase(Ident) = "int" Or _
       LCase(Ident) = "dword" Then
        AddSymbol CurrentFrame & "." & IdentII, 8 + (ArgCount * 4), 0, IS_LOCAL_DWORD
        ArgCount = ArgCount + 1
    ElseIf LCase(Ident) = "string" Then
        If IsSymbol("[") Then
            Symbol "["
            Space = NumberExpression
            Symbol "]"
        Else
            Space = 0
        End If
        AddSymbol CurrentFrame & "." & IdentII, 8 + (ArgCount * 4), 0, IS_LOCAL_STRING
        eUniqueID = eUniqueID + 1
        DeclareDataString "Local.String" & eUniqueID, "", Space
        Expr_MovEAXAdress "Local.String" & eUniqueID
        'mov [ebp+number],eax
        AddCodeByte &H89, &H85
        AddCodeDWord 8 + (ArgCount * 4)
        ArgCount = ArgCount + 1
    Else
        ErrMessage "expected identifier 'byte','int','dword' or 'string' but found" & Ident
    End If
    
    Terminator
        

    CodeBlock
End Sub

Sub Eval_Local_Variable(Name As String, Optional OnlySet As Boolean)
    Dim lvID As Long
    Dim iLabel As Long
    SkipBlank
    Symbol "="
    lvID = lvID + eUniqueID: eUniqueID = eUniqueID + 1
    DeclareDataDWord "UniqueExpression" & lvID, &H0
    Expression "UniqueExpression" & lvID
    If GetSymbolType(Name) = IS_LOCAL_STRING Then
        Expr_MovEAXAdress "UniqueExpression" & lvID
    Else
        Expr_MovEAX "UniqueExpression" & lvID
    End If
    'mov [ebp+number],eax
    AddCodeByte &H89, &H85
    For iLabel = 1 To UBound(Symbols)
        If LCase(Symbols(iLabel).Name) = LCase(CurrentFrame & "." & Name) Then
            AddCodeDWord Symbols(iLabel).Offset
        End If
    Next iLabel
    Terminator
    CodeBlock
End Sub

Sub Eval_Variable(Name As String)
    SkipBlank
    Symbol "="
    Expression Name
    Terminator
    CodeBlock
End Sub

Sub Declare_Constant()
    Dim Name As String
    Name = Identifier
    Symbol "="
    SkipBlank
    If IsStringExpression Then
        AddConstant Name, StringExpression
    ElseIf IsNumberExpression Then
        AddConstant Name, NumberExpression
    End If
    Terminator
    CodeBlock
End Sub

Sub Declare_Type()
    Dim Name As String
    
    Name = Identifier
    Symbol "{"
    AddType Name, SourcePos
    While Not IsSymbol("}")
        SourcePos = SourcePos + 1
    Wend
    Symbol "}"
    CodeBlock
End Sub

Sub Declare_Alias()
    Dim Ident As String
    Dim IdentII As String
    Dim FunctionName As String
    Dim FunctionAlias As String
    Dim Library As String
    Dim ParamCount As Long
    
    FunctionAlias = ""
    Ident = Identifier
    IdentII = Identifier
    
    If LCase(IdentII) = "for" Then
        FunctionAlias = Ident
        FunctionName = Identifier
        IdentII = Identifier
    Else
        FunctionName = Ident
        FunctionAlias = Ident
    End If
    
    If LCase(IdentII) = "lib" Then
        Library = StringExpression
    Else
        ErrMessage "expected 'lib' but found '" & Ident & "'"
        Exit Sub
    End If
    
    Symbol ","
    ParamCount = NumberExpression
    Terminator
    
    AddAlias FunctionName, Library, ParamCount, FunctionAlias

    CodeBlock
    
End Sub

Sub Declare_Label()
    AddSymbol Identifier, UBound(CodeSection), 0, IS_LABEL
    Terminator
    CodeBlock
End Sub

Sub Statement_Goto()
    Expr_Jump Identifier
    Terminator
    CodeBlock
End Sub

Sub Call_Alias(Ident As String, Optional FromExpression As Boolean, Optional OnlyInvoke As Boolean)
    Dim i As Integer
    Dim eID As Long
    Dim pCount As Long
    Dim ImportID As Long
    Dim ValueNumeric As Long
    Dim ValueString As String
    
    pCount = GetImportParameterCountByName(Ident)
    ImportID = GetImportIDByName(Ident)
    
    If pCount = -1 Then pCount = UserDefinedParameters
    ReverseParams pCount
    
    Symbol "("
    
    While pCount > 0
        SkipBlank
           
        If IsNumberExpression Then
            eID = eID + sUniqueID: eUniqueID = eUniqueID + 1
            DeclareDataDWord "UniqueExpression" & eID, &H0
            Expression "UniqueExpression" & eID
            AddCodeByte &HFF, &H35
            AddCodeFixup "UniqueExpression" & eID
        ElseIf IsStringExpression Then
            ValueString = StringExpression
            dUniqueID = dUniqueID + 1
            DeclareDataString "Unique" & dUniqueID, ValueString
            Push
            AddFixup "Unique" & dUniqueID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
        Else
            Ident = Identifier
            If IsAlias(Ident) Then ' This has to goto ReverseParams
                Call_Alias Ident, True
                Expr_StoreEAX "Internal.Return"
                AddCodeByte &HFF, &H35
                AddCodeDWord &H0
                AddFixup "Internal.Return", (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
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
                If GetSymbolType(Ident) = IS_STRING Then
                    AddCodeByte &HB8
                    AddCodeFixup Ident
                Else
                    AddCodeWord &H8B 'mov eax,eax
                    PushEAX 'push eax
                End If
            ElseIf IsVariable(Ident) Then
                Dim VarIdent As String
                VarIdent = Ident
                
                If GetSymbolType(VarIdent) = IS_BYTE Or _
                   GetSymbolType(VarIdent) = IS_WORD Or _
                    GetSymbolType(VarIdent) = IS_DWORD Then
                    'AddCodeByte &HFF, &H35
                    AddCodeByte &HFF, &H35
                    AddCodeDWord &H0
                    AddFixup VarIdent, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
                Else
                    AddCodeByte &H68
                    AddCodeDWord &H0
                    AddFixup VarIdent, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
                End If
            ElseIf IsAssignedType(Ident) Then
                AddCodeByte &H68
                AddCodeDWord &H0
                AddFixup Ident, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
            ElseIf IsSymbol("@") Then
                Symbol "@"
                Dim VarIdentII As String
                VarIdentII = Identifier
                AddCodeByte &H68
                AddCodeDWord &H0
                AddFixup VarIdentII, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
            Else
                ErrMessage "unexpected identifier '" & Ident & "'"
                Exit Sub
            End If
            
        End If
    If pCount > 1 Then Symbol ","
    pCount = pCount - 1
    Wend
    
    Symbol ")"
    If Not FromExpression Then Terminator
    
    Invoke
    AddFixup "ImageImportByName" & ImportID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Import
    
    If Not FromExpression Then CodeBlock
End Sub

Sub ReverseParams(pCount As Long)
    Dim i As Long
    Dim ii As Long
    Dim pIdent As String
    Dim Params() As String
    Dim ParamsFound As Long
    Dim OriginalString As String
    Dim ReversedString As String
    If pCount = -1 Then pCount = 0 'ErrMessage "expected '('": Exit Sub
    ReDim Params(pCount) As String
    
    On Error GoTo ErrNumParams
    i = SourcePos + 1
    ParamsFound = 0
    While Mid(Source, i, 1) <> ")"
        ii = i
        pIdent = ""
        While Mid(Source, ii, 1) <> ","
          
            If Mid(Source, ii, 1) = ")" Then
                ParamsFound = ParamsFound + 1
                Params(pCount - ParamsFound) = pIdent
                GoTo CompleteReverse
            End If
            
            If Mid(Source, ii, 1) = "(" Then
                While Mid(Source, ii, 1) <> ")"
                    pIdent = pIdent & Mid(Source, ii, 1)
                    ii = ii + 1
                    If IsEndOfCode(ii) Then Exit Sub
                Wend
            End If
            
            pIdent = pIdent & Mid(Source, ii, 1)
            ii = ii + 1
            If IsEndOfCode(ii) Then Exit Sub
        Wend
        i = ii
        ParamsFound = ParamsFound + 1
        Params(pCount - ParamsFound) = pIdent
        i = i + 1
        If IsEndOfCode(i) Then Exit Sub
    Wend

CompleteReverse:
    If pCount <> ParamsFound Then GoTo ErrNumParams
    i = SourcePos + 1
    While Mid(Source, i, 1) <> ")"
        If Mid(Source, i, 1) = "(" Then
            While Mid(Source, i, 1) <> ")"
                OriginalString = OriginalString & Mid(Source, i, 1)
                i = i + 1
                If IsEndOfCode(i) Then Exit Sub
            Wend
        End If
        OriginalString = OriginalString & Mid(Source, i, 1)
        i = i + 1
        If IsEndOfCode(i) Then Exit Sub
    Wend
    
    For i = 0 To UBound(Params) - 1
        ReversedString = ReversedString & Params(i)
        If i <> UBound(Params) - 1 Then
            ReversedString = ReversedString & ","
        End If
    Next i

    Source = Replace(Source, OriginalString, ReversedString, 1, 1, vbTextCompare)
    Exit Sub
ErrNumParams:
    ErrMessage "expected '" & pCount & "' parameters."
End Sub

Function UserDefinedParameters() As Long
    Dim i As Long
    Dim InStringExpr As Boolean

    i = SourcePos
    InStringExpr = False
    UserDefinedParameters = 1
    While Mid(Source, i, 1) <> ")"
        If Mid(Source, i, 1) = Chr(34) Then
                If InStringExpr = False Then
                    InStringExpr = True
                Else
                    InStringExpr = False
                End If
                i = i + 1
        End If
        If Mid(Source, i, 1) = "," Then
            If InStringExpr = False Then
                UserDefinedParameters = UserDefinedParameters + 1
            End If
        End If
        i = i + 1
        If IsEndOfCode(i) Then Exit Function
    Wend
End Function
