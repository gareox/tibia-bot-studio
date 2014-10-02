Attribute VB_Name = "modParser"
Option Explicit

Public Source As String
Public SourcePos As Long

Sub InitParser()
    AppType = 0
    pError = False
    Source = frmMain.Code.Text
    Source = Replace(Source, " _" & vbCrLf, "", 1, -1, vbTextCompare)
    Source = Replace(Source, "_" & vbCrLf, "", 1, -1, vbTextCompare)
    Source = Replace(Source, vbTab, "", 1, -1, vbTextCompare)
End Sub

Sub Parse()
    Dim iID As Long: SourcePos = 1: CurrentFrame = ""
    InitApplication
    AddAlias "lstrcpyA", "KERNEL32.DLL", 2
    AddAlias "lstrcmpA", "KERNEL32.DLL", 2
    AddAlias "wsprintfA", "USER32.DLL", -1, "Format"
    AddAlias "GetModuleHandleA", "KERNEL32.DLL", 1
    DeclareDataDWord "Instance", &H0
    DeclareDataDWord "Internal.Local", &H0
    DeclareDataDWord "Internal.Array", &H0
    DeclareDataDWord "Internal.Compare.I", &H0
    DeclareDataDWord "Internal.Compare.II", &H0
    DeclareDataDWord "Internal.Return", &H0
    AddCodeByte &H6A, &H0
    iID = GetImportIDByName("GetModuleHandleA")
    Invoke
    AddFixup "ImageImportByName" & iID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Code
    AddCodeByte &HA3: AddCodeFixup "Instance"
    Expr_Call "main"
    CodeBlock
End Sub

Sub CodeBlock()
    Dim Ident As String
    
    Call SkipBlank: Ident = Identifier
    If Ident = "" Then Exit Sub: If pError = True Then Exit Sub
    
    Select Case LCase(Ident)
        Case "alias": Declare_Alias
        Case "const": Declare_Constant
        Case "type": Declare_Type
        Case "frame": Declare_Frame
        Case "return": Statement_Return
        Case "call": Statement_Call
        Case "end": SourcePos = SourcePos - 3: Exit Sub
        Case "if": Statement_If
        Case "while": Statement_While
        Case "label": Declare_Label
        Case "goto": Statement_Goto
        Case "include": Statement_Include
        Case "iassembler": Statement_IA
        Case "local": Declare_Local
        Case Else
            If IsAlias(Ident) Then
                Call_Alias Ident
            ElseIf IsLocalVariable(Ident) Then
                Eval_Local_Variable Ident
            ElseIf IsVariable(Ident) Then
                Eval_Variable Ident
            ElseIf IsType(Ident) Then
                Assign_Type Ident
            ElseIf IsFrame(Ident) Then
                Call_Frame Ident
            Else
                VariableBlock Ident
            End If
    End Select
End Sub

Sub VariableBlock(Ident As String, Optional IsFrameDeclare As Boolean)
    Select Case LCase(Ident)
        Case "byte": Declare_Byte TypeDec, IsFrameDeclare
        Case "int": Declare_Integer TypeDec, IsFrameDeclare
        Case "dword": Declare_DWord TypeDec, IsFrameDeclare
        Case "string": Declare_String TypeDec, IsFrameDeclare
    Case Else
        ErrMessage "unknow Identifier '" & Ident & "'"
        Exit Sub
    End Select
End Sub

Function Identifier() As String
    SkipBlank
    While (UCase(Mid(Source, SourcePos, 1)) >= "A" And _
           UCase(Mid(Source, SourcePos, 1)) <= "Z") Or _
           Mid(Source, SourcePos, 1) = "." Or _
           Mid(Source, SourcePos, 1) = "_"
           Identifier = Identifier & Mid(Source, SourcePos, 1)
            Skip
    Wend
End Function

Sub Skip(Optional NumberOfChars As Integer)
    SourcePos = SourcePos + 1 + NumberOfChars
End Sub

Sub SkipBlank()
    While Mid(Source, SourcePos, 1) = " " Or _
          Mid(Source, SourcePos, 1) = vbCr Or _
          Mid(Source, SourcePos, 1) = vbLf Or _
          Mid(Source, SourcePos, 1) = vbTab
        Skip
    Wend
End Sub

Sub Symbol(Value As String)
    SkipBlank
    If Mid(Source, SourcePos, 1) = Value Then
        Skip
    Else
        ErrMessage "expected symbol '" & Value & "' but found '" & Mid(Source, SourcePos, 1) & "'"
    End If
End Sub

Function IsAlias(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Aliases)
        If Aliases(i).FunctionAlias = Ident Then
            IsAlias = True
            Exit Function
        End If
    Next i
End Function

Function IsIdent(Word As String) As Boolean
    If LCase(Mid(Source, SourcePos, Len(Word))) = LCase(Word) Then IsIdent = True
End Function

Function IsSymbol(Value As String) As Boolean
    If Mid(Source, SourcePos, Len(Value)) = Value Then IsSymbol = True
End Function

Sub Blank()
    If Mid(Source, SourcePos, 1) = " " Then
        Skip
    Else
        ErrMessage "expected blank ' ' but found '" & Mid(Source, SourcePos, 1) & "'"
    End If
End Sub

Sub Terminator()
    If Mid(Source, SourcePos, 1) = ";" Then
        Skip
    Else
        ErrMessage "expected terminator (;) but found '" & Mid(Source, SourcePos, 1) & "'"
    End If
End Sub

Function IsVariable(Name As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Name) Then
            If Symbols(i).SymType = IS_BYTE Or _
               Symbols(i).SymType = IS_WORD Or _
               Symbols(i).SymType = IS_DWORD Or _
               Symbols(i).SymType = IS_STRING Then
                IsVariable = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsFrame(Name As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(Symbols)
        If LCase(Symbols(i).Name) = LCase(Name) Then
            If Symbols(i).SymType = IS_FRAME Then
                IsFrame = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsEndOfCode(Value As Long) As Boolean
    If Value = Len(Source) Then
        ErrMessage "found end of code. but expected ')' or ','"
        IsEndOfCode = True
    End If
End Function

Sub EndProc()
    If IsIdent("end") Then
        Call Identifier
        Terminator
        CurrentFrame = ""
        Exit Sub
    Else
        ErrMessage "could not find end of procedure."
        Exit Sub
    End If
End Sub

Function IsVariableExpression() As Boolean
    If (UCase(Mid(Source, SourcePos, 1)) >= "A" And _
           UCase(Mid(Source, SourcePos, 1)) <= "Z") Then
        IsVariableExpression = True
    End If
End Function

Function IsStringExpression() As Boolean
    If Mid(Source, SourcePos, 1) = Chr(34) Then
        IsStringExpression = True
    End If
End Function

Function IsNumberExpression() As Boolean
    If IsNumeric(Mid(Source, SourcePos, 1)) Or _
                 Mid(Source, SourcePos, 1) = "-" Or _
                 Mid(Source, SourcePos, 1) = "$" Then
        IsNumberExpression = True
    End If
End Function

Function NumberExpression() As Variant
    Dim CToHex As Boolean
    SkipBlank
    If IsSymbol("$") Then Skip: CToHex = True
    While IsNumeric(Mid(Source, SourcePos, 1)) Or Mid(Source, SourcePos, 1) = "-" Or _
          Mid(Source, SourcePos, 1) = "A" Or Mid(Source, SourcePos, 1) = "B" Or _
          Mid(Source, SourcePos, 1) = "C" Or Mid(Source, SourcePos, 1) = "D" Or _
          Mid(Source, SourcePos, 1) = "E" Or Mid(Source, SourcePos, 1) = "F"
            NumberExpression = NumberExpression & Mid(Source, SourcePos, 1)
            Skip
    Wend
    If CToHex = True Then NumberExpression = CLng("&H" & NumberExpression)
End Function

Function VariableExpression() As String
    SkipBlank
    VariableExpression = Identifier
End Function

Function StringExpression() As String
    SkipBlank
    Symbol Chr(34)
    While Mid(Source, SourcePos, 1) <> Chr(34)
            If Mid(Source, SourcePos, 2) = "\n" Then SourcePos = SourcePos + 2: StringExpression = StringExpression & vbCrLf
            If Mid(Source, SourcePos, 2) = "\t" Then SourcePos = SourcePos + 2: StringExpression = StringExpression & vbTab
            StringExpression = StringExpression & Mid(Source, SourcePos, 1)
            If Mid(Source, SourcePos, 1) = vbCr Or _
               Mid(Source, SourcePos, 1) = "" Then
                ErrMessage "unterminated string": Exit Function
            End If
            Skip
    Wend
    Symbol Chr(34)
End Function



