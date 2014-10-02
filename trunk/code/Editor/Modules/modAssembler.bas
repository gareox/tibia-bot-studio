Attribute VB_Name = "modAssembler"
Option Explicit

Sub Push()
    AddCodeByte &H68
    AddCodeDWord &H0
End Sub

Sub PushEAX()
    AddCodeByte &H50
End Sub

Sub Invoke()
    AddCodeByte &HFF
    AddCodeByte &H15
    AddCodeDWord &H0
End Sub

Sub Expr_Set_Var(CopyTo As String, Variable As String)
    Expr_MovEAX Variable
    'mov [CopyTo],eax
    AddCodeByte &HA3
    AddCodeFixup CopyTo
End Sub

Sub Expr_Sub_Var(SubTo As String, Variable As String)
    Expr_MovEAX Variable
    'mov [SubTo],eax
    AddCodeByte &H29, &H5
    AddCodeFixup SubTo
End Sub

Sub Expr_Add_Var(AddTo As String, Variable As String)
    Expr_MovEAX Variable
    'mov [AddTo],eax
    AddCodeByte &H1, &H5
    AddCodeFixup AddTo
End Sub

Sub Expr_Mul_Var(Name As String, Variable As String)
    Expr_MovEAX Name
    'mov ebx,[Variable]
    AddCodeByte &H8B, &H1D
    AddCodeFixup Variable
    'mul ebx
    AddCodeByte &HF7, &HE3
    'mov [Name],eax
    AddCodeByte &HA3
    AddCodeFixup Name
End Sub

Sub Expr_Div_Var(Name As String, Variable As String)
    Expr_MovEAX Name
    'mov ebx,[Variable]
    AddCodeByte &H8B, &H1D
    AddCodeFixup Variable
    'div bl
    AddCodeByte &HF6, &HF3
    'mov [Name],eax
    AddCodeByte &HA3
    AddCodeFixup Name
End Sub

Sub Expr_Set(Name As String, Value As Long)
    AddCodeByte &HC7, &H5
    AddCodeFixup Name
    AddCodeDWord Value
End Sub

Sub Expr_Mul(Name As String, Value As Long)
    Expr_MovEAX Name
    'mov ebx,Value
    AddCodeByte &HBB
    AddCodeDWord Value
    'mul ebx
    AddCodeByte &HF7, &HE3
    'mov [Name],eax
    AddCodeByte &HA3
    AddCodeFixup Name
End Sub

Sub Expr_Div(Name As String, Value As Long)
    Expr_MovEAX Name
    'mov ebx,Value
    AddCodeByte &HBB
    AddCodeDWord Value
    'div bl
    AddCodeByte &HF6, &HF3
    'mov [Name],eax
    AddCodeByte &HA3
    AddCodeFixup Name
End Sub

Sub Expr_Add(Name As String, Value As Long)
    AddCodeByte &H81, &H5
    AddCodeFixup Name
    AddCodeDWord Value
End Sub

Sub Expr_Sub(Name As String, Value As Long)
    AddCodeByte &H81, &H2D
    AddCodeFixup Name
    AddCodeDWord Value
End Sub

Sub Expr_MovEAX(Name As String)
    'mov eax,[name]
    AddCodeByte &HA1
    AddCodeFixup Name
End Sub

Sub Expr_MovEAXAdress(Name As String)
    'mov eax,name
    AddCodeByte &HB8
    AddCodeFixup Name
End Sub

Sub Expr_MovECX(Name As String)
    'mov ecx,value
    AddCodeByte &H8B, &HD
    AddCodeFixup Name
End Sub

Sub Expr_MovEDX(Name As String)
    'mov edx,[name]
    AddCodeByte &H8B, &H15
    AddCodeFixup Name
End Sub

Sub Expr_Compare(VarI As String, VarII As String)
    Expr_MovEAX VarI
    Expr_MovEDX VarII
    'cmp eax,edx
    AddCodeByte &H39, &HD0
End Sub

Sub Expr_StringCompare(VarI As String, VarII As String)
    Dim ImportID As Long: ImportID = GetImportIDByName("lstrcmpA")
    AddCodeByte &HFF, &H35
    AddCodeDWord &H0
    AddFixup VarII, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
    AddCodeByte &HFF, &H35
    AddCodeDWord &H0
    AddFixup VarI, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
    Invoke
    AddFixup "ImageImportByName" & ImportID, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Import
    'cmp eax,0
    AddCodeByte &H3D, &H0, &H0, &H0, &H0
End Sub

Sub Expr_StoreEAX(Variable As String)
    AddCodeByte &HA3
    AddCodeFixup Variable
End Sub

Sub Expr_JumpEqual(Name As String)
    AddCodeByte &HF, &H84
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_JumpNotEqual(Name As String)
    AddCodeByte &HF, &H85
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_JumpBelow(Name As String)
    AddCodeByte &HF, &H82
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_JumpBelowEqual(Name As String)
    AddCodeByte &HF, &H86
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_JumpAbove(Name As String)
    AddCodeByte &HF, &H8F
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_JumpAboveEqual(Name As String)
    AddCodeByte &HF, &H8D
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_Jump(Name As String)
    AddCodeByte &HE9
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub Expr_Call(Name As String)
    AddCodeByte &HE8
    AddCodeDWord &HFFFFFFFF
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), UBound(CodeSection), Code
End Sub

Sub StartFrame()
    AddCodeByte &H55, &H89, &HE5
End Sub

Sub EndFrame(Value As Integer)
    AddCodeByte &HC9
    AddCodeByte &HC2
    AddCodeWord Value
End Sub

