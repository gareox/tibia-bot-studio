Attribute VB_Name = "modCompiler"
Option Explicit

Public CodeSection() As Byte
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub InitCompiler()
    ReDim CodeSection(0) As Byte
End Sub

Sub Compile(sFile As String, Run As Boolean)
    On Error GoTo RaiseError
    
    Call InitCompiler: InitErrors: InitSymbols
    Call InitFixups: InitTypes: InitFrames
    Call InitImports: InitData: InitParser
    Call Parse: If CheckSectionSize = True Then Compile sFile, Run: Exit Sub
    Call InitLinker: DoFixups
    If pError = True Then MsgBox Errors, vbInformation: Exit Sub
    Call Link(sFile, Run): Exit Sub

RaiseError:
    ErrMessage Err.Description
    InfMessage "Libry.Internal.Error: " & Err.Description
End Sub

Sub AddCodeDWord(Value As Long)
    AddCodeWord LoWord(Value)
    AddCodeWord HiWord(Value)
End Sub

Sub AddCodeWord(Value As Integer)
    AddCodeByte LoByte(Value), HiByte(Value)
End Sub

Sub AddCodeByte(ParamArray Bytes() As Variant)
    Dim i As Integer
    For i = 0 To UBound(Bytes)
        ReDim Preserve CodeSection(UBound(CodeSection) + 1) As Byte
        CodeSection(UBound(CodeSection)) = CByte(Bytes(i))
    Next i
End Sub
