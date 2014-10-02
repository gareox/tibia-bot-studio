Attribute VB_Name = "modInline"
Option Explicit

Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public FasmPath As String

Public Sub DeclareFasmPath()
    FasmPath = InputBox("Specify the fasm path (ex: C:\fasm164\) without the exe file." & vbNewLine & "You need to have 'flat assembler for DOS' installed. You only require this if you want to use inline assembler. (http://www.flatassembler.net)", "Fasm Path", FasmPath)
    SaveSetting "Libry", "Settings", "FasmPath", FasmPath
End Sub

Sub IABlock()
    Dim i As Integer
    Dim Ident As String
    Dim AssemblerSource As String

    AssemblerSource = AssemblerSource & "use32" & vbNewLine & "org 0x" & CStr((512 + (256 * SectionSize))) & vbNewLine
    
    If Dir(FasmPath) = "" Or FasmPath = "" Then ErrMessage "IA.Error: flat assembler path not specified or incorrect.": Exit Sub
    While Mid(Source, SourcePos, 1) <> "}"
        If Mid(Source, SourcePos, 1) = "[" Then
            Ident = ""
            While Mid(Source, SourcePos, 1) <> "]"
                SourcePos = SourcePos + 1
                Ident = Ident & Mid(Source, SourcePos, 1)
                If IsEndOfCode(SourcePos) Then Exit Sub
            Wend
            Ident = Left(Ident, Len(Ident) - 1)
            For i = 1 To UBound(Symbols)
                If Symbols(i).Name = Ident Then
                    AssemblerSource = AssemblerSource & "dword [" & (Symbols(i).Offset - 512) + &H401000 & "]"
                    GoTo FoundSymbol
                End If
            Next i
            ErrMessage "IA.Error '" & Ident & "' is not defined."
FoundSymbol:
        Else
            AssemblerSource = AssemblerSource & Mid(Source, SourcePos, 1)
        End If
        SourcePos = SourcePos + 1
        If IsEndOfCode(SourcePos) Then Exit Sub
    Wend
    
    Dim WaitReturn As Long
    Dim LIAEXESOURCEBINARY As String
    
    If Dir("C:\LIA.ASM") <> "" Then Kill "C:\LIA.ASM"
    If Dir("C:\LIA.EXE") <> "" Then Kill "C:\LIA.EXE"
    Open "C:\LIA.ASM" For Binary As #1
        Put #1, , AssemblerSource
    Close #1

    Call ShellAndWait
   
    Open "C:\LIA.EXE" For Binary As #1
        LIAEXESOURCEBINARY = Space(LOF(1))
        Get #1, , LIAEXESOURCEBINARY
    Close #1
    
    For i = 1 To Len(LIAEXESOURCEBINARY)
        AddCodeByte CByte(Asc(Mid(LIAEXESOURCEBINARY, i, 1)))
    Next i
    If Dir("C:\LIA.ASM") <> "" Then Kill "C:\LIA.ASM"
    If Dir("C:\LIA.EXE") <> "" Then Kill "C:\LIA.EXE"
End Sub

Function ShellAndWait() As Long
    Dim ProgramID As Long
    Dim ExitCode As Long
    Dim hWndProg As Long
    
    ProgramID = Shell(FasmPath & "fasm.exe C:\LIA.ASM C:\LIA.EXE", vbHide)
    hWndProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgramID)
    
    GetExitCodeProcess hWndProg, ExitCode
    Do While ExitCode = STILL_ACTIVE&
        GetExitCodeProcess hWndProg, ExitCode
        DoEvents
    Loop
    CloseHandle hWndProg
    ShellAndWait = ExitCode
End Function

