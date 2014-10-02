Attribute VB_Name = "modGetHProcExe"
'I DID NOT CREATE THIS MODULE!  Im in love with who ever did though   ;)
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Option Explicit
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)


Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Function GetHProcExe(strExeName As String) As Long
Dim hSnap As Long
    'Create a snapshot of all of the processes, and information
    'about them (saving the handle so we can iterate through the
    'processes)
    hSnap = CreateToolhelpSnapshot(2, 0)
    
Dim peProcess As PROCESSENTRY32
    peProcess.dwSize = LenB(peProcess)
    
Dim nProcess As Long
    nProcess = Process32First(hSnap, peProcess)
    
    'Loop through the processes until we find the one we want
    'and return its process handle
    Do While nProcess
        If StrComp(Trim$(peProcess.szExeFile), strExeName, vbTextCompare) _
            = 0 Then
            GetHProcExe = OpenProcess(PROCESS_ALL_ACCESS, False, peProcess.th32ProcessID)
            Exit Function
        End If
        peProcess.szExeFile = vbNullString
        nProcess = Process32Next(hSnap, peProcess)
    Loop
    CloseHandle hSnap
End Function

Public Function FindProc(ProcName As String) As Long
Dim Hwnd As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
Hwnd = FindWindow(ProcName, vbNullString)
GetWindowThreadProcessId Hwnd, ProcessID
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
FindProc = ProcessHandle
End Function




