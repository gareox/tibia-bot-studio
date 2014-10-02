Attribute VB_Name = "ModAdvanced"
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORY_BASIC_INFORMATION ' 28 bytes
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    state As Long
    Protect As Long
    lType As Long
End Type

Private Type SYSTEM_INFO ' 36 Bytes
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDNEXT = 2

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_VM_WRITE = (&H20)
Private Const PROCESS_VM_OPERATION = (&H8)
Private Const PROCESS_QUERY_INFORMATION = (&H400)
Private Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION

Private Const MEM_PRIVATE& = &H20000
Private Const MEM_COMMIT& = &H1000

Public useDynamicOffset As String
Public tibiaModuleRegionSize As Long  '= &H2FC000
Public theBase As Long

Public Function getProcessBase(ByVal hProcess As Long, ByVal expectedRegionSize As Long, Optional PIDinsteadHp As Boolean = False) As Long
    On Error GoTo goterr
    ' expectedRegionSize is used again
    Dim lpMem As Long, ret As Long, lLenMBI As Long
    Dim lWritten As Long, CalcAddress As Long, lPos As Long
    Dim sBuffer As String
    Dim sSearchString As String, sReplaceString As String
    Dim si As SYSTEM_INFO
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim realH As Long
    Dim pid As Long
    Dim res As Long
    If PIDinsteadHp = True Then
        res = GetWindowThreadProcessId(hProcess, pid)
        realH = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
        hProcess = realH
    End If
    Call GetSystemInfo(si)
    lpMem = si.lpMinimumApplicationAddress
    lLenMBI = Len(mbi)
    ' Scan memory
    Do While lpMem < si.lpMaximumApplicationAddress
        mbi.RegionSize = 0
        ret = VirtualQueryEx(hProcess, ByVal lpMem, mbi, lLenMBI)
        If ret = lLenMBI Then
           If (mbi.state = MEM_COMMIT) Then
                If mbi.AllocationProtect = &H80 Then
                If mbi.BaseAddress - mbi.AllocationBase = &H1000 Then
                If mbi.Protect = &H20 Then
                If (mbi.RegionSize = expectedRegionSize) Then
                    res = mbi.AllocationBase
                    'Debug.Print "The new result is " & CStr(res)
                    If PIDinsteadHp = True Then
                      CloseHandle hProcess
                    End If
                    getProcessBase = res
                    Exit Function
                End If
                End If
                End If
                End If
           End If
           lpMem = mbi.BaseAddress + mbi.RegionSize
        Else
           Exit Do
        End If
    Loop
    If PIDinsteadHp = True Then
       CloseHandle hProcess
    End If
goterr:
    getProcessBase = 0
End Function


Public Function theOffset() As Long
If useDynamicOffset = "yes" Then
    theBase = getProcessBase(Tibia_Hwnd, tibiaModuleRegionSize, True)
    theOffset = theBase - &H400000
End If
End Function
    
