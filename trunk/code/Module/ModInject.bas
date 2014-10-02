Attribute VB_Name = "ModInject"
Option Explicit

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lPAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const MEM_COMMIT As Long = &H1000
Private Const PAGE_READWRITE As Long = &H4
 
 
Public Sub Inject()
Dim lLL As Long, lhWnd As Long, lThID As Long, pID As Long, dllPath As String, hRemMem As Long, lInject As Long, hRemThread As Long
Dim lMyMsg As Long, lmhWnd As Long, lLibraryHandle As Long, lFunctionHandle As Long, iMyMsg As Long
lLL = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
lhWnd = FindWindow(Proseso, vbNullString)
WriteFile
If lhWnd Then
    GetWindowThreadProcessId lhWnd, lThID

    pID = OpenProcess(PROCESS_ALL_ACCESS, False, lThID)

    If pID Then
        dllPath = App.Path & "\TsbHook.dll"
        hRemMem = VirtualAllocEx(pID, ByVal 0, Len(dllPath), MEM_COMMIT, ByVal PAGE_READWRITE)
            If hRemMem Then
            lInject = WriteProcessMemory(pID, ByVal hRemMem, ByVal dllPath, Len(dllPath), vbNull)
                
            DoEvents
                If lInject Then
                hRemThread = CreateRemoteThread(pID, 0, 0, lLL, hRemMem, 0, 0)
                WaitForSingleObject hRemThread, 5000
                
                FrmMain.StatusBar.SimpleText = "Injected!"
                Wait 1
                FrmMain.StatusBar.SimpleText = "Tibia Studio"
                
                    'If hRemThread Then
                    '     lLibraryHandle = LoadLibrary(dllPath)
                    '    iMyMsg = GetProcAddress(pID, "InitConnect")
                    'End If
                End If
            End If
    End If
End If
End Sub

