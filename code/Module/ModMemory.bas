Attribute VB_Name = "ModMemory"
 Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
             
 ' Constants we need for the functions that read memory
'
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public ClientSelected As Long
Public myBpos As Long
Public myID As Long

Public Function Tibia_Hwnd() As Long
    Tibia_Hwnd = ClientSelected
    'Tibia_Hwnd = FindWindowEx(0&, Tibia_Hwnd, Proseso, vbNullString)
End Function

Public Function FindProcess(window As Long) As Long

Dim procID As Long
Dim process As Long

     Call GetWindowThreadProcessId(window, procID)
     process = OpenProcess(PROCESS_ALL_ACCESS, False, procID)

    FindProcess = process
End Function

Public Function CountTibiaWindows() As Long
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim countt As Long
  countt = 0
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, Proseso, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      countt = countt + 1
    End If
   
  Loop
  CountTibiaWindows = countt
End Function

Public Function LongToByte(Number As Long, nByte As Long) As Byte
        Dim ByteArray(0 To 3) As Byte
    
        CopyMemory ByteArray(0), ByVal VarPtr(Number), Len(Number)
        LongToByte = ByteArray(nByte - 1)
    End Function

Public Function Memory_ReadByte(process_Hwnd As Long, address As Long) As Byte
  
   ' Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Byte   ' Byte
    
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, pID)
   If (phandle = 0) Then Exit Function
   
   ' Read Long
   ReadProcessMemory phandle, address, valbuffer, 1, 0&
       
   ' Return
   Memory_ReadByte = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function

Public Function Memory_ReadLong(ByVal process_Hwnd As Long, ByVal address As Long) As Long
  
   ' Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
   
   Dim offset As Long
    
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pID
   
   ' Use the pid to get a Process Handle
   'phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pID) ' more powerfull
   If (phandle = 0) Then Exit Function
   
   
   ' Read Long
   ReadProcessMemory phandle, address, valbuffer, 4, 0&
       
   ' Return
   Memory_ReadLong = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function

Public Function Memory_ReadString(Windowhwnd As Long, address As Long) As String
  
   ' Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim str(255) As Byte    ' Read String Array
 
   ' First get a handle to the "game" window
   If (Windowhwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId Windowhwnd, pID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, pID)
   If (phandle = 0) Then Exit Function
    
   ' Read String
   ReadProcessMemory phandle, address, str(0), 255, 0&
       
   ' Return String
   Memory_ReadString = StrConv(str, vbUnicode)
       
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Function Memory_WriteString(TibiaHwnd As Long, address As Long, tibiaString As String)

   '// Write String

   'Get length of current text in the status
   Dim tempstr As String
   tempstr = Memory_ReadString(Tibia_Hwnd, address)

   'Write to the status text now, with the LENGTH
   Call Memory_WriteString_Len(Tibia_Hwnd, address, tibiaString, Len(tempstr))

End Function
Public Function Memory_WriteString_Len(Windowhwnd As Long, address As Long, valbufferr As String, vallength As Long)

   'Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle

   ' First get a handle to the "game" window
   If (Windowhwnd = 0) Then Exit Function

   ' We can now get the pid
   GetWindowThreadProcessId Windowhwnd, pID

   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pID)
   If (phandle = 0) Then Exit Function

   ' Now we can write to memory space with length of that string
   WriteProcessMemory phandle, address, ByVal valbufferr, vallength, 0&

   ' Close the Process Handle
   CloseHandle phandle

End Function

Public Sub Memory_WriteByte(Windowhwnd As Long, address As Long, valbuffer As Byte)

   'Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (Windowhwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId Windowhwnd, pID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pID)
   If (phandle = 0) Then Exit Sub
   
   ' Write Long
   WriteProcessMemory phandle, address, valbuffer, 1, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub

Public Sub Memory_WriteLong(Windowhwnd As Long, address As Long, valbuffer As Long)

   'Declare some variables we need
   Dim pID As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (Windowhwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId Windowhwnd, pID
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pID)
   If (phandle = 0) Then Exit Sub
   
   ' Write Long
   WriteProcessMemory phandle, address, valbuffer, 4, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub

Public Function MemoryWriteLong(TibiaHwnd As Long, address As Long, valbuffer As Long)

 
   Dim pID As Long
   Dim phandle As Long
   

   If (TibiaHwnd = 0) Then

     Exit Function
   End If
   

   GetWindowThreadProcessId TibiaHwnd, pID
   

   phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pID)
   If (phandle = 0) Then

     Exit Function
   End If
    
    ' Write Long
   WriteProcessMemory phandle, address, valbuffer, 4, 0&
   
    ' Close the Process Handle
   CloseHandle phandle
  
End Function

