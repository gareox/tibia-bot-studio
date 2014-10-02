Attribute VB_Name = "ModCrack"
#Const FinalMode = 1
Option Explicit

#If FinalMode Then

Public Declare Function EncipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long
      
Public Declare Function EncipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function GetTibiaCRC Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long
    
Public Declare Function BlackdForceWrite Lib _
    "crackd.dll" (ByVal address As Long, ByRef mybuffer As Byte, ByVal mybuffersize As Long, ByVal hwndClientWindow As Long) As Long
    
    

#Else

    
Public Declare Function EncipherTibiaProtected Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtected Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long
    
Public Declare Function EncipherTibiaProtectedSP Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtectedSP Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function GetTibiaCRC Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long

Public Declare Function BlackdForceWrite Lib _
    "C:\blackdProxy\crackd.dll" (ByVal address As Long, ByRef mybuffer As Byte, ByVal mybuffersize As Long, ByVal hwndClientWindow As Long) As Long
    
'Public Declare Function GetTibiaCRC2 Lib _
'    "C:\blackdProxy\crackd2.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long
    
#End If

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal ByValcbCopy As Long)
    
 Public Type AddressPath
         BaseAddress As Long
         lastJumpIndex As Long
         jump() As Long
End Type

Public Type TypeBuffer
  numBytes As Long
  Packet() As Byte
End Type

Public Type TypeTibiaKey
 key(15) As Byte
End Type

Public ConnectionBuffer() As TypeBuffer
Public ConnectionBufferLogin() As TypeBuffer
Public ProcessID() As Long
Public AlternativeBinding As Long
Public packetKey() As TypeTibiaKey
Public loginPacketKey() As TypeTibiaKey
Public gotFirstLoginPacket() As Boolean
Public UseCrackd As Boolean
Public adrConnectionKey As Long
Public adrSelectedCharIndex As AddressPath
Public adrLastPacket As Long
Public adrCharListPtr As Long
Public adrCharListPtrEND As Long
Public debugStrangeFail As String
Public MAXCHARACTERLEN As Long
Public manualDebugOrder As Long
Public GameServerDictionary As Scripting.Dictionary  ' A dictionary server (string) -> IP (string)
Public GameServerDictionaryDOMAIN As Scripting.Dictionary
Public ProcessidIPrelations As Scripting.Dictionary

Public Sub GiveCrackdDllErrorMessage(pres As Long, ByRef packet1() As Byte, ByRef packet2() As Byte, ubound1 As Long, ubound2 As Long, p As Long)
  Dim errorstr As String
        Select Case pres
        Case -1
          errorstr = "ERROR -1 : Packet header is not multiplier of 8"
        Case -2
          errorstr = "ERROR -2 : Wrong size of key (ubound must be 15)"
        Case -3
          errorstr = "ERROR -3 : Header of packet doesn't match with real size of the packet"
        Case -4
          errorstr = "ERROR -4 : This is not a packet"
        Case Else
          errorstr = "ERROR " & CStr(pres) & " : Unknown error"
        End Select
  errorstr = errorstr & vbCrLf & "PARAMETERS:" & vbCrLf & _
    "Packet : " & showAsStr2(packet1, 2) & vbCrLf & _
    "Key : " & showAsStr2(packet2, 2) & vbCrLf & _
    "Ubound(Packet) : " & CStr(ubound1) & vbCrLf & _
    "Ubound(Key) : " & CStr(ubound2) & vbCrLf & _
    "Called at point : " & CStr(p)

  MsgBox errorstr
End Sub


Public Function readLoginTibiaKeyAtPID(idConnection As Integer, ProcessID As Long) As Long
  'On Error GoTo goterr
  Dim abyte As Byte
  Dim i As Integer

  If (ProcessID = -1) Then
    readLoginTibiaKeyAtPID = -1
  Else
    For i = 0 To 15
      abyte = Memory_ReadByte(ProcessID, theOffset + adrConnectionKey + i)
      loginPacketKey(idConnection).key(i) = abyte
    Next i
    readLoginTibiaKeyAtPID = 0
  End If
  Exit Function
goterr:
  readLoginTibiaKeyAtPID = -1
End Function

Public Function readTibiaKeyAtPID(idConnection As Integer, ProcessID As Long) As Long
  Dim abyte As Byte
  Dim i As Integer
    For i = 0 To 15
      abyte = Memory_ReadByte(ProcessID, theOffset + adrConnectionKey + i)
      packetKey(idConnection).key(i) = abyte
    Next i
  readTibiaKeyAtPID = 0
End Function

Public Function GetProcessIdFromIP(strIP As String) As Long
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If ProcessidIPrelations.Exists(strIP) = True Then
    GetProcessIdFromIP = ProcessidIPrelations.Item(strIP)
  Else
    GetProcessIdFromIP = 0 ' error
  End If
End Function


Public Function GiveProcessIDbyLastPacket(ByRef Packet() As Byte, Optional strIP As String = "", Optional FromIP As String = "?", Optional part As String = "LOGIN1") As Long
  Dim tibiaclient As Long
  Dim hWndDesktop As Long
  Dim status As Byte
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  Dim errmessage As String
  Dim res As Long
  Dim comparing1 As String
  Dim comparing2 As String
  Dim tcount As Long
  Dim trivialRes As Long
  Dim packetSizeForComparing As Long
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
  tcount = 0
  
  If AlternativeBinding <> 0 Then
    If strIP <> "" Then
      GiveProcessIDbyLastPacket = GetProcessIdFromIP(strIP)
      Exit Function
    End If
  End If
  debugStrangeFail = ""
  
  debugStrangeFail = "WARNING on GiveProcessIDbyLastPacket . Doing a complete report:"
  res = 0
  sucess = -2
  packetSizeForComparing = GetTheLong(Packet(0), Packet(1)) ' fix since 11.7 : lets only compare first subpacket
  
  comparing1 = showAsStr2(Packet, 0, packetSizeForComparing + 1)
  
  'hWndDesktop = GetDesktopWindow()
  'debugStrangeFail = debugStrangeFail & vbCrLf & "GetDesktopWindow() returned " & CStr(hWndDesktop)
  debugStrangeFail = debugStrangeFail & vbCrLf & "Now trying to determine what client sent the packet that Blackd Proxy just received."
  debugStrangeFail = debugStrangeFail & vbCrLf & "BLACKDPROXY RECEIVED, from ip [" & FromIP & "] at " & part & " :" & comparing1
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, "tibiaclient", vbNullString)
    If tibiaclient = 0 Then
      debugStrangeFail = debugStrangeFail & vbCrLf & "Found a total of " & CStr(tcount) & " Tibia client(s) opened"
      Exit Do
    Else
      trivialRes = tibiaclient
      tcount = tcount + 1
      comparing2 = GetLastPacket(tibiaclient, packetSizeForComparing + 1)
      debugStrangeFail = debugStrangeFail & vbCrLf & "CLIENT #" & CStr(tcount) & " HAVE SENT :" & comparing2
      If (comparing1 = comparing2) Then
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...MATCH at pid " & CStr(tibiaclient)
        res = tibiaclient
        sucess = 0
        'Exit Do
      Else
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...FAIL! at pid " & CStr(tibiaclient)
      End If
    End If
  Loop
  If (sucess = 0) Then
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function worked fine."
  Else
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function failed!"
    If tcount = 1 Then
        debugStrangeFail = debugStrangeFail & vbCrLf & "However, there is a trivial match since only 1 client was detected: " & CStr(trivialRes)
        sucess = 0
        res = trivialRes
    Else
        debugStrangeFail = debugStrangeFail & vbCrLf & "Please report to daniel@blackdtools.com"
    End If
  End If
  If (sucess = 0) Then
    GiveProcessIDbyLastPacket = res
    Exit Function
  End If
  'frmMain.TxtPackets.Text = frmMain.TxtPackets.Text & vbCrLf & debugStrangeFail
  'LogOnFile "errors.txt", debugStrangeFail
  GiveProcessIDbyLastPacket = 0
  Exit Function
goterr:
  errmessage = "Function failure : GiveProcessIDbyLastPacket could not match idconnection<->pid : Error number " & CStr(Err.Number) & " : " & Err.Description
  'frmMain.TxtPackets.Text = frmMain.TxtPackets.Text & vbCrLf & errmessage
  'LogOnFile "errors.txt", errmessage
  GiveProcessIDbyLastPacket = 0
End Function


Public Function GetLastPacket(pID As Long, lngp As Long) As String
  Dim res As Boolean
  Dim i As Long
  Dim B As Byte
  Dim errmessage As String
  Dim packetR() As Byte
  On Error GoTo cantdoit
  ReDim packetR(lngp)
  For i = 0 To lngp
    B = Memory_ReadByte(pID, theOffset + adrLastPacket + i)
    packetR(i) = B
  Next i
  GetLastPacket = showAsStr2(packetR, 0)
  Exit Function
cantdoit:
  GetLastPacket = "ERROR"
End Function


