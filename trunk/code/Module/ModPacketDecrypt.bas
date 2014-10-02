Attribute VB_Name = "ModPacketDecrypt"
Option Explicit
Public TibiaVersionLong As Long
  
  
  
Public Sub ReadPacket(ByVal Index As Integer, ByRef realRawPacket() As Byte)
    Dim Packet() As Byte
    Dim rawpacket() As Byte
    Dim hbytes As Long
    Dim res As Long
    Dim SPpacket() As Byte
    Dim pres As Long
    Dim SPpos, SPlen, SPlim
    
   
     SPpos = 0
     SPlim = UBound(realRawPacket)
     SPlen = GetTheLong(realRawPacket(SPpos), realRawPacket(SPpos + 1))
            ReDim SPpacket(SPlen + 1)
            RtlMoveMemory SPpacket(0), realRawPacket(SPpos), (SPlen + 2)
            pres = DecipherTibiaProtectedSP(SPpacket(0), packetKey(Index).key(0), UBound(SPpacket), UBound(packetKey(Index).key))
    

            
    If (pres = 0) Then

            If TibiaVersionLong < 830 Then
                hbytes = GetTheLong(SPpacket(2), SPpacket(3))
                'Debug.Print showAsStr(SPpacket, True)
                ReDim Packet(hbytes + 1)
                RtlMoveMemory Packet(0), SPpacket(2), (hbytes + 2)
            Else
                hbytes = GetTheLong(SPpacket(6), SPpacket(7))
                ReDim Packet(hbytes + 1)
                RtlMoveMemory Packet(0), SPpacket(6), (hbytes + 2)
            End If
        
      Else
      If pres = -1 Then
              ' somehow a login packet arrived here
              ReDim Packet(UBound(realRawPacket))
              RtlMoveMemory Packet(0), realRawPacket(0), UBound(realRawPacket) + 1
            Else
              GiveCrackdDllErrorMessage pres, SPpacket, packetKey(Index).key, UBound(SPpacket), UBound(packetKey(Index).key), 1
              Exit Sub
            End If
    End If
    
    If res <> 1 Then
        UnifiedSendToServerGame Index, Packet, True
    End If
End Sub

Public Sub UnifiedSendToServerGame(ByVal Index As Integer, ByRef Packet() As Byte, logIt As Boolean)
 Dim extrab As Long
  Dim i As Long
  Dim rnumber As Byte
  Dim totalLong As Long
  Dim goodPacket() As Byte
  Dim hbytes As Long
  Dim pres As Long
  Dim thedamnCRC As Long
  Dim fourBytesCRC(3) As Byte
  Dim onlygood As Long

If IsConnected(Index) Then
  If (UseCrackd) Then
  
      If logIt = True Then
        DebugAdd FrmMain.TxtPackets, "GAMECLIENT" & Index & ": " & showAsStr2(Packet, 0) & vbNewLine, True
      End If


    
    If TibiaVersionLong < 830 Then
        totalLong = GetTheLong(Packet(0), Packet(1))
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
        totalLong = totalLong + 2
        ReDim goodPacket(totalLong + 1)
        hbytes = UBound(Packet) + 1
        RtlMoveMemory goodPacket(2), Packet(0), (totalLong)
        goodPacket(0) = LowByteOfLong(totalLong)
        goodPacket(1) = HighByteOfLong(totalLong)
        pres = EncipherTibiaProtected(goodPacket(0), packetKey(Index).key(0), UBound(goodPacket), UBound(packetKey(Index).key))
    
    Else
        'Debug.Print "1>> " & frmMain.showAsStr(packet, True) ' DEBUGGGGGGGGGGGGGGGGGGGGGGGGGG
        totalLong = GetTheLong(Packet(0), Packet(1))
        onlygood = totalLong + 2
        extrab = 8 - ((totalLong + 2) Mod 8)
        If extrab < 8 Then
          totalLong = totalLong + extrab
        End If
    
        ReDim goodPacket(totalLong + 7)
        hbytes = UBound(Packet) + 1
        RtlMoveMemory goodPacket(6), Packet(0), (onlygood)
        goodPacket(0) = LowByteOfLong(UBound(goodPacket) - 1)
        goodPacket(1) = HighByteOfLong(UBound(goodPacket) - 1)
        pres = EncipherTibiaProtectedSP(goodPacket(0), packetKey(Index).key(0), UBound(goodPacket), UBound(packetKey(Index).key))
        thedamnCRC = GetTibiaCRC(goodPacket(6), UBound(goodPacket) - 5) ' (number of bytes - 6)
        longToBytes fourBytesCRC, thedamnCRC
        goodPacket(2) = fourBytesCRC(0)
        goodPacket(3) = fourBytesCRC(1)
        goodPacket(4) = fourBytesCRC(2)
        goodPacket(5) = fourBytesCRC(3)
        'Debug.Print "2>> " & frmMain.showAsStr(goodPacket, True) ' DEBUGGGGGGGGGGGGGGGGGGGGGGGGGG
    End If
    
    
    If (pres < 0) Then
        GiveCrackdDllErrorMessage pres, goodPacket, packetKey(Index).key, UBound(goodPacket), UBound(packetKey(Index).key), 3
        Exit Sub
    End If
    
  Else
    
      If logIt = True Then
        DebugAdd FrmMain.TxtPackets, "GAMECLIENT" & Index & ": " & showAsStr2(Packet, 0) & vbNewLine, True
      End If

    
  End If
End If
 
End Sub

Public Sub InitKEY(ByVal Index As Integer)
Dim mypid As Long
Dim strIP As String
Dim strip2 As String
Dim Packet() As Byte
Dim res As Long
Dim pres As Long


res = readTibiaKeyAtPID(Index, Tibia_Hwnd)

      If res < 0 Then
       MsgBox "WARNING: readLoginTibiaKeyAtPID failed! (this is a debug message that might be ignored)"
        Exit Sub
      End If
      

       FrmMain.TxtPacketKey.Text = "USING DECIPHER KEY = " & _
       GoodHex(packetKey(Index).key(0)) & " " & _
       GoodHex(packetKey(Index).key(1)) & " " & _
       GoodHex(packetKey(Index).key(2)) & " " & _
       GoodHex(packetKey(Index).key(3)) & " " & _
       GoodHex(packetKey(Index).key(4)) & " " & _
       GoodHex(packetKey(Index).key(5)) & " " & _
       GoodHex(packetKey(Index).key(6)) & " " & _
       GoodHex(packetKey(Index).key(7)) & " " & _
       GoodHex(packetKey(Index).key(8)) & " " & _
       GoodHex(packetKey(Index).key(9)) & " " & _
       GoodHex(packetKey(Index).key(10)) & " " & _
       GoodHex(packetKey(Index).key(11)) & " " & _
       GoodHex(packetKey(Index).key(12)) & " " & _
       GoodHex(packetKey(Index).key(13)) & " " & _
       GoodHex(packetKey(Index).key(14)) & " " & _
       GoodHex(packetKey(Index).key(15)) & vbCrLf
       

End Sub

Public Function GoodHex(B As Byte) As String
  Dim res As String
  res = Hex(B)
  If Len(res) = 1 Then
    GoodHex = "0" & res 'add a zero if VB conversion only return 1 character
  Else
    GoodHex = res
  End If
End Function

Public Function FromHexToDec(str As String) As Byte
  Dim res As Byte
  ' converts 1 character string
  ' to a byte
  res = 16 'reserved to error
  Select Case str
  Case "0"
    res = 0
  Case "1"
    res = 1
  Case "2"
    res = 2
  Case "3"
    res = 3
  Case "4"
    res = 4
  Case "5"
    res = 5
  Case "6"
    res = 6
  Case "7"
    res = 7
  Case "8"
    res = 8
  Case "9"
    res = 9
  Case "A", "a"
    res = 10
  Case "B", "b"
    res = 11
  Case "C", "c"
    res = 12
  Case "D", "d"
    res = 13
  Case "E", "e"
    res = 14
  Case "F", "f"
    res = 15
  End Select
  FromHexToDec = res
End Function
Public Function longToBytes(ByRef ByteArray() As Byte, ByVal thelong As Long) As Byte()
    CopyMemory ByteArray(0), ByVal VarPtr(thelong), Len(thelong)
End Function

Public Function ConvStrToByte(str As String) As Byte
  'converts an 1 character string into a byte
  Dim res As Byte
  Dim cad As String
  cad = "&H" & Hex(Asc(str))
  res = CLng(cad)
  ConvStrToByte = res
End Function

Public Function HighByteOfLong(address As Long) As Byte
  Dim h As Byte
  h = CByte(address \ 256) ' high byte
  HighByteOfLong = h
End Function

Public Function LowByteOfLong(address As Long) As Byte
  Dim h As Byte
  Dim l As Byte
  h = CByte(address \ 256)
  l = CByte(address - (CLng(h) * 256)) ' low byte
  LowByteOfLong = l
End Function


Public Function Hexarize2(strinput As String) As String
  Dim strByte As String
  Dim res As String
  Dim bcount As Long
  bcount = 0
  res = ""
  While Len(strinput) > 0
    strByte = Left(strinput, 1)
    strinput = Right(strinput, Len(strinput) - 1)
    res = res & GoodHex(Asc(strByte)) & " "
    bcount = bcount + 1
  Wend
  res = GoodHex(LowByteOfLong(bcount)) & " " & GoodHex(HighByteOfLong(bcount)) & " " & res
  Hexarize2 = res
End Function

Public Function FiveChrLon(num As Long) As String
  FiveChrLon = GoodHex(LowByteOfLong(num)) & " " & GoodHex(HighByteOfLong(num))
End Function

Public Function GetTheLong(byte1 As Byte, byte2 As Byte) As Long
  'get the long from 2 consecutive bytes in a tibia packet
  Dim res As Long
  res = CLng(byte2) * 256 + CLng(byte1)
  GetTheLong = res
End Function

Public Function showAsStr2(ByRef Packet() As Byte, hexad As Byte, Optional limitUbound As Long = 0) As String
  ' show a packet as string
  ' hexad:
  ' 0 -> hex with header
  ' 1 -> ascii with header
  ' 2 -> hex without header
  ' limitUbound: return result as if packet only had that ubound
  Dim i As Long
  Dim itemsNumber As Long
  Dim strShow As String
  itemsNumber = UBound(Packet)
  If limitUbound > 0 Then
    If limitUbound < itemsNumber Then
        itemsNumber = limitUbound
    End If
  End If
  
  ' depending hexad parameter, show it as hex or as ascii
  If hexad = 0 Then
     strShow = "[ hex ] "
  ElseIf hexad = 1 Then
     strShow = "[ ascii ] "
  Else
     strShow = ""
  End If
  For i = 0 To itemsNumber
   If hexad = 1 Then
     strShow = strShow & Chr(Packet(i))
   Else
     strShow = strShow & GoodHex(Packet(i)) & " "
   End If
  Next i
  showAsStr2 = strShow
End Function


Public Function HexStringToByteArray(ByRef HexString As String) As Byte()
    Dim bytOut() As Byte, bytHigh As Byte, bytLow As Byte, lngA As Long
    If LenB(HexString) Then
        ' preserve memory for output buffer
        ReDim bytOut(Len(HexString) \ 2 - 1)
        ' jump by every two characters (in this case we happen to use byte positions for greater speed)
        For lngA = 1 To LenB(HexString) Step 4
            ' get the character value and decrease by 48
            bytHigh = AscW(MidB$(HexString, lngA, 2)) - 48
            bytLow = AscW(MidB$(HexString, lngA + 2, 2)) - 48
            ' move old A - F values down even more
            If bytHigh > 9 Then bytHigh = bytHigh - 7
            If bytLow > 9 Then bytLow = bytLow - 7
            ' I guess the C equivalent of this could be like: *bytOut[++i] = (bytHigh << 8) || bytLow
            bytOut(lngA \ 4) = (bytHigh * &H10) Or bytLow
        Next lngA
        ' return the output
        HexStringToByteArray = bytOut
    End If
    End Function
