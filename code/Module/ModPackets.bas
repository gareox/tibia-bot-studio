Attribute VB_Name = "ModPackets"
 Option Explicit

Public Sub PlayerSay(Text As String)
    Dim Packet() As Byte
    Dim X As Integer

    ReDim Packet(3)
    
    Packet(0) = &H96
    Packet(1) = &H1 '1=say, 2=whisper, 3=yell
    Packet(2) = LongToByte(Len(Text), 1) 'if it doesn't work, remove the "LongToByte" function from there
    Packet(3) = LongToByte(Len(Text), 2) 'and put there 0
    
    For X = 1 To Len(Text)
        ReDim Preserve Packet(X + 4)
        Packet(X + 3) = Asc(Mid(Text, X, 1))
    Next X
    

    SendPacketToServer TibiaHwnd, Packet(0), UBound(Packet)
 End Sub
    
Public Sub PlayerMove()
    Dim Packet(2) As Byte

       Packet(0) = &H64
       Packet(1) = &H3
       Packet(2) = &H5
     
    SendPacketToServer TibiaHwnd, Packet(0), UBound(Packet)
End Sub


Public Sub MessageFromServer(ByVal Text As String)
 Dim Packet() As Byte
 Dim X As Integer
 
        ReDim Preserve Packet(2)
        Packet(0) = &HB4
        Packet(1) = &H11

        For X = 1 To Len(Text)
        ReDim Preserve Packet(X + 2)
            Packet(X + 1) = Asc(Mid(Text, X, 1))
        Next X
        
       SendPacketToClient TibiaHwnd, Packet(0), UBound(Packet)
End Sub


Public Sub PlayerAtack(ID As Long)
    Dim Packet(5) As Byte

        Packet(0) = &HA1
        Packet(1) = LongToByte(ID, 1)
        Packet(2) = LongToByte(ID, 2)
        Packet(3) = LongToByte(ID, 3)
        Packet(4) = LongToByte(ID, 4)

       SendPacketToServer TibiaHwnd, Packet(0), UBound(Packet)
 
End Sub

Public Sub ChangeOutfit(OutfitID As Long, HeadColor As Byte, BodyColor As Byte, LegsColor As Byte, FeetColor As Byte, Addons As Byte)
  Dim Packet(8) As Byte
   
        Packet(0) = &HD3
        Packet(1) = LongToByte(OutfitID, 1)
        Packet(2) = LongToByte(OutfitID, 2)
        Packet(3) = HeadColor
        Packet(4) = BodyColor
        Packet(5) = LegsColor
        Packet(6) = FeetColor
        Packet(7) = Addons 'Nothing = 0, 1 = first, 2 = second, 3 = both

        
      SendPacketToServer TibiaHwnd, Packet(0), UBound(Packet)
 
End Sub

