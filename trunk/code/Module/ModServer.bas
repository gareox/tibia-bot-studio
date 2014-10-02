Attribute VB_Name = "ModServer"
Option Explicit

Public Const MAX_CLIENTS = 5
Public Puerto As Long

Public Sub SetPort(ByVal port As Long)
    Puerto = port
End Sub

Public Sub InitSever()
Dim i As Integer

    SetPort Puerto
    
 With FrmMain
    .Socket(0).RemoteHost = .Socket(0).LocalIP
    .Socket(0).LocalPort = Puerto
    
    .Socket(0).Listen
 
 End With
 
    FrmMain.TxtPackets.Text = ""
   'Logs
    DebugAdd FrmMain.TxtPackets, "LOG >> Listening Ip: " & FrmMain.Socket(0).RemoteHost, True
    DebugAdd FrmMain.TxtPackets, "LOG >> Listening Port: " & Puerto, True
    DebugAdd FrmMain.TxtPackets, "LOG >> TibiaStudio Iniciado.", True

End Sub

Public Sub Load_Socket()
Dim i As Integer
    'Descargar Sockets
    For i = 1 To MAX_CLIENTS
        Load FrmMain.Socket(i)
    Next i

End Sub

Public Sub Unload_Socket()
Dim i As Integer
    'Descargar Sockets
    For i = 1 To MAX_CLIENTS
        Unload FrmMain.Socket(i)
    Next i

End Sub

Public Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Packet() As Byte
    FrmMain.Socket(index).GetData Packet, vbArray + vbByte
   
   If FrmMain.ChkPackets.Value = 1 Then
    ReadPacket index, Packet
   End If
   
End Sub


Public Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot
        If i <> 0 Then
            ' Whoho, we can connect them
            FrmMain.Socket(i).CloseSck
            FrmMain.Socket(i).Accept SocketId
            DebugAdd FrmMain.TxtPackets, "CLLIENT CONECTED: " & i, True
        End If
    End If

End Sub

Public Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    For i = 1 To MAX_CLIENTS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit For
        End If
    Next i

End Function

Public Function IsConnected(ByVal index As Long) As Boolean
    IsConnected = FrmMain.Socket(index).State = sckConnected
End Function

Public Function GetClientIP(ByVal index As Long) As String
    GetClientIP = FrmMain.Socket(index).RemoteHostIP
End Function
