Attribute VB_Name = "ModMain"
Option Explicit

Sub Main()

    InitCommonControls
    
    ReDim loginPacketKey(1 To MAX_CLIENTS)
    ReDim packetKey(1 To MAX_CLIENTS)
    ReDim ConnectionBuffer(1 To MAX_CLIENTS)
    Set ProcessidIPrelations = New Scripting.Dictionary
    
    
    PreLoadConfig
    LoadFileSCripts
    
    Puerto = 21215
    FrmChoose.show
    
End Sub

