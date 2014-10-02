Attribute VB_Name = "ModConfig"
Option Explicit

Public nDllVersion As String
Public TibiaVersion As String
Public ConfigPath As String
Public ConfigVersion As String
Private TmpVersion As String
Private vCount As Integer
Public Buffver() As String


Public Sub PreLoadConfig()
Dim sFilename As String

Dim i As Integer


    sFilename = App.Path & "\config.ini"
    ConfigVersion = Leer_Ini(sFilename, "tibiastudio", "configPath", "986")
    TibiaVersion = Leer_Ini(sFilename, "tibiastudio", "TibiaVersion", "9.86")
    nDllVersion = Leer_Ini(sFilename, "tibiastudio", "NewDllVersion", "10.57")
    TibiaVersionLong = Replace(TibiaVersion, ".", "")
    nDllVersion = Replace(nDllVersion, ".", "")

    Load_DllVersion
    AddVersion
    GetConfigVersion TibiaVersion
    LoadConfig ConfigVersion
End Sub

Public Sub Load_DllVersion()
On Error GoTo Error:
    If TibiaVersionLong < nDllVersion Then
        DllVersion = False
    Else
        DllVersion = True
    End If
    
 Exit Sub
Error:
  MsgBox Err.Description, vbCritical + vbSystemModal
  End
End Sub


Public Sub AddVersion()
Dim sFilename As String
Dim i As Integer
Dim tmpver As String
Dim tmpcount As String

    sFilename = App.Path & "\config.ini"
    
    vCount = Leer_Ini(sFilename, "versions", "count", "0")
     
    For i = 1 To vCount
        tmpcount = i
        tmpver = Leer_Ini(sFilename, "versions", tmpcount, "0")
        ReDim Preserve Buffver(i)
        Buffver(i) = tmpver
        FrmMain.CmdVersion.AddItem tmpver

    Next
    
    
End Sub

Public Sub GetConfigVersion(ByVal vTibia As String)
Dim i As Integer

    For i = 1 To vCount
           FrmMain.CmdVersion.ListIndex = i - 1
        If FrmMain.CmdVersion.Text = vTibia Then
            FrmChoose.CmdVersion.ListIndex = FrmMain.CmdVersion.ListIndex
            Exit For
        End If
    Next
End Sub

Public Sub LoadConfig(ByVal File As String)
Dim sFilename As String
Dim TmpUseCrackd As String

On Error GoTo Error1:

    
    sFilename = App.Path & "\configs\" & File & ".ini"
    
    If Not FileExists(sFilename) Then
        GetConfigVersion TmpVersion
        Exit Sub
    End If
    
    'Address de TibiaSock
    Address_OUTGOINGDATASTREAM CLng(Leer_Ini(sFilename, "tibiasock", "OUTGOINGDATASTREAM", "&H7C60D8"))
    Address_OUTGOINGDATALEN CLng(Leer_Ini(sFilename, "tibiasock", "OUTGOINGDATALEN", "&H9E1B98"))
    Address_SENDOUTGOINGPACKET CLng(Leer_Ini(sFilename, "tibiasock", "SENDOUTGOINGPACKET", "&H51D3D0"))
    
    Address_INCOMINGDATASTREAM CLng(Leer_Ini(sFilename, "tibiasock", "INCOMINGDATASTREAM ", "&H9E1B84"))
    Address_PARSERFUNC CLng(Leer_Ini(sFilename, "tibiasock", "PARSERFUNC", "&H46E0E0"))
    

Error1:
On Error GoTo Error2:

    'Address de TibiaClient
    tibiaModuleRegionSize = Leer_Ini(sFilename, "tibia", "tibiaModuleRegionSize", "&H2FC000")
    adrConnectionKey = Leer_Ini(sFilename, "tibia", "adrConnectionKey", "&H7B7150")
    adrXOR = Leer_Ini(sFilename, "tibia", "adrXOR", "&H7BF1F0")

    Player_ID = Leer_Ini(sFilename, "tibia", "adrNum", "&H98AEA4")
    Player_HP = Leer_Ini(sFilename, "tibia", "adrMyHP", "&H953000")
    Player_MP = Leer_Ini(sFilename, "tibia", "adrMyMana", "&H7BF244")
    
    BattleList_Start = Leer_Ini(sFilename, "tibia", "adrNameStart", "&H953008")
    adrNChar = Leer_Ini(sFilename, "tibia", "adrNChar", "&H953008")

    AdrXPos = Leer_Ini(sFilename, "tibia", "adrXPos", "&H98AEA8")
    AdrYPos = Leer_Ini(sFilename, "tibia", "adrYPos", "&H98AEAC")
    AdrZPos = Leer_Ini(sFilename, "tibia", "adrZPos", "&H98AEB0")

    GoToX = Leer_Ini(sFilename, "tibia", "adrXgo", "&H98AEA0")
    GoToY = Leer_Ini(sFilename, "tibia", "adrYgo", "&H98AE98")
    GoToZ = Leer_Ini(sFilename, "tibia", "adrZgo", "&H953004")
    
    
    IsConeted = Leer_Ini(sFilename, "tibia", "adrConnected", "&H7C8FF8")
    RedSquare = Leer_Ini(sFilename, "tibia", "RedSquare", "&H7BF240")
    
    adrGo = Leer_Ini(sFilename, "tibia", "adrGo", "&H953058")
 
    
    LAST_BATTLELISTPOS = Leer_Ini(sFilename, "tibia", "LAST_BATTLELISTPOS", "1299")
    
    CharDist = Leer_Ini(sFilename, "tibia", "CharDist", "&A0")
    MaxCreatures = Leer_Ini(sFilename, "tibia", "MaxCreatures", "187")
    Distance_Characters = Leer_Ini(sFilename, "tibia", "CharDist", "176")
    
    useDynamicOffset = Leer_Ini(sFilename, "MemoryAddresses", "useDynamicOffset", "yes")
    
    TmpUseCrackd = Leer_Ini(sFilename, "MemoryAddresses", "UseCrackd", "yes")
    
    
    If TmpUseCrackd = "yes" Then
        UseCrackd = True
    Else
        UseCrackd = False
    End If
    
     BattleList_End = BattleList_Start + (CharDist * MaxCreatures)
     Blatte_First = BattleList_Start - 4
     
    
Exit Sub

Error2:
 
 MsgBox "Error al cargar la configuracion:" & vbNewLine & sFilename, vbCritical + vbSystemModal
 End
End Sub

Public Sub SaveVersion()
Dim sFilename As String


    TmpVersion = TibiaVersion

    sFilename = App.Path & "\config.ini"
    
    Call Grabar_Ini(sFilename, "tibiastudio", "TibiaVersion", FrmMain.CmdVersion.Text)
    Call Grabar_Ini(sFilename, "tibiastudio", "configPath", "config" & Replace(FrmMain.CmdVersion.Text, ".", ""))
    
    ConfigVersion = Leer_Ini(sFilename, "tibiastudio", "configPath", "986")
    TibiaVersion = Leer_Ini(sFilename, "tibiastudio", "TibiaVersion", "9.86")
    
    LoadConfig ConfigVersion
End Sub

Public Sub SaveVersion2()
Dim sFilename As String


    TmpVersion = TibiaVersion

    sFilename = App.Path & "\config.ini"
    
    Call Grabar_Ini(sFilename, "tibiastudio", "TibiaVersion", FrmChoose.CmdVersion.Text)
    Call Grabar_Ini(sFilename, "tibiastudio", "configPath", "config" & Replace(FrmChoose.CmdVersion.Text, ".", ""))
    
    ConfigVersion = Leer_Ini(sFilename, "tibiastudio", "configPath", "986")
    TibiaVersion = Leer_Ini(sFilename, "tibiastudio", "TibiaVersion", "9.86")
    
    LoadConfig ConfigVersion
End Sub
