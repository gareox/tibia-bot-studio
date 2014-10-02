VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsScriptFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frm As FrmWindow
Private isDisabled As Boolean
Private isPause As Boolean
'==============================================================================================================================================
'Lectura de valores
'==============================================================================================================================================

Public Property Get hp() As Long
    hp = PlayerHP
End Property

Public Property Get mp() As Long
    mp = PlayerMP
End Property

Public Property Get isatack() As Boolean
    isatack = CBool(PlayerAttacking)
End Property

Public Property Get posx() As Long
    posx = PlayerX
End Property

Public Property Get posy() As Long
    posy = PlayerY
End Property

Public Property Get posz() As Long
    posz = PlayerZ
End Property

Public Property Get online() As Long
    online = PlayerOnline
End Property

Public Property Get iswalk() As Boolean
    iswalk = IsPlayerWalking
End Property


Public Sub atack(ByVal Monter As String)
        Attacker Monter
End Sub

Public Sub walk(ByVal X As Long, ByVal Y As Long, ByVal Z As Long)
    PlayerWalk X, Y, Z
End Sub

Public Sub say(ByVal Texto As String)
  Call PlayerSay(Texto)
End Sub

Public Sub MsgServer(ByVal Texto As String)
    MessageFromServer Texto
End Sub

Public Sub SendToServer(ByVal Bytes As Byte, ByVal lenght As Integer)
       SendPacketToServer TibiaHwnd, CByte(Bytes), lenght
End Sub

Public Sub SendToClient(ByRef Bytes As Byte, ByVal lenght As Integer)
       SendPacketToClient TibiaHwnd, Bytes, lenght
End Sub

'==============================================================================================================================================
'Funciones Generales
'==============================================================================================================================================

Public Sub printdebug(ByVal Texto As String)
   ' Call DebugAdd(FrmMain.TxtDebug, Texto, True)
End Sub

Public Sub Wait(seconds As Variant)
Dim dTimer As Double
 
    dTimer = Timer
    Do While Timer < dTimer + seconds
        DoEvents
    Loop

End Sub

Public Sub DoEvent()
    DoEvents
End Sub

Public Sub show(Visible As Boolean)
With frm
        Select Case Visible
        Case 0: .Visible = False
        Case 1: .Visible = True
        End Select
End With
End Sub

Public Function window(ByVal Tittle As String, Optional Width As Long = 4000, Optional Height As Long = 2500) As Boolean
On Error GoTo Error:

    Set frm = New FrmWindow
    
    'FrmMain.Enabled = False

    With frm
        .Caption = Tittle
        .Width = Width
        .Height = Height
        .Tag = fIndex
    End With
    
    frm.show , FrmMain
    window = True
    
    Exit Function
    
Error:
   window = False
End Function

Public Function Button(ByVal Funct As String, ByVal Caption As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
On Error GoTo Error:

With frm
   .CountButton = .CountButton + 1
   Load .Cmd1(.CountButton)

   With .Cmd1(.CountButton)
      .Caption = Caption
      .Left = X
      .Top = Y
      .Width = Width
      .Height = Height
      .Visible = True
      .Tag = Funct
   End With
   Button = .CountButton
 End With
 
  Exit Function
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Function

Public Function TextBox(ByVal Texto As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
On Error GoTo Error:

With frm
   .CountText = .CountText + 1
   Load .Txt1(.CountText)

   With .Txt1(.CountText)
      .Text = Texto
      .Left = X
      .Top = Y
      
      .Width = Width
      .Height = Height
      .Visible = True
   End With
   TextBox = .CountText
 End With
 
  Exit Function
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Function

Public Property Get Text(ByVal index As Integer) As String
On Error GoTo Error:

 With frm
   Text = .Txt1(index).Text
 End With
 
  Exit Property
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Property

Public Property Let Text(ByVal index As Integer, ByVal Texto As String)
On Error GoTo Error:

 With frm
   .Txt1(index).Text = Texto
 End With
 
  Exit Property
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Property

Public Function label(ByVal Caption As String, ByVal X As Long, ByVal Y As Long) As Long
On Error GoTo Error:

With frm
   .CountLabel = .CountLabel + 1
   Load .Lbl1(.CountLabel)

   With .Lbl1(.CountText)
      .Caption = Caption
      .Left = X
      .Top = Y
      .Visible = True
   End With
   label = .CountLabel
 End With
 
  Exit Function
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Function

Public Function Listbox(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Integer
On Error GoTo Error:

With frm
    
   .CountList = .CountList + 1
   Load .Lst1(.CountList)
   
   With .Lst1(.CountList)
      .Left = X
      .Top = Y
      
      .Width = Width
      .Height = Height
      .Visible = True
      .Tag = frm.CountList
   End With
   
   Listbox = .CountList
 End With
 
  Exit Function
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Function

Public Sub AddItem(ByVal index As Integer, ByVal Item As String)
With frm
    .Lst1(index).AddItem Item
End With
End Sub

Public Property Get ListText(ByVal index As Integer) As String
With frm
    ListText = .Lst1(index).Text
End With
End Property

Public Property Get ListCount(ByVal index As Integer) As Integer
On Error GoTo Error:
With frm
    ListCount = .Lst1(index).ListCount
End With
  Exit Property
    
Error:
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Property

Public Property Get ListIndex(ByVal index As Integer) As Integer
On Error GoTo Error:
With frm
    ListIndex = .Lst1(index).ListIndex
End With
Exit Property
Error:
 MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal
End Property

Public Property Let ListIndex(ByVal index As Integer, ByVal Valor As Integer)
On Error GoTo Error:
With frm
    .Lst1(index).ListIndex = Valor
End With
Exit Property
Error:
 MsgBox Err.Number & vbNewLine & Err.Description, vbCritical + vbSystemModal

End Property


Public Property Get readconfig(ByVal File As String, ByVal key As String) As String
    Dim sFilename As String

    sFilename = App.Path & "\Scripts\" & File & ".ini"
    readconfig = Leer_Ini(sFilename, "config", key, "")
    
End Property

Public Sub saveconfig(ByVal File As String, ByVal key As String, ByVal Valor)
    Dim sFilename As String

    sFilename = App.Path & "\Scripts\" & File & ".ini"
     Call Grabar_Ini(sFilename, "config", key, Valor)
End Sub

Public Sub playsound(ByVal WavFile As String)
    Call Reproducir_WAV(WavFile, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT)
End Sub

Public Function Path() As String
    Path = App.Path
End Function


