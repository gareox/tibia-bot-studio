VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tibia Bot Studio"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9270
   FillColor       =   &H00676767&
   ForeColor       =   &H00676767&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
   Begin TibiaStudio.AeroStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   5880
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      SimpleText      =   "Tibia Studio"
   End
   Begin VB.PictureBox PicTop 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "FrmMain.frx":164A
      ScaleHeight     =   735
      ScaleWidth      =   45000
      TabIndex        =   1
      Top             =   0
      Width           =   45000
      Begin VB.Label LblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8880
         TabIndex        =   2
         Top             =   20
         Width           =   195
      End
   End
   Begin TibiaStudio.AeroTab AeroTab 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8916
      TabCount        =   4
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabCaption(3)   =   "0 3"
      ActiveTabBackEndColor=   16777215
      ActiveTabBackStartColor=   16777215
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ActiveTabForeColor=   0
      BackColor       =   16777215
      DisabledTabBackColor=   13355721
      DisabledTabForeColor=   10526880
      InActiveTabBackEndColor=   13619151
      InActiveTabBackStartColor=   15461355
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InActiveTabForeColor=   0
      OuterBorderColor=   9800841
      TabStripBackColor=   6776679
      TabStyle        =   1
      TabOffset       =   15835
      TabCaption      =   "0"
      Begin VB.CommandButton CmdChoose 
         Caption         =   "Choose Player"
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   4500
         Width           =   1575
      End
      Begin VB.TextBox TxtSend 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -29870
         TabIndex        =   14
         Top             =   4650
         Width           =   6000
      End
      Begin VB.CommandButton CmdInject 
         Caption         =   "Inject!"
         Height          =   350
         Left            =   -31470
         TabIndex        =   12
         Top             =   4620
         Width           =   1455
      End
      Begin VB.CheckBox ChkPackets 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enabled"
         Height          =   255
         Left            =   -23750
         TabIndex        =   11
         Top             =   4680
         Width           =   975
      End
      Begin TibiaStudio.Socket Socket 
         Index           =   0
         Left            =   -31550
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin TibiaStudio.AeroGroupBox Group4 
         Height          =   4095
         Left            =   -31550
         TabIndex        =   9
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7223
         BorderColor     =   14408667
         BackColor       =   16777215
         BackColor2      =   15395562
         HeadColor1      =   -2147483633
         HeadColor2      =   15000804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Analyze Packets "
         Begin VB.TextBox TxtPacketKey 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   150
            Width           =   7335
         End
         Begin VB.TextBox TxtPackets 
            Appearance      =   0  'Flat
            Height          =   3495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   480
            Width           =   8655
         End
      End
      Begin VB.Timer T1 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   100
         Left            =   -15715
         Top             =   4320
      End
      Begin TibiaStudio.AeroGroupBox Group2 
         Height          =   4095
         Left            =   -15715
         TabIndex        =   7
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7223
         BorderColor     =   -6974058
         BackColor       =   16777215
         BackColor2      =   2829099
         HeadColor1      =   6776679
         HeadColor2      =   -1118481
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Custom Scripts "
         Begin MSComctlLib.ListView LScripts 
            Height          =   3495
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nombre"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Autor"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descripcion"
               Object.Width           =   6791
            EndProperty
         End
      End
      Begin VB.ComboBox CmdVersion 
         Height          =   315
         ItemData        =   "FrmMain.frx":30A3
         Left            =   960
         List            =   "FrmMain.frx":30A5
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4550
         Width           =   1215
      End
      Begin TibiaStudio.AeroGroupBox Group1 
         Height          =   3855
         Left            =   195
         TabIndex        =   3
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6800
         BorderColor     =   -6974058
         BackColor       =   16777215
         BackColor2      =   2829099
         HeadColor1      =   6776679
         HeadColor2      =   -1118481
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Welcome "
         Begin VB.Label LblWelcome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bienvenido a Tibia Bot Script."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2100
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version: "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   4580
         Width           =   615
      End
   End
   Begin VB.Menu MainPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnu_Config 
         Caption         =   "Configurar"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Reload 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Start 
         Caption         =   "Iniciar"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Stop 
         Caption         =   "Detener"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aCtl As Boolean

Private Sub AeroTab_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
Select Case iNewActiveTab
    Case 0: DoEvents: Group1.Visible = True: Group2.Visible = False:
    Case 1: DoEvents: Group1.Visible = False: Group2.Visible = True:
    Case 2: DoEvents:
    Case 3: DoEvents: FrmEditor.show: bCancel = True
End Select
End Sub

Private Sub CmdChoose_Click()
    Socket(0).CloseSck
    Unload_Socket
    FrmMain.Visible = False
    FrmChoose.GetClients
    FrmChoose.Visible = True
End Sub

Private Sub CmdInject_Click()
    Inject
End Sub

Private Sub Form_Load()
    Call FixThemeSupport(Me.Controls) 'add by theBatch 1.0
    LoadSizeMain
    DoEvents
    
    TabCaptions
    
    LblVersion.Caption = "Tibia: " & TibiaVersion
    
    LblWelcome.Caption = "Bienvenido a Tibia Bot Script [ " & App.Major & "." & App.Minor & "." & App.Revision & " ]" & vbNewLine & _
                        "" & vbNewLine & _
                         "Features:" & vbNewLine & _
                         "CaveBot [by AvalonTM]" & vbNewLine & _
                         "Read Send Papckets [by AvalonTM]" & vbNewLine & _
                         "Custom Scripts [by AvalonTM]" & vbNewLine & _
                         "Script Editor [by AvalonTM]" & vbNewLine & _
                         "Multi Tibia Versions [by AvalonTM]" & vbNewLine
                         

 
End Sub

Private Sub CmdVersion_Click()
    SaveVersion
End Sub

Private Sub TabCaptions()
    AeroTab.TabCaption(0) = "Main"
    AeroTab.TabCaption(1) = "Scripts"
    AeroTab.TabCaption(2) = "  Analyze Packets  "
    AeroTab.TabCaption(3) = "Editor"

End Sub


Public Sub LoadSizeMain()
On Error Resume Next
If WindowState = 2 Then
    WindowState = 0
End If
    FrmMain.Height = 6555
    FrmMain.Width = 9200
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 On Error Resume Next
 
    Dim i As Integer
    
    For i = 1 To LScripts.ListItems.Count
       If T1(i).Enabled Then
            T1(i).Enabled = False
        End If
    Next
    
    FreeLibrary hDLL
    Socket(0).CloseSck
 End
End Sub

Private Sub Form_Resize()
On Error Resume Next

LblVersion.Left = FrmMain.Width - LblVersion.Width - 300
'AeroTab.Width = FrmMain.Width
'AeroTab.Height = FrmMain.Height - StatusBar.Height

End Sub

Private Sub LScripts_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Item.Checked Then
   Item.Checked = False
Else
   Item.Checked = True
End If
End Sub

Private Sub LScripts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim index As Integer
Dim Count As Integer
DoEvents
If Button = 2 Then
       Count = LScripts.ListItems.Count
    If Count <> 0 Then
        index = LScripts.SelectedItem.index
        
        For i = 1 To Count
            DoEvents
            Mnu_Start(i).Visible = False
            Mnu_Stop(i).Visible = False
        Next
            DoEvents
            Mnu_Start(index).Visible = True
            Mnu_Stop(index).Visible = True
        PopupMenu MainPopup
    End If
End If
End Sub

Private Sub mnu_Config_Click()
On Error GoTo Error:
DoEvents
    If LScripts.ListItems.Count <> 0 Then
       fIndex = LScripts.SelectedItem.index

       Call Script(fIndex).ExecuteStatement("config")
       Exit Sub
    End If
    
Error:
    MsgBox2 "El Script no contiene configuracion", vbInformation + vbSystemModal
End Sub

Private Sub Mnu_Reload_Click()
Dim index As Integer
On Error GoTo Error:
  DoEvents
    If LScripts.ListItems.Count <> 0 Then
          index = LScripts.SelectedItem.index
        If Not Active(index) Then
            Me.Enabled = False
            Screen.MousePointer = 11
            Reset_ObjScript index
            Load_ObjSCript index
            Call Load_Scripts(index, LScripts.ListItems.Item(index).Text)
            Me.Enabled = True
            Screen.MousePointer = 0
            MsgBox "Script se a cargado correctamente!", vbInformation + vbSystemModal
            Else
            MsgBox "No se puede Recargar el script mientras este en ejecucion.", vbInformation + vbSystemModal
        End If
    End If
    
    Exit Sub
Error:
    Me.Enabled = True
    Screen.MousePointer = 0
    MsgBox2 "Error al Recargar el Script", vbCritical + vbSystemModal
End Sub

Private Sub Mnu_Start_Click(index As Integer)
Dim cIndex As Integer
On Error GoTo Error:
DoEvents
    If LScripts.ListItems.Count <> 0 Then
        cIndex = LScripts.SelectedItem.index
       If Not Active(cIndex) Then
            T1(cIndex).Enabled = True
            Mnu_Start(index).Enabled = False
            Mnu_Stop(index).Enabled = True
            Active(cIndex) = True
            LScripts.SelectedItem.Checked = True
            Else
            MsgBox "No se puede iniciar el script.", vbExclamation + vbSystemModal
       End If
    End If
    
    Exit Sub
Error:
    MsgBox2 "El Script no contiene funcion main", vbInformation + vbSystemModal
End Sub

Private Sub Mnu_Stop_Click(index As Integer)
Dim cIndex As Integer
On Error GoTo Error:
 DoEvents
    If LScripts.ListItems.Count <> 0 Then
        cIndex = LScripts.SelectedItem.index
       If Active(cIndex) Then
            T1(cIndex).Enabled = False
            Mnu_Start(index).Enabled = True
            Mnu_Stop(index).Enabled = False
            Active(cIndex) = False
            LScripts.SelectedItem.Checked = False
            Else
            MsgBox "No se puede terminar el script por que no esta en ejecucion.", vbExclamation + vbSystemModal
       End If
    End If
    
    Exit Sub
Error:
    MsgBox2 "El Script no contiene funcion terminate", vbInformation + vbSystemModal
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
     AcceptConnection index, requestID
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
If IsConnected(index) Then
    InitKEY index
    IncomingData index, bytesTotal
 End If
End Sub

Private Sub T1_Timer(index As Integer)
DoEvents
    Call ScriptExecute(index)
End Sub

Private Sub TxtPackets_KeyPress(KeyAscii As Integer)

If KeyAscii = 3 Then
' empty the clipboard
    Clipboard.Clear
    ' and copy the selected text to the clipboard
    Clipboard.SetText TxtPackets.SelText
End If

 KeyAscii = 0
End Sub



Private Sub TxtSend_KeyPress(KeyAscii As Integer)
Dim Packet() As Byte

On Error GoTo Error:

If KeyAscii = 13 Then
   If TxtSend.Text = "" Then Exit Sub
   HexStrToByteArray TxtSend, Packet
   SendPacketToServer TibiaHwnd, Packet(0), UBound(Packet)
   KeyAscii = 0
   TxtSend.Text = ""
   StatusBar.SimpleText = "packet sent"
   Wait 1
   StatusBar.SimpleText = "Tibia Studio"
End If
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub
