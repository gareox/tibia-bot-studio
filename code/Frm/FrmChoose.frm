VERSION 5.00
Begin VB.Form FrmChoose 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Client Select"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   Icon            =   "FrmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox CmdVersion 
      Height          =   315
      ItemData        =   "FrmChoose.frx":164A
      Left            =   840
      List            =   "FrmChoose.frx":164C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1410
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1470
      Width           =   615
   End
End
Attribute VB_Name = "FrmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isload  As Boolean

Private Sub CmdVersion_Click()
    If Not isload Then Exit Sub
    SaveVersion2
     FrmMain.CmdVersion.ListIndex = FrmChoose.CmdVersion.ListIndex
    GetClients
End Sub

Private Sub Command2_Click()
    GetClients
End Sub

Private Sub Load_tVersions()
Dim i As Integer
    CmdVersion.Clear
    For i = 1 To UBound(Buffver)
        FrmChoose.CmdVersion.AddItem Buffver(i)
    Next
    FrmChoose.CmdVersion.ListIndex = FrmMain.CmdVersion.ListIndex
End Sub
Private Sub Form_Load()
    Load_tVersions
    GetClients
    isload = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub List1_DblClick()

If List1.ListIndex <> -1 Or List1.Text <> "" Then
    List2.ListIndex = List1.ListIndex
    List3.ListIndex = List1.ListIndex
    
    ClientSelected = List2.Text
    TibiaHwnd = FindProcess(Tibia_Hwnd)
    
    myID = List3.Text
    myBpos = MyBattleListPosition
    
    FrmMain.Caption = "Tibia Bot Studio - [ " & List1.Text & " ]"
    FrmMain.show
    Load_Socket
    InitSever
    FrmChoose.Visible = False
Else
    MsgBox "No hay ningun cliente seleccionado."
End If

End Sub

Public Sub GetClients()
Dim Thwnd As Long
Dim TmptheBase As Long
Dim TmptheOffset As Long
Dim i As Integer
Dim tmpName As String

List1.Clear
List2.Clear
List3.Clear

For i = 1 To CountTibiaWindows

    Thwnd = FindWindowEx(0&, Thwnd, Proseso, vbNullString)
    If useDynamicOffset = "yes" Then
        TmptheBase = getProcessBase(Thwnd, tibiaModuleRegionSize, True)
        TmptheOffset = TmptheBase - &H400000
    End If

    If Thwnd <> 0& Then

         For BattleList_Address = BattleList_Start To BattleList_End Step CharDist
            If Memory_ReadLong(Thwnd, TmptheOffset + Player_ID) = Memory_ReadLong(Thwnd, TmptheOffset + BattleList_Address + distance.cID) Then
               tmpName = (Memory_ReadString(Thwnd, TmptheOffset + BattleList_Address + distance.Name))
                
                If Asc(Mid(tmpName, 1, 1)) <> 0 Then    'Si no esta conectado no lo mostramos en la lista.
                    List1.AddItem tmpName
                    List2.AddItem Thwnd
                    List3.AddItem Memory_ReadLong(Thwnd, TmptheOffset + Player_ID)    'PlayerPos
                    Exit For
                End If
                
            End If
        Next BattleList_Address
    Else
        MsgBox "Cliente no encontrado", vbInformation
        End
    End If
 Next


End Sub

