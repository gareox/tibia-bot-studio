VERSION 5.00
Begin VB.Form FrmWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2970
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "FrmWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Lst1 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Text            =   "Text"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "button"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "FrmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CountButton As Integer
Public CountText As Integer
Public CountLabel As Integer
Public CountList As Integer

Private Sub Form_Load()
    Me.Left = (FrmMain.Left + FrmMain.Width / 2) - (Me.Width / 2)
    Me.Top = (FrmMain.Top + FrmMain.Height / 2) - (Me.Height / 2)

End Sub

Private Sub Cmd1_Click(index As Integer)
Dim tmpcall As String
On Error GoTo Error:

     tmpcall = Cmd1(index).Tag
     Call Script(Me.Tag).ExecuteStatement(tmpcall)
     Exit Sub
     
Error:
     MsgBox2 Err.Description, vbCritical + vbSystemModal
End Sub


Function ConvertTwipsToPixels(lngTwips As Long, bytDirection As Byte) As Long
   
   Dim lngRetVal As Long
   
   If (bytDirection = 0) Then       'Horizontal
      lngRetVal = lngTwips / Screen.TwipsPerPixelX
   Else                            'Vertical
      lngRetVal = lngTwips / Screen.TwipsPerPixelY
   End If
   ConvertTwipsToPixels = lngRetVal

End Function


Private Sub Form_Unload(Cancel As Integer)
  Me.Visible = False
  Cancel = 1
End Sub

Private Sub Lst1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
     Lst1(index).RemoveItem Lst1(index).ListIndex
End If
End Sub
