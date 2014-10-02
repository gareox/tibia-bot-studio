VERSION 5.00
Begin VB.UserControl AddList 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   Begin TibiaStudio.UserBar UserBar 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1085
      Picture         =   "AddList.ctx":0000
      PictureHover    =   "AddList.ctx":0058
   End
End
Attribute VB_Name = "AddList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public UserBarCount As Integer
Dim iXUsername(0 To 100) As String
Private cIndex As Integer
Private cName As String

Private Sub UserBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'iForm.UserbarID = UserBar(Index).GetUsername
    cIndex = Index
End Sub

Private Sub UserControl_Initialize()
UserBarCount = 0
End Sub
Private Sub UserBar_Click(Index As Integer)

For i = 1 To UserBarCount
    If i <> Index Then
        UserBar(i).Whiteme
    End If
Next i

UserBar(Index).Blueme
cIndex = Index
'iForm.UserbarID = UserBar(Index).GetUsername
End Sub

Private Sub UserBar_DblClick(Index As Integer)
On Error Resume Next
Select Case UserBar(Index).GetVisibileStat
Case "offline"
    'iForm.CreateNewChat UserBar(Index).GetUsername, 1, False
Case "online"
    'iForm.CreateNewChat UserBar(Index).GetUsername, 2, False
Case "busy"
    'iForm.CreateNewChat UserBar(Index).GetUsername, 3, False
End Select
End Sub
Sub CloseCon()
On Error Resume Next
For i = 1 To UserBarCount
Unload UserBar(i)
Next i
UserBarCount = 0
End Sub

Function clearUsername(rUser As String)
For i = 1 To UserBarCount
If rUser = UserBar(i).GetUsername Then Unload UserBar(i)
Next i
End Function

Private Sub UserBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

For i = 0 To UserBar.Count - 1
    If i <> Index Then
        UserBar(i).Whiteme
    End If
Next i
    UserBar(Index).Blueme

If Button = 2 Then ' 1=left, 2=right
    PopupMenu FrmMain.Mnu_popup
End If


End Sub

Function GetUsername(T As Integer) As String
    GetUsername = UserBar(T).GetUsername
End Function

Function UserExist(ID As String) As Boolean
For i = 1 To UserBarCount
If LCase(ID) = LCase(UserBar(i).GetUsername) Then
UserExist = True
Exit Function
End If
Next i
UserExist = False
End Function
Function ChangeStatus(ID As String, Status As String) As String
For i = 1 To UserBarCount
If LCase(ID) = LCase(UserBar(i).GetUsername) Then
UserBar(i).ChangeStatus Status
End If
Next i
End Function
Function UserVisible(ID As String, Status As String) As String
For i = 1 To UserBarCount
If LCase(ID) = LCase(UserBar(i).GetUsername) Then
UserBar(i).UserVisible Status
End If
Next i
End Function

Sub AddUsertoBar(XUsername As String, XStatus As String)
On Error Resume Next
If UserExist(XUsername) = False Then
    i = UserBarCount + 1
UserBarCount = UserBarCount + 1
r = UserBarCount
Load UserBar(r)
iXUsername(r) = XUsername
UserBar(r).UserVisible LCase(XStatus)
UserBar(r).Top = i * 42 - 42
UserBar(r).Left = 0
UserBar(r).Tag = XUsername
UserBar(r).ChangeDisplayName XUsername
UserBar(r).ChangeStatus XStatus
UserBar(r).Visible = True
UserBar(r).Width = UserControl.Width
UserBar(r).UserIndex = r
End If
cIndex = r
cName = XUsername
End Sub

Private Sub UserControl_Resize()
For i = 1 To UserBarCount
UserBar(i).Width = UserControl.ScaleWidth
Next i
End Sub

Public Property Get UserIndex() As Integer
    UserIndex = cIndex
End Property

Public Property Get UserText() As String
    UserText = cName
End Property
