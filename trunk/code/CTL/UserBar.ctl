VERSION 5.00
Begin VB.UserControl UserBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   MouseIcon       =   "UserBar.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   840
   End
   Begin VB.PictureBox BusyPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1920
      Picture         =   "UserBar.ctx":030A
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox OnlinePic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1320
      Picture         =   "UserBar.ctx":111E
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox offlinePic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   720
      Picture         =   "UserBar.ctx":1F32
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox MyPictureGhab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      MouseIcon       =   "UserBar.ctx":2D46
      MousePointer    =   99  'Custom
      Picture         =   "UserBar.ctx":3050
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   4
      Top             =   0
      Width           =   630
      Begin VB.Image MyPicture 
         Height          =   510
         Left            =   60
         MouseIcon       =   "UserBar.ctx":4594
         MousePointer    =   99  'Custom
         Picture         =   "UserBar.ctx":489E
         Stretch         =   -1  'True
         Top             =   60
         Width           =   510
      End
   End
   Begin VB.Label iDisplayName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      MouseIcon       =   "UserBar.ctx":56B2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   75
      Width           =   135
   End
   Begin VB.Label iStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...."
      Height          =   195
      Left            =   720
      MouseIcon       =   "UserBar.ctx":59BC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   360
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   0
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   0
      Width           =   510
   End
   Begin VB.Label LblIndex 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "UserBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public iURL As String
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim m_Picture As StdPicture
Dim m_PictureHover As StdPicture
Dim m_PictureDown As StdPicture
Dim myUserVisible As String
Dim isOver As Boolean
Dim m_State As Integer
Dim imMaster As Boolean
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()
Event MousePayin()


Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(m_New_Picture As StdPicture)
    Set m_Picture = m_New_Picture
    DrawState
    PropertyChanged "Picture"
End Property

Public Property Get PictureHover() As StdPicture
    Set PictureHover = m_PictureHover
End Property

Public Property Set PictureHover(m_New_PictureHover As StdPicture)
    Set m_PictureHover = m_New_PictureHover
    DrawState
    PropertyChanged "PictureHover"
End Property

Public Property Get PictureDown() As StdPicture
    Set PictureDown = m_PictureDown
End Property

Public Property Set PictureDown(m_New_PictureDown As StdPicture)
    Set m_PictureDown = m_New_PictureDown
    DrawState
    PropertyChanged "PictureDown"
End Property

Public Sub Blueme()
imMaster = True
Image1.Visible = False
End Sub
Public Sub Whiteme()
imMaster = False
Image1.Visible = True
End Sub

Public Function GetMasterStatus() As Boolean
GetMasterStatus = imMaster
End Function
Private Sub DrawState()
    On Error Resume Next
    If m_State = 1 Then 'mouse hover
    If Not imMaster = True Then
        UserControl.Cls
        UserControl.PaintPicture m_PictureHover, 0, 0
        'Image1.Visible = True
        Image1.Picture = m_PictureHover
        UserControl.Width = Image1.Width
    End If
    ElseIf m_State = 2 Then 'mouse down
        UserControl.Cls
        'UserControl.PaintPicture m_PictureDown, 0, 0
        'Image1.Picture = m_PictureDown
        'Image1.Visible = False
        UserControl.BackColor = &HFFE495
        'UserControl.Width = Image1.Width
    Else
        UserControl.Cls
        UserControl.PaintPicture m_Picture, 0, 0
        'Image1.Visible = True
        Image1.Picture = m_Picture
        UserControl.Width = Image1.Width
    End If
End Sub

Private Sub iDisplayName_Click()
RaiseEvent Click
End Sub

Private Sub iDisplayName_DblClick()
RaiseEvent DblClick
End Sub

Private Sub iDisplayName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub iDisplayName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub



Private Sub Image1_Click()
RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub

Private Sub iStatus_Click()
RaiseEvent Click
End Sub

Private Sub iStatus_DblClick()
RaiseEvent DblClick
End Sub

Private Sub iStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not iURL = "" Then GotoSite iURL
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub iStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub

Private Sub MyPicture_DblClick()
RaiseEvent DblClick
End Sub

Private Sub MyPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub

Private Sub MyPictureGhab_DblClick()
RaiseEvent DblClick
End Sub

Private Sub MyPictureGhab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        Call DrawState
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
UserControl.Height = 630
imMaster = False
End Sub

Private Sub UserControl_InitProperties()
    Set m_Picture = Nothing
    Set m_PictureHover = Nothing
    Set m_PictureDown = Nothing
End Sub

Private Sub UserControl_MousePayin()
RaiseEvent MousePayin
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
                  '  UserControl.BackColor = vbBlue
        m_State = 2
        Call DrawState

    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call iMove(Button, Shift, X, Y)
End Sub
Sub iMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                Call DrawState

            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                Call DrawState
                isOver = False
            End If
        End If
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    Call DrawState
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_PictureHover = .ReadProperty("PictureHover", Nothing)
        Set m_PictureDown = .ReadProperty("PictureDown", Nothing)
    End With
End Sub

Private Sub UserControl_Resize()
    Image1.Width = UserControl.Width
    DrawState
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Picture", m_Picture, Nothing)
        Call .WriteProperty("PictureHover", m_PictureHover, Nothing)
        Call .WriteProperty("PictureDown", m_PictureDown, Nothing)
    End With
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hwnd)
End Function

Public Sub ChangeDisplayName(Displayname As String)
iDisplayName.Caption = Displayname
End Sub
Public Function UserVisible(iUserVisible As String) As String
UserVisible = iUserVisible
myUserVisible = iUserVisible
Select Case iUserVisible
Case "online"
MyPicture.Picture = OnlinePic.Picture
Case "offline"
MyPicture.Picture = offlinePic.Picture
Case "busy"
MyPicture.Picture = BusyPic.Picture
End Select

End Function

Public Sub ChangeStatus(Status As String)
iStatus.Caption = Status
If InStr(Status, "www.") Or InStr(Status, "http://") Then
iStatus.ForeColor = &HFB8400
iStatus.MousePointer = 99
SeperateURL Status
If Not Status = "Offline" And Not Status = "Add Request Sent..." Then
UserVisible ("online")
Else
UserVisible ("offline")
End If
Else
If InStr(LCase(iStatus), " busy ") Or InStr(LCase(iStatus), " busy") Or InStr(LCase(iStatus), "busy ") Then
iStatus.ForeColor = vbRed
iStatus.MousePointer = 0
UserVisible ("busy")
iURL = ""
Else
iStatus.ForeColor = vbBlack
iStatus.MousePointer = 0
iURL = ""
If Not Status = "Offline" And Not Status = "Add Request Sent..." Then
UserVisible ("online")
Else
UserVisible ("offline")
End If
End If
End If
End Sub
Function GetUsername() As String
GetUsername = iDisplayName.Caption
End Function
Function SeperateURL(MixedStatus As String) As String
Dim iCount As Long, FirstW As Long
iCount = 0
For i = 1 To Len(MixedStatus)
If LCase(Mid(MixedStatus, i, 1)) = "w" Then
iCount = iCount + 1
If iCount = 3 Then
FirstW = i - 3
iURL = FindUntilSpace(MixedStatus, FirstW)
End If
End If
Next i
End Function

Function FindUntilSpace(Xtring As String, FirstW As Long) As String
Xtring = Right(Xtring, Len(Xtring) - FirstW)
For i = 1 To Len(Xtring)
If LCase(Mid(Xtring, i, 1)) = " " Then
FindUntilSpace = Left(Xtring, i - 1)
'iStatus.ToolTipText = iStatus.Caption
'iStatus = Replace(iStatus, FindUntilSpace, "")
Exit Function
End If
Next i
For i = 1 To Len(Xtring)
If i = Len(Xtring) Then
FindUntilSpace = Left(Xtring, i)
'iStatus.ToolTipText = iStatus.Caption
'iStatus = Replace(iStatus, FindUntilSpace, "")
Exit Function
End If
Next i
End Function

Function GetVisibileStat() As String
GetVisibileStat = myUserVisible
End Function

Public Property Get UserIndex() As Integer
    UserIndex = LblIndex.Caption
End Property

Public Property Let UserIndex(valor As Integer)
    LblIndex.Caption = valor
End Property

Public Property Get UserText() As Integer
    UserText = iDisplayName.Caption
End Property
