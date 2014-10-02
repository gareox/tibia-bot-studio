VERSION 5.00
Begin VB.UserControl ctlTab 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   300
   ScaleWidth      =   1515
   Begin VB.Line linBorderTop 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Caption]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   60
      Width           =   780
   End
   Begin VB.Line linBorderRight 
      BorderColor     =   &H80000015&
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   300
   End
   Begin VB.Line linBorderLeft 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Image imgIcon 
      Height          =   255
      Left            =   60
      Top             =   30
      Width           =   255
   End
End
Attribute VB_Name = "ctlTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Description]
'EBTabStrip tab control (Private)

Option Explicit

Private flgIsActive         As Boolean      'Current draw state of tab (front=true)

Event Click()

Private Sub imgIcon_Click()

'[Description]
'   Raise the click event for this tab

'[Code]

    RaiseEvent Click
    
End Sub

Private Sub lblCaption_Click()

'[Description]
'   Raise the click event for this tab

'[Code]

    RaiseEvent Click
    
End Sub

Private Sub UserControl_Click()

'[Description]
'   Raise the click event for this tab

'[Code]

    RaiseEvent Click
    
End Sub

Private Sub UserControl_InitProperties()

'[Description]
'   Initialise the control's properties

'[Code]

    IsActive = False
    Caption = "[No Caption]"
    Set Image = Nothing
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'[Description]
'   Retrieve the controls properties from it's PropBag

'[Code]

    With PropBag
        IsActive = .ReadProperty("IsActive", False)
        Caption = .ReadProperty("Caption", "[No Caption]")
        Set Image = .ReadProperty("Image", Nothing)
    End With
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Resize the controls constiuant controls

'[Code]

    With UserControl
    
        If imgIcon.Picture = 0 Then
            'no picture - lable flush with left of tab
            lblCaption.left = 60
            .Width = lblCaption.left + lblCaption.Width + 120
        Else
            'leave space for image
            lblCaption.left = 360
            .Width = lblCaption.left + lblCaption.Width + 120
        End If
        
        linBorderLeft.Y2 = .Height - 15
        linBorderTop.X2 = .Width - 15
        linBorderRight.X1 = .Width - 15
        linBorderRight.X2 = .Width - 15
        linBorderRight.Y2 = .Height - 0.15
    End With
    
End Sub

Public Property Get IsActive() As Boolean

'[Description]
'   Returns the current draw state of the tab
'   True:   Control is currenly in Front mode
'   False:  Controls is current in Back mode

'[Code]

    IsActive = flgIsActive
    
End Property

Public Property Let IsActive(NewValue As Boolean)

'[Description]
'   Sets the current draw style for the tab

'[Code]

    flgIsActive = NewValue
    
    Select Case NewValue
        
        Case True
            'Draw tab in front mode
            lblCaption.ForeColor = vb3DHighlight
            UserControl.BackColor = vb3DFace
            linBorderLeft.BorderColor = vb3DHighlight
            linBorderTop.BorderColor = vb3DHighlight
            linBorderRight.BorderColor = vb3DDKShadow
            linBorderTop.BorderWidth = 1
            UserControl.Height = 315
        
        Case False
            'Draw tab in back mode
            lblCaption.ForeColor = vbButtonText
            UserControl.BackColor = vb3DShadow
            linBorderLeft.BorderColor = Ambient.BackColor
            linBorderTop.BorderColor = Ambient.BackColor
            linBorderRight.BorderColor = Ambient.BackColor
            linBorderTop.BorderWidth = 4
            UserControl.Height = 300
            
    End Select
    
    PropertyChanged "IsActive"
    
End Property

Public Property Get Caption() As String

'[Description]
'   Return the tabs caption

'[Code]

    Caption = lblCaption.Caption
    
End Property

Public Property Let Caption(NewValue As String)

'[Description]
'   Set the tabs caption

'[Code]

    lblCaption.Caption = NewValue
    
    UserControl_Resize
    
    PropertyChanged "Caption"
    
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'[Description]
'   Store the controls properties in it's PropBag

'[Code]

    With PropBag
        .WriteProperty "IsActive", flgIsActive, False
        .WriteProperty "Caption", lblCaption.Caption, "[No Caption]"
        .WriteProperty "Image", imgIcon.Picture, Nothing
    End With
    
End Sub

Public Property Get Image() As StdPicture

'[Description]
'   Returns the tab's Image

'[Code]

    Set Image = imgIcon.Picture
    
End Property

Public Property Set Image(NewValue As StdPicture)

'[Description]
'   Set the tabs image

'[Code]

    Set imgIcon.Picture = NewValue
    UserControl_Resize
    
    PropertyChanged "Image"
    
End Property

