VERSION 5.00
Begin VB.UserControl ctlEBTabStrip 
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ControlContainer=   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   3675
   ToolboxBitmap   =   "ctlEBTabStrip.ctx":0000
   Begin EBTabStrip.ctlTab ctlTab 
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      IsActive        =   -1  'True
   End
   Begin VB.Line linBorderLeft 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   0
      Y1              =   300
      Y2              =   1860
   End
   Begin VB.Line linBorderBottom 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   3660
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line linBorderRight 
      BorderColor     =   &H80000010&
      X1              =   3660
      X2              =   3660
      Y1              =   300
      Y2              =   1860
   End
   Begin VB.Line linBorderTop 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   3660
      Y1              =   300
      Y2              =   300
   End
End
Attribute VB_Name = "ctlEBTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'[Description]
'   Tabbed dialog control

'[Notes]
'   Contained Controls:
'   Even though this control can contain other controls, it cannot manage
'   controls on multiple tabs (version 2?) like the Tabbed Dialog control.
'   The ability to contain other controls in this version is to host the page
'   containers with the control to make it easier to move the control about
'   on its host and keep the paged containers together.

'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@earlybirdmarketing.com

'[History]
'   V1.0.0  18/06/2001
'   Initial release

'[Declarations]
Option Explicit

'Property Storage
Private intCurrentTab       As Integer      'Currently selected tab

'Events
Event TabClick(NewTabIndex As Integer, OldTabIndex As Integer)

'Error enums
Private Enum EBTabStrip_Error
    ERR_MINTAB
    ERR_NOSUCHTAB
End Enum

Private Sub ctlTab_Click(Index As Integer)

'[Description]
'   If the user has clicked a new tab:
'   * Bring the new tab to the front
'   * Send the old tab to the back
'   * Raise a TabClick event

'[Code]

    If Index <> intCurrentTab Then
        'Bring selected tab to front
        ctlTab(Index).IsActive = True
        ctlTab(Index).ZOrder 0
        
        'Send previously selected tab to back
        'Suppress any error as we may be referencing a tab which has been
        'removed
        On Error Resume Next
        ctlTab(intCurrentTab).IsActive = False
        On Error GoTo 0
        
        RaiseEvent TabClick(Index, intCurrentTab)
    
        intCurrentTab = Index
    End If
    
End Sub

Private Sub UserControl_InitProperties()

'[Description]
'   Initialise the controls properties

'[Code]

    Tabs = 3
    CurrentTab = 0
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'[Description]
'   Restore the control properties from the PropBag

'[Notes]
'   As there can be multiple caption and image properties (one for each tab),
'   we cycle through the tabs and use the Property Name & Index as a property
'   key

'[Declarations]
Dim intIndex                As Integer      'Used to cycle through the tabs

'[Code]

    With PropBag
        Tabs = .ReadProperty("Tabs", 3)
        
        'Cycle through each tab setting its properties
        For intIndex = 0 To ctlTab.Count - 1
            ctlTab(intIndex).Caption = .ReadProperty("Caption" & intIndex, "[No Caption]")
            Set ctlTab(intIndex).Image = .ReadProperty("Image" & intIndex, Nothing)
        Next
        
        CurrentTab = .ReadProperty("CurrentTab", 0)
        
    End With
    
    'Force a redraw
    RedrawTabs
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Resize the controls constituant controls

'[Code]

    With UserControl
        linBorderTop.X2 = .Width - 15
        linBorderRight.X1 = .Width - 15
        linBorderRight.X2 = .Width - 15
        linBorderRight.Y2 = .Height - 15
        linBorderBottom.Y1 = .Height - 15
        linBorderBottom.Y2 = .Height - 15
        linBorderBottom.X2 = .Width - 15
        linBorderLeft.Y2 = .Height - 15
    End With
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'[Description]
'   Store the current control properties in the control's PropBag

'[Notes]
'   As there can be multiple caption and image properties (one for each tab),
'   we cycle through the tabs and use the Property Name & Index as a property
'   key

'[Declarations]
Dim intIndex                As Integer      'Used to cycle through the tabs

'[Code]
    
    With PropBag
        .WriteProperty "Tabs", ctlTab.Count, 3
        .WriteProperty "CurrentTab", intCurrentTab, 0
        
        'Cycle through each tab setting its properties
        For intIndex = 0 To ctlTab.Count - 1
            .WriteProperty "Caption" & intIndex, ctlTab(intIndex).Caption, "[No Caption]"
            .WriteProperty "Image" & intIndex, ctlTab(intIndex).Image, Nothing
        Next
        
    End With
    
End Sub

Public Property Get CurrentTab() As Integer

'[Description]
'   Return the current (foremost) tab

'[Code]

    CurrentTab = intCurrentTab
    
End Property

Public Property Let CurrentTab(NewValue As Integer)

'[Description]
'   Set the current tab

'[Code]

    If NewValue < 0 Or NewValue > ctlTab.Count - 1 Then
        'User attempted to set to a tab which does not exist
        Err.Raise vbObjectError + ERR_NOSUCHTAB, "CurrentTab [Let]", "Tab " & NewValue & " does not exist"
    Else
        'Change to the new tab
        ctlTab_Click NewValue
        PropertyChanged "CurrentTab"
    End If
        
End Property

Public Property Get Caption() As String

'[Description]
'   Return the caption of the current tab

'[Code]

    Caption = ctlTab(intCurrentTab).Caption
    
End Property

Public Property Let Caption(NewValue As String)

'[Description]
'   Set the caption of the current tab and force a redraw

'[Code]

    ctlTab(intCurrentTab).Caption = NewValue
    
    RedrawTabs
    
    PropertyChanged "Caption"
    
End Property

Private Function RedrawTabs()

'[Description]
'   Cycles through the tab collection and repositions each tab depending on the
'   width of any previous tabs

'[Declarations]
Dim lngLeftPos              As Long         'Next left position
Dim intIndex                As Integer      'Used to cycle through the tabs

'[Code]

    lngLeftPos = 0 'flush with left hand side of control
    
    For intIndex = 0 To ctlTab.Count - 1
        ctlTab(intIndex).Move lngLeftPos, 0
        lngLeftPos = lngLeftPos + ctlTab(intIndex).Width - 15
    Next

End Function

Public Property Get Tabs() As Integer

'[Description]
'   Return the current number of tabs

'[Code]

    Tabs = ctlTab.Count
    
End Property

Public Property Let Tabs(NewValue As Integer)

'[Description]
'   Set the number of tabs

'[Declarations]
Dim intIndex                As Integer      'Used to cycle through tabs we are
                                            'adding/removing

'[Code]

    If NewValue < 1 Then
        'Must have at lease 1 tab
        Err.Raise vbObjectError + ERR_MINTAB, "Tabs [Property Let]", "You must have at least 1 tab"
        
    ElseIf NewValue < ctlTab.Count Then
        'New tabs value is less than current
        'Remove extra tabs from end
        
        For intIndex = ctlTab.Count - 1 To NewValue Step -1
            Unload ctlTab(intIndex)
        Next
        
        If intCurrentTab > NewValue - 1 Then
            'Current tab has just been removed!
            'set it to the last tab
            CurrentTab = NewValue - 1
        End If
        
    ElseIf NewValue > ctlTab.Count Then
        'New tabs value is greater than current
        'Add extra tabs to end
        
        For intIndex = ctlTab.Count To NewValue - 1
            Load ctlTab(intIndex)
            ctlTab(intIndex).Caption = "[No Caption]"
            ctlTab(intIndex).IsActive = False
            ctlTab(intIndex).Visible = True
        Next
        
        'Force a redraw
        RedrawTabs
    
    End If
    

    
    PropertyChanged "Tabs"
    
End Property

Public Property Get Image() As StdPicture

'[Description]
'   Return the current tab's image

'[Returns]
'   StdPicture or 0 if no image set

'[Code]

    Set Image = ctlTab(intCurrentTab).Image
    
End Property

Public Property Set Image(NewValue As StdPicture)

'[Description]
'   Set the current tabs image

'[Code]

    Set ctlTab(intCurrentTab).Image = NewValue
    
    RedrawTabs
    PropertyChanged "Image"
    
End Property

