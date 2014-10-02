Attribute VB_Name = "Modwindow"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE As Long = (-16&)
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1


Public Sub ChangeFrom(win As Form, Index As Integer)
 Dim style As Long
 
With win
    Select Case Index
        Case 0:
        
            style = GetWindowLong(.hwnd, GWL_STYLE)
            style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
            style = SetWindowLong(.hwnd, GWL_STYLE, style)
  
        Case 1:
            style = GetWindowLong(.hwnd, GWL_STYLE)
            style = style And Not WS_THICKFRAME
            style = style And Not WS_MAXIMIZEBOX
            style = SetWindowLong(.hwnd, GWL_STYLE, style)
            FrmMain.LoadSizeMain
    End Select
End With
End Sub

