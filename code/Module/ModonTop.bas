Attribute VB_Name = "ModonTop"
Option Explicit

      Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2

      Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

      Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
            SetTopMostWindow = False
         End If
      End Function

