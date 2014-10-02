Attribute VB_Name = "modCDialog"
Option Explicit

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_EXPLORER = &H80000

Public Type OPENFILENAME
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 flags As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim ofn As OPENFILENAME

'Muestra el cuadro de dialogo para abrir archivos:
Public Function OpenFile(hwnd As Long, filter As String, Title As String, InitDir As String, Optional FileName As String, Optional FilterIndex As Long) As String
 On Local Error Resume Next

 Dim ofn As OPENFILENAME
 Dim a As Long
 
 ofn.lStructSize = Len(ofn)
 ofn.hwndOwner = hwnd
 ofn.hInstance = App.hInstance
 
 If VBA.Right$(filter, 1) <> "|" Then filter = filter + "|"
 
 For a = 1 To Len(filter)
 If Mid$(filter, a, 1) = "|" Then Mid(filter, a, 1) = Chr(0)
 Next
 
 ofn.lpstrFilter = filter
 ofn.lpstrFile = Space$(254)
 ofn.nMaxFile = 255
 ofn.lpstrFileTitle = Space$(254)
 ofn.nMaxFileTitle = 255
 ofn.lpstrInitialDir = InitDir
 If Not FileName = vbNullString Then ofn.lpstrFile = FileName & Space$(254 - Len(FileName))
 ofn.nFilterIndex = FilterIndex
 ofn.lpstrTitle = Title
 ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
 a = GetOpenFileName(ofn)

 If a Then
   OpenFile = Trim$(ofn.lpstrFile)
   If VBA.Right$(VBA.Trim$(OpenFile), 1) = Chr(0) Then OpenFile = VBA.Left$(VBA.Trim$(ofn.lpstrFile), Len(VBA.Trim$(ofn.lpstrFile)) - 1)
   
 Else
   OpenFile = vbNullString
   
 End If
 
End Function

'Muestra el cuadro de dialogo para guardar archivos:
Public Function SaveFile(hwnd As Long, filter As String, Title As String, InitDir As String, Optional FileName As String, Optional FilterIndex As Long) As String
 On Local Error Resume Next
 Dim ofn As OPENFILENAME
 Dim a As Long
Dim Ext As String


 ofn.lStructSize = Len(ofn)
 ofn.hwndOwner = hwnd
 ofn.hInstance = App.hInstance
 
 If VBA.Right$(filter, 1) <> "|" Then filter = filter + "|"
 
 For a = 1 To Len(filter)
 If Mid(filter, a, 1) = "|" Then Mid(filter, a, 1) = Chr(0)
 Next
 
 ofn.lpstrFilter = filter
 ofn.lpstrFile = Space(254)
 ofn.nMaxFile = 255
 ofn.lpstrFileTitle = Space(254)
 ofn.nMaxFileTitle = 255
 ofn.lpstrInitialDir = InitDir
 If Not FileName = vbNullString Then ofn.lpstrFile = FileName & Space(254 - Len(FileName))
 ofn.nFilterIndex = FilterIndex
 ofn.lpstrTitle = Title
 ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT Or OFN_EXPLORER
 a = GetSaveFileName(ofn)


       Ext = GetExtension(ofn.lpstrFilter, ofn.nFilterIndex)
 If a Then
   SaveFile = Trim$(ofn.lpstrFile)
   'If VBA.Right$(Trim$(SaveFile), 1) = Chr(0) Then SaveFile = VBA.Left$(Trim$(ofn.lpstrFile), Len(Trim$(ofn.lpstrFile)) - 1)
    If VBA.Right$(Trim$(SaveFile), 1) = Chr(0) Then
    SaveFile = Left$(ofn.lpstrFile, InStr(1, ofn.lpstrFile, Chr(0)) - 1) & Ext
    End If

       'Comprobamos si el nombre ya contiene la extension, si no la tiene se la añadimos:
       
       'If Not UCase(Right(DLG_SaveFile, 4)) = UCase(Ext) Then DLG_SaveFile = DLG_SaveFile + Ext

 Else
   SaveFile = vbNullString
   
 End If
 
End Function

Private Function GetExtension(sfilter As String, pos As Long) As String
 Dim Ext() As String
 
 Ext = Split(sfilter, vbNullChar)
 
 If pos = 1 And Ext(pos) <> "*.*" Then
 GetExtension = "." & Replace(Ext(pos), "*.", "")
  
 ElseIf pos = 1 And Ext(pos) = "*.*" Or InStr(Ext(pos + 1), "*.*") Then
 GetExtension = vbNullString
 
 Else
 GetExtension = "." & Replace(Ext(pos + 1), "*.", "")
 
 End If
 
End Function


