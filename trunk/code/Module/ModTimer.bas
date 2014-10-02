Attribute VB_Name = "ModTimer"
Option Explicit
  
'Declaraciones para implemntar el timer con el Api
'***********************************************************
  
  
' Función que crea un timer
Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
  
' Función que detiene el timer iniciado
Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
  
  
' Función Callback que se dispara al iniciar el timer
Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
  Dim i As Integer
If ScriptActivate Then
    For i = 1 To NumScript
        If FrmMain.List1.Selected(i - 1) = True Then
            DoEvents
            Call ScriptExecute(i)
        End If
    
    Next
End If
End Sub
  
'Inicia
Sub Iniciar_Timer(Hwnd_Form As Long, Intervalo As Long, ID As Long)
    SetTimer Hwnd_Form, ID, Intervalo, AddressOf TimerProc
End Sub
  
'Detiene
Sub Detener_Timer(Hwnd_Form As Long, ID As Long)
    KillTimer Hwnd_Form, ID
End Sub

