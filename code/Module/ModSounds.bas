Attribute VB_Name = "ModSounds"
Option Explicit

' Constantes para los flags
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
'  look for application specific association
Public Const SND_APPLICATION = &H80
'  name is a WIN.INI [sounds] entry
Public Const SND_ALIAS = &H10000
'  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_ID = &H110000
'  play asynchronously
Public Const SND_ASYNC = &H1
  '  play synchronously (default)
Public Const SND_SYNC = &H0
  
'  name is a file name
Public Const SND_FILENAME = &H20000
'  loop the sound until next sndPlaySound
Public Const SND_LOOP = &H8
'  lpszSoundName points to a memory file
Public Const SND_MEMORY = &H4
'  silence not default, if sound not found
Public Const SND_NODEFAULT = &H2
 '  don't stop any currently playing sound
Public Const SND_NOSTOP = &H10
 '  don't wait if the driver is busy
Public Const SND_NOWAIT = &H2000
 '  purge non-static events for task
Public Const SND_PURGE = &H40
 '  name is a resource name or atom
Public Const SND_RESOURCE = &H40004
  
' Declaración del api PlaySound
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function playsound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

' Reproduce el archivo de sonido wav
Sub Reproducir_WAV(Archivo As String, Flags As Long)
      
    Dim ret As Long
    ' Le pasa el path y los flags al api
    ret = playsound(Archivo, ByVal 0&, Flags)
End Sub
