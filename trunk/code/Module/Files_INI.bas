Attribute VB_Name = "Files_INI"
Option Explicit



Public Path_Archivo_Ini As String

'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long


'Lee un dato _
-----------------------------
'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Public Function Leer_Ini(Path_INI As String, APPLICATION As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Public Function Grabar_Ini(Path_INI As String, APPLICATION As String, Key As String, Valor As Variant) As String

    WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI

End Function







