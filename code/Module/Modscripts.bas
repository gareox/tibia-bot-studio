Attribute VB_Name = "Modscripts"
Option Explicit

Public Script(999) As MSScriptControl.ScriptControl
Public MainScript As New ClsScriptFunctions
Public Active(999) As Boolean
Public NumScript As Integer
Public loadedFile As String
Public fIndex As Integer

Public Const ScriptExtencion = "*.tbs"
Public Const ScriptDescription = "Tibia Bot Script"

Public Sub Reset_ObjScript(ByVal index As Integer)
    Set Script(index) = Nothing
    Load_ObjSCript index
End Sub

Public Sub Load_ObjSCript(ByVal index As Integer)
 ' Inicializa la variable para usar el ScriptControl
    Set Script(index) = New MSScriptControl.ScriptControl
      
    Script(index).Language = "VBScript"
    Script(index).AddObject "my", MainScript
    
    Script_Propiedades index

End Sub

Sub Script_Propiedades(ByVal index As Long)
    Script(index).Language = "VBScript"
    Script(index).AllowUI = True
End Sub


Public Sub Load_Scripts(ByVal index As Integer, ByVal File As String)
   Dim ScriptCode As String
   Dim intFile As Integer

    intFile = FreeFile
   
    'Cargamos el archivo del script.
    Open App.Path & "\Scripts\" & File & ".tbs" For Input As #intFile
        ScriptCode = Input$(LOF(intFile), #intFile) '  LOF returns Length of File
    Close #intFile
    
    'Ejecutamos el Script.
    'Script(Index).RunCode ScriptCode
    On Error GoTo ErrSub:
    Script_Propiedades index
    Script(index).AddCode ScriptCode
    Exit Sub
ErrSub:
    MsgBox "Error: " & Script(index).Error.Description & vbNewLine & _
        "Linea: " & Script(index).Error.Line & vbNewLine & _
        "Columna: " & Script(index).Error.Column & vbNewLine
    'End
End Sub

Public Function ScriptExecute(ByVal index As Integer)
On Error GoTo Error:
   Call Script(index).ExecuteStatement("main")
   Exit Function
Error:
MsgBox "Error: " & Script(index).Error.Description & vbNewLine & _
        "Linea: " & Script(index).Error.Line & vbNewLine & _
        "Columna: " & Script(index).Error.Column & vbNewLine
        
            FrmMain.T1(index).Enabled = False
            FrmMain.Mnu_Start(index).Enabled = True
            FrmMain.Mnu_Stop(index).Enabled = False
            Active(index) = False
            FrmMain.LScripts.SelectedItem.Checked = False
End Function

Public Sub LoadFileSCripts()
Dim sArchivo As String
Dim Count As Integer
Dim Autor As String
Dim Desc As String


    sArchivo = Dir(App.Path & "\Scripts\" & ScriptExtencion)
    Count = 0
    'FrmMain.AddScript.clearUsername
    
Do While sArchivo <> ""
    
    Count = Count + 1
    
    Load FrmMain.T1(Count)
    Load FrmMain.Mnu_Start(Count)
    Load FrmMain.Mnu_Stop(Count)
    
    FrmMain.Mnu_Start(Count).Visible = False
    FrmMain.Mnu_Stop(Count).Visible = False
    
    Call Load_ObjSCript(Count)
    Call Load_Scripts(Count, Mid(sArchivo, 1, Len(sArchivo) - 4))
    
    Autor = Script(Count).Eval("Autor")
    Desc = Script(Count).Eval("Description")
    If Len(Autor) = 0 Then Autor = "Desconocido"
    If Len(Desc) = 0 Then Desc = "Sin Descripcion"

  FrmMain.LScripts.ListItems.Add , , Mid(sArchivo, 1, Len(sArchivo) - 4)
  FrmMain.LScripts.ListItems.Item(Count).SubItems(1) = Autor
  FrmMain.LScripts.ListItems.Item(Count).SubItems(2) = Desc

       
    sArchivo = Dir
Loop
  NumScript = Count

End Sub
