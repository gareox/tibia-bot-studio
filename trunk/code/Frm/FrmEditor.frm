VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEditor 
   Caption         =   "Editor"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   Icon            =   "FrmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9165
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicTop 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "FrmEditor.frx":164A
      ScaleHeight     =   735
      ScaleWidth      =   45000
      TabIndex        =   6
      Top             =   0
      Width           =   45000
   End
   Begin VB.CommandButton CmdRun 
      Caption         =   "Run Script"
      Height          =   300
      Left            =   7800
      TabIndex        =   3
      Top             =   5550
      Width           =   1215
   End
   Begin TibiaStudio.CodeEdit Code 
      Height          =   4935
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      LineNumbers     =   -1  'True
      BoldSelectedKeyWords=   -1  'True
   End
   Begin MSComctlLib.TreeView PicTree 
      Height          =   4935
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8705
      _Version        =   393217
      LabelEdit       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "Nuevo Script"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   5550
      Width           =   1215
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Abrir Script"
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   5550
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Guardar Script"
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   5550
      Width           =   1215
   End
   Begin TibiaStudio.AeroStatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   5910
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      SimpleText      =   "Editor"
   End
End
Attribute VB_Name = "FrmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tScript As MSScriptControl.ScriptControl

Public isDirty As Boolean
Public ScriptFile As String

Public Sub Load_tSCript()
 ' Inicializa la variable para usar el ScriptControl
    Set tScript = New MSScriptControl.ScriptControl
      
    tScript.Language = "VBScript"
    tScript.AddObject "my", MainScript
    tScript.AllowUI = True
End Sub

Public Sub Reset_tScript()
    Set tScript = Nothing
    Load_tSCript
End Sub

Public Sub Load_tScripts(ByVal File As String)
   Dim ScriptCode As String
   Dim intFile As Integer

    intFile = FreeFile
    Me.Enabled = False
    Screen.MousePointer = 11
    'Cargamos el archivo del script.
    Open File For Input As #intFile
        ScriptCode = Input$(LOF(intFile), #intFile) '  LOF returns Length of File
    Close #intFile
    

    On Error GoTo ErrSub:

    tScript.AddCode ScriptCode
    Call tScript.ExecuteStatement("main")

    MsgBox "Script Se Ejecuto Correctamente!", vbInformation + vbSystemModal
    Me.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrSub:
    MsgBox "Error: " & tScript.Error.Description & vbNewLine & _
        "Linea: " & tScript.Error.Line & vbNewLine & _
        "Columna: " & tScript.Error.Column & vbNewLine
        Me.Enabled = True
         Screen.MousePointer = 0
    'End
End Sub

Private Sub CmdRun_Click()
    Reset_tScript
    Load_tScripts ScriptFile
End Sub

Private Sub Code_Change()
    isDirty = True
End Sub


Private Sub Form_Load()
    DoEvents
    LoadCode
    Create_Template
    Load_tSCript
End Sub

Public Sub LoadCode()
    Code.ceKeyWords = "*my*sub*dim*end*alias*for*while*const*byte*if*else*type*as*redim*preserve*ubound*next*exit*byval*not*Option*Explicit*"
    Code.ceOperators = "**"
    Code.ColourEntireRTB

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If isDirty Then
        Select Case MsgBox("El Script Tiene Cambias, Deseas Guardarlos?", vbYesNoCancel)
            Case vbYes: Call CmdSave_Click
            Case vbCancel: Cancel = 1
        End Select
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next 'PicTop
    
    PicTree.Top = PicTop.Height
    PicTree.Height = Me.Height - StatusBar2.Height - PicTop.Height - 1000
    
    Code.Top = PicTop.Height
    Code.Height = PicTree.Height
    Code.Width = Me.Width - PicTree.Width - 200
    
    CmdNew.Top = PicTree.Height + PicTop.Height + 50
    CmdOpen.Top = CmdNew.Top
    CmdSave.Top = CmdOpen.Top
    
    CmdRun.Top = CmdSave.Top
    CmdRun.Left = Me.Width - CmdRun.Width - 350
End Sub


Private Sub CmdNew_Click()

Dim tmpNew As String
    tmpNew = tmpNew & "Option Explicit" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "'Template [Tibia Bot Script]" & vbNewLine
    tmpNew = tmpNew & "Const ScriptName = " & """TibiaStudio""" & vbNewLine
    tmpNew = tmpNew & "Const Autor = " & """TibiaStudio""" & vbNewLine
    tmpNew = tmpNew & "Const Description = " & """Template TibiaStudio""" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "sub config()" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "end sub" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "sub main()" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    tmpNew = tmpNew & "end sub" & vbNewLine
    tmpNew = tmpNew & "" & vbNewLine
    
    If isDirty Then
            If MsgBox("Deseas guardar los cambiaos antes de continuar?", vbInformation + vbYesNo) = vbYes Then
                If Len(ScriptFile) = 0 Then
                    ScriptFile = SaveFile(Me.hwnd, ScriptDescription & " |" & ScriptExtencion, "Abrir SCript", App.Path & "\Scripts\", vbNullString)
                    If Len(ScriptFile) = 0 Then Exit Sub
                End If
                Code.SaveFile ScriptFile
            End If
    End If
    
    ScriptFile = SaveFile(Me.hwnd, ScriptDescription & "|" & ScriptExtencion, "Abrir SCript", App.Path & "\Scripts\", vbNullString)
    If Len(ScriptFile) = 0 Then Exit Sub
    Code.Text = Empty
    Code.Text = tmpNew
    Code.SaveFile ScriptFile
    Code.LoadFile ScriptFile
    isDirty = False
    
End Sub

Private Sub CmdOpen_Click()
    ScriptFile = OpenFile(Me.hwnd, ScriptDescription & " |" & ScriptExtencion, "Abrir SCript", App.Path & "\Scripts\", vbNullString)
    
    If Len(ScriptFile) = 0 Then Exit Sub
    Code.LoadFile ScriptFile
    isDirty = False
End Sub

Private Sub CmdSave_Click()

    If ScriptFile <> "" Then
        Code.SaveFile ScriptFile
        isDirty = False
        MsgBox "Script Guardado Correctamente.", vbInformation + vbSystemModal
    Else
    
        ScriptFile = SaveFile(Me.hwnd, ScriptDescription & " |" & ScriptExtencion, "Abrir SCript", App.Path & "\Scripts\", vbNullString)
        Code.SaveFile ScriptFile
        isDirty = False
        MsgBox "Script Guardado Correctamente.", vbInformation + vbSystemModal
    End If
    
        

End Sub

Private Sub Create_Template()
    PicTree.Nodes.Add , , "root", "Template"

    PicTree.Nodes.Add "root", tvwChild, "troot", "tibiasock"
    PicTree.Nodes.Add "troot", tvwChild, "t1", "my.SendToServer"
    PicTree.Nodes.Add "troot", tvwChild, "t2", "my.SendToClient"
    
    PicTree.Nodes.Add "root", tvwChild, "proot", "Player"
    PicTree.Nodes.Add "proot", tvwChild, "p1", "my.hp"
    PicTree.Nodes.Add "proot", tvwChild, "p2", "my.mp"
    PicTree.Nodes.Add "proot", tvwChild, "p3", "my.online"
    PicTree.Nodes.Add "proot", tvwChild, "p4", "my.atack"
    PicTree.Nodes.Add "proot", tvwChild, "p5", "my.isatack"
    PicTree.Nodes.Add "proot", tvwChild, "p6", "my.walk"
    PicTree.Nodes.Add "proot", tvwChild, "p7", "my.iswalk"
    PicTree.Nodes.Add "proot", tvwChild, "p8", "my.say"
    PicTree.Nodes.Add "proot", tvwChild, "p9", "my.posx"
    PicTree.Nodes.Add "proot", tvwChild, "p10", "my.poxy"
    PicTree.Nodes.Add "proot", tvwChild, "p11", "my.poxz"
    
    PicTree.Nodes.Add "root", tvwChild, "wroot", "Window"
    PicTree.Nodes.Add "wroot", tvwChild, "w1", "my.window"
    PicTree.Nodes.Add "wroot", tvwChild, "w2", "my.button"
    PicTree.Nodes.Add "wroot", tvwChild, "w3", "my.TextBox"
    PicTree.Nodes.Add "wroot", tvwChild, "w4", "my.text"
    PicTree.Nodes.Add "wroot", tvwChild, "w5", "my.label"
    
    PicTree.Nodes.Add "root", tvwChild, "sroot", "System"
    PicTree.Nodes.Add "sroot", tvwChild, "s1", "my.readconfig"
    PicTree.Nodes.Add "sroot", tvwChild, "s2", "my.saveconfig"
    PicTree.Nodes.Add "sroot", tvwChild, "s3", "my.wait"
    PicTree.Nodes.Add "sroot", tvwChild, "s4", "my.DoEvent"
    PicTree.Nodes.Add "sroot", tvwChild, "s5", "my.playsound"
    PicTree.Nodes.Add "sroot", tvwChild, "s6", "my.path"
End Sub


Private Sub PicTree_NodeClick(ByVal Node As MSComctlLib.Node)

If LCase(Node.Text) = LCase("template") Or LCase(Node.Text) = LCase("player") Or LCase(Node.Text) = LCase("window") Or LCase(Node.Text) = LCase("system") Or LCase(Node.Text) = LCase("tibiasock") Then Exit Sub

   EscribeXY Code.CurrentColumnNumber + 1, Code.CurrentLineNumber, Node.Text, Code

End Sub
