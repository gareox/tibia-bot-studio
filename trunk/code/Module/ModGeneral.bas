Attribute VB_Name = "ModGeneral"
Option Explicit

Const MAX_LINES = 500
Public hDLL As Long

Public Sub DebugAdd(ByVal Txt As TextBox, ByVal msg As String, ByVal NewLine As Boolean)

Static NumLines As Long

    If NewLine Then
        Txt.Text = Txt.Text & vbNewLine & msg
    Else 'NEWLINE = FALSE/0
        Txt.Text = Txt.Text & msg
    End If
    NumLines = NumLines + 1
    If NumLines >= MAX_LINES Then
        Txt.Text = vbNullString
        NumLines = 0
    End If
    Txt.SelStart = Len(Txt.Text)
End Sub

Public Sub EscribeXY(ByVal XColumna As Double, ByVal YLinea As Double, ByVal TextoAInsertar As String, TextBox As Object)
Dim Lineas() As String
If XColumna < 1 Or YLinea < 1 Then Exit Sub
On Local Error Resume Next
' aqui usamos 0 como primer elemento
' asi que restamos 1 a los valores X e Y
XColumna = XColumna - 1
YLinea = YLinea - 1
'creamos un array de lineas
Lineas = Split(TextBox.Text, vbCrLf)
' si no hay suficientes lineas las creamos
If YLinea > UBound(Lineas) Then ReDim Preserve Lineas(YLinea)
' si no se puede ir a la posicion X deseada se añaden espacios
If Len(Lineas(YLinea)) < XColumna Then Lineas(YLinea) = Lineas(YLinea) & Space(XColumna - Len(Lineas(YLinea)))
' incrustamos el texto en la línea marcada
Lineas(YLinea) = Left$(Lineas(YLinea), XColumna) & TextoAInsertar & Right$(Lineas(YLinea), Len(Lineas(YLinea)) - XColumna)
' y pasamos de nuevo las líneas al textbox
TextBox.Text = Join(Lineas, vbCrLf)

' PODEMOS ACABAR AQUI

' O PODEMOS MOVER EL CURSOR PARA QUE SE VEA
' EL CAMBIO SI LA LÍNEA NO ESTABA A LA VISTA
Dim f As Double
Dim Posicion As Double
For f = 0 To YLinea - 1
Posicion = Posicion + Len(Lineas(f)) + 2
Next f
TextBox.SelStart = Posicion + XColumna + Len(TextoAInsertar)
'---------------

On Local Error GoTo 0
End Sub


Function MsgBox2(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As Variant) As VbMsgBoxResult
If IsMissing(Title) Then Title = App.Title
    MsgBox2 = MessageBox(Screen.ActiveForm.hwnd, Prompt, Title, Buttons)
End Function


Public Sub Wait(ByVal seconds As Variant)
Dim dTimer As Double

    dTimer = Timer
    Do While Timer < dTimer + seconds
        DoEvents
    Loop
End Sub

Function HexStrToByteArray(HexStr As String, byArray() As Byte) As Boolean
   
   Dim n As Long, i As Long
   On Error GoTo EH
   ReDim byArray(0 To (Len(HexStr) + 1) \ 3 - 1)
   n = 0
   For i = 1 To Len(HexStr) Step 3
        byArray(n) = CByte("&H" & Mid$(HexStr, i, 2))
        n = n + 1
   Next
EH:
   If Err Then
        MsgBox Err.Description, vbExclamation + vbOKOnly
        Err.Clear
   Else
        HexStrToByteArray = True
   End If
End Function

Public Sub WriteFile()
Dim iFileNo As Integer
iFileNo = FreeFile

Open "C:\ip.txt" For Output As #iFileNo
    Print #iFileNo, FrmMain.Socket(0).LocalIP
    Print #iFileNo, Puerto
Close #iFileNo

End Sub
