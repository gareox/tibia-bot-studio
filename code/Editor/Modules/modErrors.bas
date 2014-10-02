Attribute VB_Name = "modErrors"
Option Explicit

Public pError As Boolean
Public Errors As String

Sub InitErrors()
    Errors = ""
End Sub

Sub ErrMessage(Text As String)
    Errors = Errors & Text & " [Line:" & GetLineNumber(SourcePos) & "]" & vbCrLf
    SkipBlank
    frmMain.Code.ErrorSelectLineBySourcePos SourcePos
    pError = True
End Sub

Sub InfMessage(Text As String)
    MsgBox Text, vbInformation, "Libry Compiler"
    pError = False
End Sub

Function GetLineNumber(CurrentPosition As Long)
    Dim ActualLine As Integer
    Dim i As Long
    
    ActualLine = 1
    For i = 1 To CurrentPosition
        If Mid(frmMain.Code.Text, i, 2) = vbCrLf Then
            ActualLine = ActualLine + 1
        End If
    Next i
    GetLineNumber = ActualLine
End Function


