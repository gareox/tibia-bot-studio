Attribute VB_Name = "ModScriptControl"
  Option Explicit
  
Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
Const SC_MARK_BACKGROUND = 22

Public FunctionPrototypes() As String

Function ReadFile(FileName)
Dim f, temp
  f = FreeFile
  temp = ""
   Open FileName For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(FileName), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

  Public Sub Load_Scripcontrol()
    
      
  End Sub
  
Sub push(ary, Value) 'this modifies parent ary object
Dim x
    On Error GoTo Init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init: ReDim ary(0): ary(0) = Value
End Sub

Sub LoadFile(fpath As String)
   
   loadedFile = fpath
   FrmEditor.scivb.DeleteAllMarkers
   FrmEditor.scivb.LoadFile loadedFile
 
End Sub

Function LoadFunctionPrototypes(fpath As String)
Dim a
  Erase FunctionPrototypes
  If Not FileExists(fpath) Then Exit Function
  
  Dim tmp() As String, x
  Const endMarker = "#modules"
  
  tmp = Split(ReadFile(fpath), vbCrLf)
  For Each x In tmp
        x = Trim(x)
        If Left(x, Len(endMarker)) = endMarker Then Exit For
        If Len(x) > 0 And Left(x, 1) <> "#" And Left(x, 1) <> "'" Then
            a = InStr(x, "(")
            If a > 2 Then x = Mid(x, 1, a - 1)
            push FunctionPrototypes, x
        End If
        
  Next
  
  LoadFunctionPrototypes = UBound(FunctionPrototypes)
  
End Function
