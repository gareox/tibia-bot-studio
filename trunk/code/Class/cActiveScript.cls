VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActiveScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Implements IActiveScriptSite
Implements IActiveScriptSiteWindow 'this allows script to have a ui (like messageboxes)
'Implements IActiveScriptSiteDebug

Dim IScript As IActiveScript
Dim IParse As IActiveScriptParse

Private cObjects As New Collection
Private AppCookie As Long

Const TEXT_DOC_ATTR_READONLY As Long = 1
Const TYPE_E_ELEMENTNOTFOUND = 134227970

Event Error(description As String, ScriptSource As String, lineNumber As Long, charposition As Long)

Dim debugSourceContext As Long


Private Sub Class_Initialize()

    Dim hRes As Long
    Dim oUnk As vbActiveScript3.IUnknown
    Dim clsidVBS As vbActiveScript3.UUID
    Dim uuidActScr As vbActiveScript3.UUID
    
    'Create instance of the language engine
    CLSIDFromProgID "VBScript", clsidVBS
    CLSIDFromString IID_IActiveScript, uuidActScr
    
    'Start inproc script engine, VBSCRIPT.DLL
    hRes = CoCreateInstance(clsidVBS, Nothing, CLSCTX_INPROC_SERVER, uuidActScr, oUnk)
    
    'Get the IActiveScript interface
    Set IScript = oUnk
    
    IScript.SetScriptSite Me
        
    'get the iparse interface
    Set IParse = IScript
    
    IParse.InitNew
    
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    cleanup cObjects
    IScript.Close
    Set IParse = Nothing
    Set IScript = Nothing
End Sub

Private Sub cleanup(c As Collection)
    Dim o As Object
    For Each o In c
        Set o = Nothing
    Next
    Set c = Nothing
End Sub

Sub Reset()
    
    cleanup cObjects
    Set cObjects = New Collection
    
    IScript.Close
    IScript.SetScriptSite Me

End Sub

Sub RunCode(scode As String)
    
    Dim exep As vbActiveScript3.EXCEPINFO

 On Error GoTo Error:

    
    IParse.ParseScriptText scode, _
                           Empty, _
                           Nothing, _
                           Empty, _
                           0, _
                           0, _
                           0, _
                           Null, _
                           exep
                           
 Exit Sub
 
Error:
DebugAdd FrmMain.TxtDebug, "Error en RunCode: " & Err.description, True

End Sub

Function Eval(sExpression As String) As Variant

    On Error Resume Next
    
    Dim exep As vbActiveScript3.EXCEPINFO
    Dim vAnswer As Variant

    IParse.ParseScriptText sExpression, _
                            Empty, _
                            Nothing, _
                            "", _
                            0, _
                            0, _
                            SCRIPTTEXT_ISEXPRESSION Or SCRIPTTEXT_ISVISIBLE, _
                            vAnswer, _
                            exep
                            
    Eval = vAnswer

End Function

Function CallFunction(funcName As String, ParamArray args())
 
    Dim idisp As vbActiveScript3.IDispatch
    Dim exp As vbActiveScript3.EXCEPINFO
    Dim numArgs As Integer

    'On Error GoTo Error

    
    IScript.GetScriptDispatch "", idisp

    numArgs = UBound(args)
    
    If numArgs = -1 Then
        CallFunction = CallByName(idisp, funcName, VbMethod)
    Else
        Select Case numArgs
            Case 0: CallFunction = CallByName(idisp, funcName, VbMethod, args(0))
            Case 1: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1))
            Case 2: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2))
            Case 3: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2), args(3))
            Case 4: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2), args(3), args(4))
            Case 5: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5))
            Case 6: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
            Case 7: CallFunction = CallByName(idisp, funcName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
            Case Else:
                    MsgBox "CallFunction only supports a max of 7 Args to call function.", vbInformation
        End Select
    End If

   Exit Function
   
Error:
   'MsgBox "Error en Ejecutar Script: " & """" & funcName & """" & vbCrLf & Err.description, vbCritical + vbSystemModal
  Call DebugAdd(FrmMain.TxtDebug, "Error en CallFunction: " & """" & funcName & """" & vbCrLf & Err.description, True)
End Function


Sub SetScriptState(scrState As SCRIPTSTATE)
    IScript.SetScriptState scrState
End Sub

Sub AddObject(sName As String, obj As Object)
    
    If obj Is Nothing Then
        Err.Raise vbObjectError, "AddObject", "Object can not be nothing"
    End If
    
    cObjects.Add obj, sName
    IScript.AddNamedItem StrPtr(sName), SCRIPTITEM_ISVISIBLE Or SCRIPTITEM_GLOBALMEMBERS

End Sub


'*********************************************************************
' IActiveScript Implementation
'
Private Sub IActiveScriptSite_GetItemInfo(ByVal pstrName As String, ByVal dwReturnMask As vbActiveScript3.SCRIPTINFO, ppiunkItem As Long, ppti As Long)
     Dim idisp As vbActiveScript3.IDispatch
     Dim obj As Object
     
     On Error Resume Next
     
     
     Set obj = cObjects(pstrName)
     If Err.Number = 0 Or ObjPtr(obj) = 0 Then
          If dwReturnMask = SCRIPTINFO_IUNKNOWN Then
                Set idisp = obj
                Debug.Print "Asking for an Iunknown ** Using Ptr to Idispatch: " & ObjPtr(idisp)
                ppiunkItem = ObjPtr(idisp)
          Else
              Debug.Print "Looking for * to ITypeInfo, not supported"
              Err.Raise TYPE_E_ELEMENTNOTFOUND
          End If
          Set obj = Nothing
     Else
        Debug.Print "Cant find object named: " & pstrName
        Err.Raise TYPE_E_ELEMENTNOTFOUND
     End If
  
     
End Sub

Private Sub IActiveScriptSite_OnScriptError(ByVal pscripterror As vbActiveScript3.IActiveScriptError)
    
    Dim exp As vbActiveScript3.EXCEPINFO
    Dim tmp As String
    Dim context As Long, lineNo As Long, charPos As Long
    
    On Error Resume Next
    
    With pscripterror
        .GetExceptionInfo exp
        .GetSourceLineText tmp
        .GetSourcePosition context, lineNo, charPos
    End With
    

    'RaiseEvent Error(exp.description, tmp, lineNo, charPos)
 'MsgBox "Error: " & exp.description & tmp & vbNewLine & "Linea: " & lineNo & vbNewLine & "Pos: " & charPos, vbCritical + vbSystemModal, "Script"
Call DebugAdd(FrmMain.TxtDebug, "Error: " & exp.description & tmp & " | " & "Linea: " & lineNo & " | " & "Pos: " & charPos, True)
End Sub

Private Function IActiveScriptSite_GetLCID() As Long
'On Error Resume Next
    'Err.Raise E_NOTIMPL
End Function

Private Function IActiveScriptSite_GetDocVersionString() As String

End Function

Private Sub IActiveScriptSite_OnEnterScript()

End Sub

Private Sub IActiveScriptSite_OnLeaveScript()

End Sub

Private Sub IActiveScriptSiteWindow_EnableModeless(ByVal fEnable As vbActiveScript3.Bool)
    
End Sub

Private Function IActiveScriptSiteWindow_GetWindow() As Long
    'without this our script could not show any kind of UI like messagboxes
    IActiveScriptSiteWindow_GetWindow = 0
End Function

Private Sub IActiveScriptSite_OnScriptTerminate(pvarResult As Variant, pexcepinfo As vbActiveScript3.EXCEPINFO)

End Sub

Private Sub IActiveScriptSite_OnStateChange(ByVal ssScriptState As vbActiveScript3.SCRIPTSTATE)

End Sub

 

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


