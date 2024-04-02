VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGJ_State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_InsertS = &H200

Private Declare Function DispCallFunc Lib "oleaut32" _
   (ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal CallConv As Long, _
    ByVal vtReturn As VbVarType, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As Long, _
    ByRef pvargResult As Variant) As Long
    
Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private PrivateUrl As String

Public webviewWindow As Object
Public webviewController As Object
Public ICoreWebView2EnvironmentObjB As Object
Public ICoreWebView2EnvironmentObjA As Object

'CallWebView2
Public WebViewShowOK As Boolean
Public webHostHwnd As Long

Private hResult As Long
Private v As Variant

Property Get Url() As String
Url = PrivateUrl
End Property

Function OpenURL(ByVal Url As String)
    If Left(LCase(Url), 4) <> "http" Then Url = "http://" & Url
    PrivateUrl = Url
    v = VTableCallEx("Navigate", webviewWindow, 3, StrPtr(Url))
End Function

Function OpenDevToolsWindow()
    v = VTableCallEx("OpenDevToolsWindow", webviewWindow, 49)
End Function

Function ExecuteScript(JS As String)
    v = VTableCallEx("ExecuteScript", webviewWindow, 27, StrPtr(JS), 0&)
End Function

Function AddHostObjectToScript(ObjName As String, Obj1 As Object)
    v = VTableCallEx("AddHostObjectToScript", webviewWindow, 47, StrPtr(ObjName), ObjPtr(Obj1))

    AddHostObjectToScript = hResult = 0
End Function

Function WebReSize() As Boolean
    Dim RECT1 As RECT
    GetClientRect webHostHwnd, RECT1
    'webviewController->put_Bounds(bounds);
    v = VTableCallEx("put_Bounds", webviewController, 4, RECT1.Left, RECT1.Top, RECT1.Right, RECT1.Bottom)
    WebReSize = hResult = 0
End Function

Function VTableCallEx(DoAction As String, pUnk, ByVal MethodIndex As Long, ParamArray A() As Variant) As Variant
    If pUnk Is Nothing Then Exit Function
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), (MethodIndex + 2) * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), VTableCallEx)
    DoErr hResult, DoAction
End Function
 
Sub DoErr(ByVal ErrCode As Long, Optional DoAction As String)
    Const HideErrInfo  As Boolean = True
    If ErrCode = 0 Then Exit Sub
    Dim Info As String
    If ErrCode = -1 Then
        ErrCode = ERR.Number
        Info = ERR.Description
    Else
        Info = GetLastDllErr(hResult)
    End If
    ERR.Raise ErrCode, "DispCallFunc"
    Info = "Err:" & DoAction & ",Errid=" & ErrCode & ",Info:" & Info
    If Not HideErrInfo Then Debug.Print Info
End Sub

Private Function GetLastDllErr(ByVal lErr As Long) As String
Dim sReturn As String
sReturn = String$(256, 32)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_InsertS, 0&, lErr, 0&, sReturn, Len(sReturn), ByVal 0
sReturn = Trim(sReturn)
GetLastDllErr = sReturn
End Function

