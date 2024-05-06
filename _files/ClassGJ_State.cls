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

Event Ready()

Private Type DeferredCall
    DoAction As String
    MethodIndex As Long
    Args() As Variant
End Type

Dim DeferredCalls() As DeferredCall
Dim DeferredCallsMax As Long

Public Sub RaiseReady()
If webviewWindow Is Nothing Then Exit Sub
RaiseEvent Ready
If DeferredCallsMax = -1 Then Exit Sub
Dim Z As Long
For Z = 0 To DeferredCallsMax
    VTableCallEx DeferredCalls(Z).DoAction, webviewWindow, DeferredCalls(Z).MethodIndex, DeferredCalls(Z).Args
Next
DeferredCallsMax = -1
Erase DeferredCalls
End Sub

Property Get Url() As String
Url = PrivateUrl
End Property

Function OpenURL(ByVal Url As String)
    If Left(LCase(Url), 4) <> "http" And Left(LCase(Url), 4) <> "file" Then Url = "http://" & Url
    PrivateUrl = Url
    Dim Args() As Variant
    Args = Array(StrPtr(Url))
    VTableCallEx "Navigate", webviewWindow, 3, Args
End Function

Sub WebReSize()
    Dim RECT1 As RECT
    GetClientRect webHostHwnd, RECT1
    'webviewController->put_Bounds(bounds);
    Dim Args() As Variant
    Args = Array(RECT1.Left, RECT1.Top, RECT1.Right, RECT1.Bottom)
    VTableCallEx "put_Bounds", webviewController, 4, Args
End Sub

Sub DeferableWindowVTableCallEx(DoAction As String, ByVal MethodIndex As Long, ParamArray Args() As Variant)
If webviewWindow Is Nothing Then
    DeferredCallsMax = DeferredCallsMax + 1
    ReDim Preserve DeferredCalls(DeferredCallsMax)
    DeferredCalls(DeferredCallsMax).DoAction = DoAction
    DeferredCalls(DeferredCallsMax).MethodIndex = MethodIndex
    DeferredCalls(DeferredCallsMax).Args = Args
Else
    Dim Args2() As Variant
    Args2 = Args
    VTableCallEx DoAction, webviewWindow, MethodIndex, Args2
End If
End Sub

Private Function VTableCallEx(DoAction As String, pUnk, ByVal MethodIndex As Long, Args() As Variant) As Variant
    If pUnk Is Nothing Then Exit Function
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = Args
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), (MethodIndex + 2) * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), VTableCallEx)
    DoErr hResult, DoAction
End Function
 
Private Sub DoErr(ByVal ErrCode As Long, Optional DoAction As String)
    Const HideErrInfo As Boolean = True
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

Private Sub Class_Initialize()
DeferredCallsMax = -1
End Sub
