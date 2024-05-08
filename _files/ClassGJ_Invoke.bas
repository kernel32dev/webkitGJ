Attribute VB_Name = "ClassGJ_Invoke"
Option Explicit
'A: For Call back:ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
'B: AAICoreWebView2CreateCoreWebView2ControllerCompletedHandler

Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function DispCallFunc Lib "oleaut32" _
   (ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal CallConv As Long, _
    ByVal vtReturn As VbVarType, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As Long, _
    ByRef pvargResult As Variant) As Long
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal FillByte As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal lpVoid As Long)


Public Type ClassGJ_Layout
    pVTable As Long
    RefC As Long
    State As ClassGJ_State
End Type
 
Private m_VTable_Initialized As Boolean
Private m_VTableA(3) As Long
Private m_VTableB(3) As Long
Private m_VTableC(3) As Long


Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)



Private Declare Function NewWebView Lib "WebView2Loader.dll" _
Alias "CreateCoreWebView2EnvironmentWithOptions" (browserExecutableFolder As String, userDataFolder As String _
, ByVal environmentOptions As Long, ByVal createdEnvironmentCallback As Long) As Long
 
Private Const E_FAIL As Long = &H80004005
Private Const S_OK As Long = &H0

Private Const E_NOINTERFACE As Long = &H80004002
Private Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long
Private Declare Function CoInitialize Lib "ole32" ( _
                        ByRef pvReserved As Any) As Long




Private Type POINTAPI
        X As Long
        Y As Long
End Type
Public IsVb6 As Boolean
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Function SetDebug() As Boolean
    SetDebug = True
    IsVb6 = True
End Function

Private Function ObjFromPtr(Optr As Long) As Object
CopyMemory ByVal VarPtr(ObjFromPtr), ByVal Optr, 4
 
End Function

Private Function DispCallByVtbl(pUnk, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long

    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    'MsgBox "DispCallFunc hResult=" & hResult
    'hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), 0)
    If hResult < 0 Then
        ERR.Raise hResult, "DispCallFunc"
    End If
End Function

Private Function GetAddress(ByVal pfn As Long) As Long
    GetAddress = pfn
End Function


'Invoke=ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
'CreateCoreWebView2EnvironmentWithOptions(nullptr, nullptr, nullptr,
 '       Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>

Function PrivateNewClassGJ(State As ClassGJ_State, ByVal VTableSelect As Long) As IUnknown

   If Not m_VTable_Initialized Then
        m_VTableA(0) = GetAddress(AddressOf QueryInterface)
        m_VTableA(1) = GetAddress(AddressOf AddRef)
        m_VTableA(2) = GetAddress(AddressOf Release)
        m_VTableA(3) = GetAddress(AddressOf InvokeA)
        m_VTableB(0) = GetAddress(AddressOf QueryInterface)
        m_VTableB(1) = GetAddress(AddressOf AddRef)
        m_VTableB(2) = GetAddress(AddressOf Release)
        m_VTableB(3) = GetAddress(AddressOf InvokeB)
        m_VTableC(0) = GetAddress(AddressOf QueryInterface)
        m_VTableC(1) = GetAddress(AddressOf AddRef)
        m_VTableC(2) = GetAddress(AddressOf Release)
        m_VTableC(3) = GetAddress(AddressOf InvokeC)
        m_VTable_Initialized = True
    End If
    
    Dim StackInstance As ClassGJ_Layout
    Dim HeapInstance As Long
    
    If VTableSelect = 0 Then StackInstance.pVTable = VarPtr(m_VTableA(0))
    If VTableSelect = 1 Then StackInstance.pVTable = VarPtr(m_VTableB(0))
    If VTableSelect = 2 Then StackInstance.pVTable = VarPtr(m_VTableC(0))
    
    StackInstance.RefC = 1
    Set StackInstance.State = State
    
    HeapInstance = LocalAlloc(0, LenB(StackInstance))
    CopyMemory ByVal HeapInstance, StackInstance, LenB(StackInstance)
    CopyMemory PrivateNewClassGJ, HeapInstance, 4
    
    FillMemory StackInstance.State, 4, 0
End Function

Private Function QueryInterface(This As ClassGJ_Layout, riid As Long, pvObj As Long) As Long
    Debug.Assert False
    Const E_NOINTERFACE As Long = &H80004002
    pvObj = 0
    QueryInterface = E_NOINTERFACE
    Debug.Print "QueryInterface"
End Function

Private Function AddRef(This As ClassGJ_Layout) As Long
This.RefC = This.RefC + 1
Debug.Print "AddRef"
End Function

Private Function Release(This As ClassGJ_Layout) As Long
This.RefC = This.RefC - 1
Debug.Assert This.RefC >= 0
If This.RefC <= 0 Then
    Set This.State = Nothing
    LocalFree VarPtr(This)
End If
Debug.Print "Release"
End Function

Private Function InvokeA(This As ClassGJ_Layout, ByVal errorCode As Long, ByVal OBJ1Address As Long) As Long

    On Error GoTo ERR
    Dim id As Long
    Debug.Print "InvokeA"
    Dim IUnknown1 As IUnknown
    

 'MsgBox "CALLBAKC1 OBJ1Address=" & OBJ1Address
'Static ICoreWebView2Environment As Object
Set This.State.ICoreWebView2EnvironmentObjA = ObjFromPtr(VarPtr(OBJ1Address))
 
Set IUnknown1 = PrivateNewClassGJ(This.State, 1)
Dim Hwnd As Long
Hwnd = This.State.webHostHwnd
'env->CreateCoreWebView2Controller(hWnd, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>
DispCallByVtbl This.State.ICoreWebView2EnvironmentObjA, 4 - 1, Hwnd, ObjPtr(IUnknown1)

'Set This.State = Nothing

'DoEvents
'MsgBox "wait"
Exit Function
ERR:
MsgBox "err-ICoreWebView2Environment ,ID=" & id & ",ERR:" & ERR.Description
End Function

Private Function InvokeB(This As ClassGJ_Layout, ByVal ErrCode As Long, ByVal OBJ1Address As Long) As Long
'IUnknown
    Debug.Print "InvokeB"
On Error GoTo ERR
Dim id As Long
If OBJ1Address = 0 Then
    MsgBox "RUN ON IDE HAVE ERR HERE"
    Exit Function
Else
 'MsgBox "Callback2 OBJ1Address=" & OBJ1Address
End If

'Isso funciona
Set This.State.ICoreWebView2EnvironmentObjB = ObjFromPtr(VarPtr(OBJ1Address))
Set This.State.webviewController = This.State.ICoreWebView2EnvironmentObjB
'Isso não
'Set This.State.webviewController = ObjFromPtr(VarPtr(OBJ1Address))


Dim webviewWindowPtr As Long
Dim r As Long
'For :webviewController->get_CoreWebView2(&webviewWindow);
r = DispCallByVtbl(This.State.webviewController, 23 + 2, VarPtr(webviewWindowPtr))

 Set This.State.webviewWindow = ObjFromPtr(VarPtr(webviewWindowPtr))
This.State.WebViewShowOK = True
    Dim RECT1 As RECT
    GetClientRect This.State.webHostHwnd, RECT1
'webviewController->put_Bounds(bounds);
r = DispCallByVtbl(This.State.webviewController, 4 + 2, RECT1.Left, RECT1.Top, RECT1.Right, RECT1.Bottom)
r = DispCallByVtbl(This.State.webviewWindow, 3 + 2, StrPtr(This.State.Url))

Dim Trash(32) As Byte

Dim IUnknown1 As IUnknown
Set IUnknown1 = PrivateNewClassGJ(This.State, 2)
r = DispCallByVtbl(This.State.webviewWindow, 32 + 2, ObjPtr(IUnknown1), VarPtr(Trash(0)))

This.State.RaiseReady

Exit Function
ERR:
MsgBox "id=" & id & ",CLASS3 ERR:" & ERR.Description
End Function

Private Function InvokeC(This As ClassGJ_Layout, ByVal ErrCode As Long, ByVal OBJ1Address As Long) As Long
Dim r As Variant
Dim Obj As IUnknown
Dim StringPtr As Long
Dim StringLen As Long
Dim Text As String

CopyMemory Obj, OBJ1Address, 4
r = DispCallByVtbl(Obj, 5, VarPtr(StringPtr))

StringLen = lstrlen(StringPtr)
FillMemory Obj, 4, 0

Text = String$(StringLen, 0)
CopyMemory ByVal StrPtr(Text), ByVal StringPtr, StringLen * 2

If Not This.State.WebMessageCallback Is Nothing Then
    This.State.WebMessageCallback.PostWebMessage Text
End If

CoTaskMemFree StringPtr

FillMemory Obj, 4, 0
End Function
'IUnknown
