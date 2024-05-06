VERSION 5.00
Begin VB.UserControl WebViewGJ 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "WebViewGJ.ctx":0000
   Begin VB.Timer JSCallTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer ReadyTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "WebViewGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal createdEnvironmentCallback As Long) As Long

Event Ready()
Event JSCall(StrArg As String, ByVal NumArg As Long)

Dim WithEvents State As ClassGJ_State
Attribute State.VB_VarHelpID = -1
Dim WithEvents JSCall As JSCall
Attribute JSCall.VB_VarHelpID = -1
Dim WebViewCreated As Boolean

Dim LastJSCallStr As String
Dim LastJSCallNum As Long

Public Sub Dispose()
Set State = Nothing
Set JSCall = Nothing
WebViewCreated = False
End Sub

Public Sub OpenURL(Url As String)
If State Is Nothing Then Set State = New ClassGJ_State
State.OpenURL Url
CreateWebView
End Sub

Private Sub JSCall_JavascriptCalled(StrArg As String, ByVal NumArg As Long)
If JSCallTimer.Enabled Then JSCallTimer_Timer
LastJSCallStr = StrArg
LastJSCallNum = NumArg
JSCallTimer.Enabled = True
End Sub

Private Sub JSCallTimer_Timer()
JSCallTimer.Enabled = False
RaiseEvent JSCall(LastJSCallStr, LastJSCallNum)
LastJSCallStr = vbNullString
LastJSCallNum = 0
End Sub

Private Sub ReadyTimer_Timer()
ReadyTimer.Enabled = False
If State Is Nothing Then Exit Sub
Set JSCall = New JSCall
AddHostObjectToScript "JSCall", JSCall
RaiseEvent Ready
End Sub

Private Sub State_Ready()
ReadyTimer.Enabled = True
End Sub

Private Sub UserControl_Resize()
If State Is Nothing Then Exit Sub
CreateWebView
If State.WebViewShowOK Then State.WebReSize
End Sub

Private Sub CreateWebView()

If State Is Nothing Then Exit Sub

If WebViewCreated Then Exit Sub
WebViewCreated = True

On Error Resume Next
MkDir App.Path & "\userdata\"
On Error GoTo 0

State.webHostHwnd = Hwnd

Dim userdata As String
Dim edgesdk As String

userdata = App.Path & "\userdata\"

Dim WebCompletedHandler As IUnknown
Set WebCompletedHandler = PrivateNewClassGJ(State, False)

If CreateCoreWebView2EnvironmentWithOptions(StrPtr(vbNullString), StrPtr(userdata), 0&, ObjPtr(WebCompletedHandler)) <> S_OK Then
    MsgBox "Failed to create webview environment", vbOKOnly, "Error"
End If
 
End Sub

Sub OpenDevToolsWindow()
    State.DeferableWindowVTableCallEx "OpenDevToolsWindow", 49
End Sub

Sub ExecuteScript(JS As String)
    State.DeferableWindowVTableCallEx "ExecuteScript", 27, StrPtr(JS), 0&
End Sub

Sub AddHostObjectToScript(ObjName As String, Obj1 As Object)
    State.DeferableWindowVTableCallEx "AddHostObjectToScript", 47, StrPtr(ObjName), ObjPtr(Obj1)
End Sub

