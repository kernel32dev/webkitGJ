VERSION 5.00
Begin VB.UserControl WebViewGJ 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "WebViewGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal createdEnvironmentCallback As Long) As Long

Dim State As New ClassGJ_State
Dim WebViewCreated As Boolean

Public Sub OpenURL(Url As String)
State.OpenURL Url
CreateWebView
End Sub

Public Sub AddHostObjectToScript(Name As String, Obj As Object)
State.AddHostObjectToScript Name, Obj
End Sub

Public Sub ExecuteScript(Javascript As String)
State.ExecuteScript Javascript
End Sub

Public Sub OpenDevToolsWindow()
State.OpenDevToolsWindow
End Sub

Private Sub UserControl_Resize()
CreateWebView
If Not State Is Nothing Then If State.WebViewShowOK Then State.WebReSize
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
