VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6_Edge WebView2 Sample"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   17310
   StartUpPosition =   2  'CenterScreen
   Begin VB6_WebView2.WebViewGJ WebViewGJ1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3413
   End
   Begin VB.CommandButton BTOpenUrl 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OpenUrl"
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox UrlTxt 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Text            =   "https://www.baidu.com"
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton BTCreateWebView 
      Caption         =   "CreateWebView"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Menu DoOpenDevToolsWindow 
      Caption         =   "OpenDevToolsWindow"
   End
   Begin VB.Menu DoExecuteScript 
      Caption         =   "ExecuteScript"
   End
   Begin VB.Menu DoAddHostObjectToScript 
      Caption         =   "AddHostObjectToScript"
   End
   Begin VB.Menu doObjectTest 
      Caption         =   "jsObjectTest"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' WebView2Loader.dll can download from here:
'www.vbrichclient.com/en/Downloads.htm

Public Function GetCaption() As String
GetCaption = Me.Caption
End Function

Private Sub DoAddHostObjectToScript_Click()
WebViewGJ1.AddHostObjectToScript "Form1", Me
End Sub

Private Sub DoExecuteScript_Click()
MsgBox "Runjs alert(33+44)"
WebViewGJ1.ExecuteScript "alert(33+44)"
End Sub

Private Sub doObjectTest_Click()
'WebViewGJ1.ExecuteScript "const ObjForm1=window.chrome.webview.hostObjects.Form1;alert(ObjForm1.GetCaption)"
WebViewGJ1.ExecuteScript "const ObjForm1=window.chrome.webview.hostObjects.Form1;ObjForm1.GetCaption().then(alert)"
'error ,need fix
'const ObjForm1=window.chrome.webview.hostObjects.Form1;alert(ObjForm1.GetCaption);
'var result=await ObjForm1.GetCaption();alert(result)
End Sub

Private Sub DoOpenDevToolsWindow_Click()
WebViewGJ1.OpenDevToolsWindow
End Sub

Private Sub BTOpenUrl_Click()
WebViewGJ1.OpenURL UrlTxt.Text
End Sub

Private Sub Form_Load()
WebViewGJ1.OpenURL "google.com"
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    WebViewGJ1.Move WebViewGJ1.Left, WebViewGJ1.Top, ScaleWidth - WebViewGJ1.Left - 50, ScaleHeight - WebViewGJ1.Top - 50
End If
End Sub

Private Sub UrlTxt_DblClick()
UrlTxt.Text = ""
End Sub

Private Sub UrlTxt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    WebViewGJ1.OpenURL UrlTxt.Text
End If
End Sub
