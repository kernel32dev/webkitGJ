VERSION 5.00
Object = "*\A..\..\VB6_WebView2.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin webkitGJ.WebViewGJ WebViewGJ1 
      Height          =   6735
      Left            =   180
      TabIndex        =   3
      Top             =   660
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DevTools"
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dispose"
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11700
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  '   WebViewGJ1.Dispose
End Sub

Private Sub Command2_Click()
WebViewGJ1.OpenDevToolsWindow
WebViewGJ1.AddHostObjectToScript "eu", Me
End Sub

Private Sub Form_Load()
Text1 = "google.com"
WebViewGJ1.OpenURL Text1

End Sub

Private Sub Form_Resize()
WebViewGJ1.Move 0, WebViewGJ1.Top, ScaleWidth, ScaleHeight - WebViewGJ1.Top
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    WebViewGJ1.OpenURL Text1
End If
End Sub

Private Sub WebViewGJ1_JSCall(StrArg As String, ByVal NumArg As Long)
MsgBox "StrArg = " & StrArg & vbNewLine & "NumArg = " & NumArg
End Sub

Private Sub WebViewGJ1_WebMessagePosted(Message As String)
    MsgBox Message
End Sub
