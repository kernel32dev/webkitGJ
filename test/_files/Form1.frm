VERSION 5.00
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11700
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   10095
   End
   Begin webkitGJ.WebViewGJ WebViewGJ1 
      Height          =   6735
      Left            =   600
      TabIndex        =   0
      Top             =   540
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PaginaInicial As String = "google.com"

Private Sub Form_Load()
Text1 = PaginaInicial
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

Private Sub Timer1_Timer()
Timer1 = False
WebViewGJ1.OpenURL PaginaInicial
End Sub

