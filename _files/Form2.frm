VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12390
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   10095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   360
   End
   Begin webkitGJ.WebViewGJ WebViewGJ1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      _ExtentX        =   1720
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
