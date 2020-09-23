VERSION 5.00
Begin VB.Form frmPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Which demo would you like?"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "&Basic Text Only Web Browser - No WebBrowser Control"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdSyntax 
      Caption         =   "&Syntax Highlight with Proper Formatting Demo"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just to display the demos
Private Sub cmdBrowser_Click()
frmBrowser.Show
End Sub

Private Sub cmdSyntax_Click()
frmFormat.Show
End Sub

