VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFormat 
   Caption         =   "Proper Formatting and Syntax Highlighting Demo"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5295
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin prjHTMLParser.JamieHTMLParser JamieHTMLParser1 
      Height          =   1695
      Left            =   4800
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2990
   End
   Begin VB.Frame fraFormatted 
      Caption         =   "Formatted with Syntax Highlighting"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   6855
      Begin RichTextLib.RichTextBox rtbFormatted 
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmFormat.frx":0000
      End
   End
   Begin VB.Frame fraUnformatted 
      Caption         =   "Unformatted HTML"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtUnformatted 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Keep track of our indenting
Option Explicit

Dim Indent As Integer

Private Sub Form_Resize()
'Resize our controls
fraUnformatted.Width = Me.Width - 360
fraFormatted.Width = Me.Width - 360

fraUnformatted.Height = Me.Height / 2 - 360
fraFormatted.Top = Me.Height / 2 - 120
fraFormatted.Height = Me.Height / 2 - 360

txtUnformatted.Width = fraUnformatted.Width - 240
rtbFormatted.Width = fraFormatted.Width - 240

txtUnformatted.Height = fraUnformatted.Height - 360
rtbFormatted.Height = fraFormatted.Height - 360
End Sub

Private Sub JamieHTMLParser1_HTMLProperty(Property As String, PropertyValue As String)
'Highlight the values using ValueColor, than
'return the color to BaseColor
With rtbFormatted
    .SelText = " " & Property & "=" & Chr(34)
    .SelColor = &HFF&
    .SelText = PropertyValue
    .SelColor = &HFF0000
    .SelText = Chr(34)
End With
End Sub

Private Sub JamieHTMLParser1_HTMLTagBegin(Tag As String)
'If our tag is a comment, highlight with CommentColor
'Comments don't have closing tags, so we have to
'subtract 1 from our indent
If Left(Tag, 3) = "!--" Then
    rtbFormatted.SelColor = &HC0C0C0
    Indent = Indent - 1
Else
    rtbFormatted.SelColor = &HFF0000
    Select Case LCase(Tag)
        'these tags don't have closing tags i.e. </ so
        'we have to decrease the indent manually
        Case "meta"
            Indent = Indent - 1
        Case "img"
            Indent = Indent - 1
        Case "br"
            Indent = Indent - 1
    End Select
End If
'add our tag to the RTF box
rtbFormatted.SelText = TabIndent(Indent) & "<" & Tag
'always increase the indent
Indent = Indent + 1
End Sub

Private Sub JamieHTMLParser1_HTMLTagClose(Tag As String)
'decrease our indent
Indent = Indent - 1
'change our text color back to BaseColor and then
'add our closing tag
rtbFormatted.SelColor = &HFF0000
rtbFormatted.SelText = TabIndent(Indent) & "</" & Tag & ">" & vbCrLf
End Sub

Private Sub JamieHTMLParser1_HTMLTagEnd(Tag As String)
'add the final > to our tag
rtbFormatted.SelText = ">" & vbCrLf
End Sub

Private Sub JamieHTMLParser1_HTMLText(Text As String)
'change the text color to TextColor and then add our
'text
rtbFormatted.SelColor = &H0&
rtbFormatted.SelText = TabIndent(Indent) & Text & vbCrLf
End Sub

Private Sub txtUnformatted_Change()
'trigger our syntax highlighting and parser, set our
'indent
Indent = 0
rtbFormatted.Text = ""
JamieHTMLParser1.ParseHTML (Trim(txtUnformatted.Text))
End Sub
Private Function TabIndent(Number As Integer)
'useful little function, adds two spaces as an indent
If Number >= 0 Then
    TabIndent = Replace(Space(Number * 2), " ", " ")
End If
End Function
