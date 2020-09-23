VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Basic Text Only Browser"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBrowser 
      Caption         =   "Untitled Document"
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   2280
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin prjHTMLParser.JamieHTMLParser JamieHTMLParser1 
         Height          =   1215
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2143
      End
      Begin RichTextLib.RichTextBox rtbBrowser 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3625
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmBrowser.frx":0000
      End
   End
   Begin VB.Frame fraURL 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "http://www.microsoft.com/default_text.htm"
         Top             =   210
         Width           =   3735
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is by no means a finished web browser
'it supports only a handful of tags and it does
'not even do that very well :)
'
'This example is more for people looking for some
'advanced uses of the parser control, I didn't comment
'as much, the code is pretty self-explanatory.
'
'Basically, we just look for beginning tags, set the
'RTF box the corresponding state and then add our text
'
'I urge everyone to try to make a better web
'browser using the control, it'd be very cool.
Option Explicit

Dim CurrentTag As String
Dim ColorExtract As String
Dim FaceArray() As String
Private Sub Form_Resize()
fraURL.Width = Me.Width - 360
txtURL.Width = fraURL.Width - 720


fraBrowser.Width = Me.Width - 360
rtbBrowser.Width = fraBrowser.Width - 240
fraBrowser.Height = Me.Height - 1185
rtbBrowser.Height = fraBrowser.Height - 360
End Sub

Private Sub JamieHTMLParser1_HTMLProperty(Property As String, PropertyValue As String)
Select Case CurrentTag
    Case "body"
        If LCase(Property) = "bgcolor" Then
            ColorExtract = Right(PropertyValue, 6)
            rtbBrowser.BackColor = RGB(CByte("&H" & Left(ColorExtract, 2)), CByte("&H" & Mid(ColorExtract, 3, 2)), CByte("&H" & Right(ColorExtract, 2)))
        End If
    Case "p"
        If Property = "align" Then
            Select Case LCase(PropertyValue)
                Case "left"
                    rtbBrowser.SelAlignment = 0
                Case "right"
                    rtbBrowser.SelAlignment = 1
                Case "center"
                    rtbBrowser.SelAlignment = 2
                Case Else
                    rtbBrowser.SelAlignment = 0
            End Select
        End If
    Case "div"
        If Property = "align" Then
            Select Case LCase(PropertyValue)
                Case "left"
                    rtbBrowser.SelAlignment = 0
                Case "right"
                    rtbBrowser.SelAlignment = 1
                Case "center"
                    rtbBrowser.SelAlignment = 2
                Case Else
                    rtbBrowser.SelAlignment = 0
            End Select
        End If
    Case "font"
        Select Case LCase(Property)
            Case "face"
                FaceArray = Split(PropertyValue, ",")
                rtbBrowser.SelFontName = FaceArray(0)
            Case "size"
                Select Case Abs(PropertyValue)
                    Case 1
                        rtbBrowser.SelFontSize = 8
                    Case 2
                        rtbBrowser.SelFontSize = 10
                    Case 3
                        rtbBrowser.SelFontSize = 12
                    Case 4
                        rtbBrowser.SelFontSize = 14
                    Case 5
                        rtbBrowser.SelFontSize = 18
                    Case 6
                        rtbBrowser.SelFontSize = 24
                    Case 7
                        rtbBrowser.SelFontSize = 36
                End Select
            Case "color"
                ColorExtract = Right(PropertyValue, 6)
                rtbBrowser.SelColor = RGB(CByte("&H" & Left(ColorExtract, 2)), CByte("&H" & Mid(ColorExtract, 3, 2)), CByte("&H" & Right(ColorExtract, 2)))
                
        End Select
End Select
End Sub

Private Sub JamieHTMLParser1_HTMLTagBegin(Tag As String)
Select Case LCase(Tag)
    Case "br"
        rtbBrowser.SelText = vbCrLf
    Case "b"
        rtbBrowser.SelBold = True
    Case "i"
        rtbBrowser.SelItalic = True
    Case "u"
        rtbBrowser.SelUnderline = True

'Bad emulation of the heading tags
    Case "h1"
        rtbBrowser.SelBold = True
    Case "h2"
        rtbBrowser.SelBold = True
    Case "h3"
        rtbBrowser.SelBold = True
    Case "h4"
        rtbBrowser.SelBold = True
    Case "h5"
        rtbBrowser.SelBold = True
    Case "h6"
        rtbBrowser.SelBold = True
End Select
CurrentTag = LCase(Tag)
End Sub

Private Sub JamieHTMLParser1_HTMLTagClose(Tag As String)
Select Case LCase(Tag)
    Case "b"
        rtbBrowser.SelBold = False
    Case "i"
        rtbBrowser.SelItalic = False
    Case "u"
        rtbBrowser.SelUnderline = False
    Case "div"
        rtbBrowser.SelText = vbCrLf
'Again, needs to be fixed
    Case "h1"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
    Case "h2"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
    Case "h3"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
    Case "h4"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
    Case "h5"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
    Case "h6"
        rtbBrowser.SelText = vbCrLf
        rtbBrowser.SelBold = False
End Select
End Sub

Private Sub JamieHTMLParser1_HTMLText(Text As String)
Text = Replace(Text, "&nbsp;", " ")
Text = Replace(Text, "&amp;", "&")
Text = Replace(Text, "&reg;", "®")
Select Case LCase(CurrentTag)
    Case "title"
        fraBrowser.Caption = Trim(Text)
    Case Else
        rtbBrowser.SelText = Trim(Text)
End Select
End Sub

Private Sub txtURL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    rtbBrowser.Text = ""
    JamieHTMLParser1.ParseHTML (Inet1.OpenURL(Trim(txtURL.Text)))
End If
End Sub

