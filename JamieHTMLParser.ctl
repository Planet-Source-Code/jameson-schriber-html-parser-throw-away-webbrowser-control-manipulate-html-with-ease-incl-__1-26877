VERSION 5.00
Begin VB.UserControl JamieHTMLParser 
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   1350
   ScaleWidth      =   2190
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "JamieHTMLParser.ctx":0000
      Top             =   680
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "< v 1.0.0.0 >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "< Jamie HTML Parser >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "JamieHTMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ta da!
'This is first control of its kind on PSC. It parses
'HTML and then fires events on the beginning of tags,
'tag properties and their values, the ending of tags
'(the end of the beginning tag), the closing of
'tags (</close>), and the text in between tags
'allowing you to do practically everything with HTML.
'Yeehaw! Extract information with ease, manipulate tags
'like never before, create syntax highlighters, even
'your own web browser!
'
'Please vote and always support PSC
'
'Copyright 2001 - Jameson Schriber
'JamesonSchriber@aol.com

Option Explicit

Public Event HTMLTagBegin(Tag As String)
Attribute HTMLTagBegin.VB_Description = "Fires when a tag begins"
Public Event HTMLProperty(Property As String, PropertyValue As String)
Attribute HTMLProperty.VB_Description = "Returns a property's name and value"
Public Event HTMLTagEnd(Tag As String)
Attribute HTMLTagEnd.VB_Description = "Fires when a tag ends, i.e. tag begin, properties and values, tag end,html text, tag close"
Public Event HTMLTagClose(Tag As String)
Attribute HTMLTagClose.VB_Description = "Fires when a tag closes i.e. </"
Public Event HTMLText(Text As String)
Attribute HTMLText.VB_Description = "Fires when text in between HTML tags is found"
Public Sub ParseHTML(HTML As String)
Attribute ParseHTML.VB_Description = "Parses HTML and fires HTML events"
'We go through the HTML, character by character
'checking first for <, then for spaces, then
'quotation marks, and finally /. As we find
'them we fire events and continue parsing.
'
'Clean code with few relevant comments is better than
'unwieldy code commented to death, IMHO
'
Dim IsValue, IsProperty, IsTag, RaisedTagBegin As Boolean
Dim i As Long
Dim CurrentChar As String
Dim CurrentProperty As String
Dim CurrentPropertyValue As String
Dim CurrentTag As String
Dim CurrentText As String
'Remove tabs and returns, they have no place in HTML
HTML = Replace(HTML, vbCrLf, "")
HTML = Replace(HTML, vbTab, "")
'Start our searching
For i = 1 To Len(HTML)
    CurrentChar = Mid(HTML, i, 1)
    If IsTag = True Then
        If IsProperty = True Then
            If IsValue = True Then
                If CurrentChar = Chr(34) Then
                    IsValue = False
                    IsProperty = False
                    CurrentPropertyValue = Trim(CurrentPropertyValue)
                    CurrentProperty = Trim(CurrentProperty)
                    RaiseEvent HTMLProperty(Left(CurrentProperty, Len(CurrentProperty) - 1), CurrentPropertyValue)
                    CurrentPropertyValue = ""
                    CurrentProperty = ""
                Else
                    CurrentPropertyValue = CurrentPropertyValue & CurrentChar
                End If
            ElseIf CurrentChar = Chr(34) Then
                IsValue = True
            Else
                CurrentProperty = CurrentProperty & CurrentChar
            End If
        Else
            If CurrentChar = " " Then
                IsProperty = True
                CurrentTag = Trim(CurrentTag)
                CurrentTag = CurrentTag
                If RaisedTagBegin = False Then
                    RaiseEvent HTMLTagBegin(CurrentTag)
                    RaisedTagBegin = True
                End If
            ElseIf CurrentChar = ">" Then
                IsTag = False
                If Left(CurrentTag, 1) = "/" Then
                    RaiseEvent HTMLTagClose(Right(CurrentTag, Len(CurrentTag) - 1))
                ElseIf RaisedTagBegin = False Then
                    RaiseEvent HTMLTagBegin(CurrentTag)
                    RaiseEvent HTMLTagEnd(CurrentTag)
                    RaisedTagBegin = True
                Else
                    RaiseEvent HTMLTagEnd(CurrentTag)
                End If
                CurrentTag = ""
                
            Else
                CurrentTag = CurrentTag & CurrentChar
            End If
        End If
    Else
        If CurrentChar = "<" Then
            IsTag = True
            RaisedTagBegin = False
            If Trim(CurrentText) <> "" Then
                RaiseEvent HTMLText(Trim(CurrentText))
                CurrentText = ""
            End If
        Else
            CurrentText = CurrentText & CurrentChar
        End If
    End If
Next i
End Sub

