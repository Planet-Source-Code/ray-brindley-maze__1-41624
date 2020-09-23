VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   7875
   ClientLeft      =   2145
   ClientTop       =   2010
   ClientWidth     =   10500
   Icon            =   "MenuChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "MenuChar.frx":08CA
   ScaleHeight     =   7215.313
   ScaleMode       =   0  'User
   ScaleWidth      =   9786.958
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPlay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7875
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   10500
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblPlayExit 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   7320
         MousePointer    =   10  'Up Arrow
         TabIndex        =   15
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label lblPlayHelp 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   2040
         MousePointer    =   10  'Up Arrow
         TabIndex        =   14
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label lblSubmit 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   4440
         MousePointer    =   10  'Up Arrow
         TabIndex        =   13
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Image imgChar1 
         Height          =   1050
         Left            =   3360
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":1194
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Image imgCharI1 
         Height          =   1050
         Left            =   3360
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":1B5F
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgChar2 
         Height          =   1050
         Left            =   4680
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":26DF
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Image imgCharI2 
         Height          =   1050
         Left            =   4680
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":2EFB
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgChar3 
         Height          =   1050
         Left            =   6000
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":372D
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Image imgCharI3 
         Height          =   1050
         Left            =   6000
         MouseIcon       =   "MenuChar.frx":40B2
         MousePointer    =   10  'Up Arrow
         Picture         =   "MenuChar.frx":497C
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgCharO3 
         Height          =   1050
         Left            =   6000
         Picture         =   "MenuChar.frx":5467
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgCharO2 
         Height          =   1050
         Left            =   4680
         Picture         =   "MenuChar.frx":60A3
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgCharO1 
         Height          =   1050
         Left            =   3360
         Picture         =   "MenuChar.frx":6926
         Top             =   4560
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Garu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   5
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MashiMaro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   6
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pucca"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   7
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lblPickChar 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the character of your choice:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   4200
         Width           =   3735
      End
      Begin VB.Label lblTypeName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Type in your name [max. 10 characters]:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Image imgGettingStarted 
         Height          =   7875
         Left            =   0
         Picture         =   "MenuChar.frx":76FE
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10500
      Begin VB.Label lblMainAbout 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   1320
         MouseIcon       =   "MenuChar.frx":33DDB
         MousePointer    =   10  'Up Arrow
         TabIndex        =   12
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblMainExit 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   7320
         MouseIcon       =   "MenuChar.frx":346A5
         MousePointer    =   10  'Up Arrow
         TabIndex        =   11
         Top             =   6120
         Width           =   1425
      End
      Begin VB.Label lblMainHelp 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   7680
         MouseIcon       =   "MenuChar.frx":34F6F
         MousePointer    =   10  'Up Arrow
         TabIndex        =   10
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label lblMainHighScores 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Left            =   3720
         MouseIcon       =   "MenuChar.frx":35839
         MousePointer    =   10  'Up Arrow
         TabIndex        =   9
         Top             =   3360
         Width           =   3240
      End
      Begin VB.Label lblMainPlay 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   1920
         MouseIcon       =   "MenuChar.frx":36103
         MousePointer    =   10  'Up Arrow
         TabIndex        =   8
         Top             =   960
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   7875
         Left            =   0
         Picture         =   "MenuChar.frx":369CD
         Top             =   0
         Width           =   10500
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'**********************************************************************************'
'Ok a bit about myself... My name is Ray and i enjoy programming. I am not the best'
'but i do my best and i always am willing to try and experience new things. I don't'
'get everything right in programming but i don't often give up unless i can't teach'
'myself as i have taught myself basically everything i know......School gave me the'
'comeplete basics and i read help file after help file and other things from planet'
'source code which have helped me out. I made this little program for a girl in USA'
'she asked me for this. i made her do most of the graphics cos i really suck at em '
'i don't often make something unless i am asked to or have a need to....I used the '
'same basics from one of my earlier programs called rayman a pacman game which did '
'pretty well i thought. this has the same basics and is just a little more advanced'
'i had never used a database EVER before so that i had to learn quite a lot about. '
'But i managed and i think it worked out well. Other than that most of it is pretty'
'basic but i quite proud of it considering it is almost all my own work if not it's'
'all edited to suit my needs. and as far as i know it all works except for the stuf'
'that i didn't understand but that i couldn't be bothered with and the girl from US'
'wanted it quicker than i could have figured it out. So i left it as is and it was '
'good enough for me. Anyway Hope you enjoy it. And by all means vote for me at PSC '
'if you liked it...Hell even if you didn't like it vote...Anything i can improve on'
'         tell me. I would love some feedback either on PSC or e-mail me @         '
'                            brindleyray@hotmail.com                               '
'**********************************************************************************'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'

Sub Form_Load()
  For J = 0 To 2  'makes it simpler to do everything the same color not really nessesary
    lblChar(J).ForeColor = RGB(77, 104, 173)
  Next J
    
  lblTypeName.ForeColor = RGB(77, 104, 173)
  lblPickChar.ForeColor = RGB(77, 104, 173)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = True     'stops the program from exiting until it is confirmed
  Quit              'calls a function so i don't have to copy the same thing over and over again
End Sub

Private Sub imgChar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Char1_MouseMove       'Calls the sub, i put them all in a module
                        'simply because it makes it a lot simpler
                        'to see and understand. i didn't make this
                        'part i just edited it to make it neater.
                        'from here down is really very simple so i left it uncommented
End Sub

Private Sub imgChar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Char2_MouseMove
End Sub

Private Sub imgChar3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Char3_MouseMove
End Sub

Private Sub imgCharI1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CharI1_MouseDown
End Sub

Private Sub imgCharI2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CharI2_MouseDown
End Sub

Private Sub imgCharI3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CharI3_MouseDown
End Sub

Private Sub imgGettingStarted_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  GettingStarted
End Sub

Sub lblMainAbout_Click()
  frmAbout.Visible = True
End Sub

Sub lblMainExit_Click()
  Quit
End Sub

Sub lblMainHelp_Click()
  frmHelp.Visible = True
End Sub

Sub lblMainHighScores_Click()
  frmHighScores.Visible = True
End Sub

Sub lblMainPlay_Click()
  fraMain.Visible = False
  fraPlay.Visible = True
  frmMenu.Caption = "Character Select"
End Sub

Sub lblPlayExit_Click()
  Quit
End Sub

Sub lblPlayHelp_Click()
  frmHelp.Visible = True
End Sub

Sub lblSubmit_Click()
  frmMaze.LblName.Caption = txtName.Text
  If txtName.Text = "" Then
    MsgBox ("Please enter a name."), vbOKOnly, "No Name"    'Simple message box routine
    Exit Sub
  End If
  If imgCharO1.Visible = False And imgCharO2.Visible = False And imgCharO3.Visible = False Then
    MsgBox ("Please select a character to play as."), vbOKOnly, "No Character"
    Exit Sub                  'quits this sub so that they enter a name and select a character
  End If
  fraMain.Visible = True      'puts everything back to normal for a restart
  fraPlay.Visible = False
  imgChar1.Visible = True
  imgCharI1.Visible = False
  imgCharO1.Visible = False
  imgChar2.Visible = True
  imgCharI2.Visible = False
  imgCharO2.Visible = False
  imgChar3.Visible = True
  imgCharI3.Visible = False
  imgCharO3.Visible = False
  txtName.Text = ""
  frmMenu.Visible = False
  MakeGrid                    'calls the sub to start everything
  frmMaze.Visible = True
End Sub
