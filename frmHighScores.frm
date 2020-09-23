VERSION 5.00
Begin VB.Form frmHighScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HighScores"
   ClientHeight    =   7845
   ClientLeft      =   2310
   ClientTop       =   1830
   ClientWidth     =   10500
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHighScores.frx":08CA
   ScaleHeight     =   7845
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5280
      TabIndex        =   29
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   28
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5280
      TabIndex        =   27
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   26
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   25
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   24
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   23
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   21
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   19
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   5880
      TabIndex        =   18
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   17
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   16
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   15
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   12
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblNames 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   5880
      TabIndex        =   10
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   3480
      TabIndex        =   9
      Top             =   6240
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   8
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   3480
      TabIndex        =   7
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   6
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   4
      Top             =   4080
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   2
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   8400
      TabIndex        =   0
      Top             =   6600
      Width           =   1005
   End
End
Attribute VB_Name = "frmHighScores"
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

Private Sub Form_Load()
  ShowScores
End Sub
'another simple file reading example

Public Sub ShowScores()
Dim Name, Score
  Scoresfile = FreeFile
  Open "highscores.hss" For Input As #Scoresfile
    For I = 1 To 10 'selects all of the lines.
      Input #Scoresfile, Name, Score  'In the file the name and score are separated by comma's so you have to say which one is which by separating them in here
      lblNames(I).Caption = Name    'Put the name of the line into the corresponding label
      lblScores(I).Caption = Score  'Put the score of the line into the corresponding label
    Next I  'repeat process untill I = 10
  Close Scoresfile  'close the file
End Sub
