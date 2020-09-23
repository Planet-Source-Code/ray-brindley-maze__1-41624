VERSION 5.00
Begin VB.Form frmQuestions 
   Caption         =   "Select the correct answer to continue"
   ClientHeight    =   5175
   ClientLeft      =   3195
   ClientTop       =   3240
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmQuestions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   5625
   Begin VB.Timer tmrImage 
      Interval        =   100
      Left            =   4080
      Top             =   3840
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   120
      Top             =   3960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox txtFields 
         DataField       =   "Answer1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Answer2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   915
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Answer3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1245
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Answer4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
      End
      Begin VB.OptionButton opt 
         Caption         =   "a)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton opt 
         Caption         =   "b)"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   915
         Width           =   615
      End
      Begin VB.OptionButton opt 
         Caption         =   "c)"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1245
         Width           =   615
      End
      Begin VB.OptionButton opt 
         Caption         =   "d)"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Questions"
      Top             =   4830
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00FF99CC&
      DataField       =   "Correct"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Question"
      DataSource      =   "Data1"
      Height          =   1320
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   40
      Width           =   3375
   End
   Begin VB.Image imgQuestions 
      Height          =   495
      Left            =   2040
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   0
      Left            =   5520
      Picture         =   "frmQuestions.frx":08CA
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   1
      Left            =   5520
      Picture         =   "frmQuestions.frx":0DFD
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   2
      Left            =   5520
      Picture         =   "frmQuestions.frx":12EF
      Top             =   2400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   3
      Left            =   5520
      Picture         =   "frmQuestions.frx":1843
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   4
      Left            =   5520
      Picture         =   "frmQuestions.frx":1CF4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   5
      Left            =   5520
      Picture         =   "frmQuestions.frx":21EC
      Top             =   3840
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   6
      Left            =   5520
      Picture         =   "frmQuestions.frx":269D
      Top             =   4320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   8
      Left            =   5520
      Picture         =   "frmQuestions.frx":2BF1
      Top             =   5280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   720
      Index           =   7
      Left            =   5520
      Picture         =   "frmQuestions.frx":3124
      Top             =   4800
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "60"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Question:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmQuestions"
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

Private Sub cmdClose_Click()
  Unload Me     'closes the form not the project
End Sub

Private Sub Form_Load()
  Maths     'Call sub
  TotalTime = 60  'set everything up
  Current = 0
  Start = Timer
End Sub

Private Sub opt_Click(Index As Integer)
  tmrTime.Enabled = False
                            
  If txtFields(Index).Text = txtFields(5).Text Then   'if they chose correctly
  
    MsgBox "Correct, Congradulations.", vbOKOnly, "Correct!"      'simple message box
    frmMaze.lblScore.Caption = frmMaze.lblScore.Caption + Int(lblTime.Caption / 6)  'add some points to there score for getting it right
    
    Maths     'set up the next question
    frmMaze.lblCover.BackStyle = 0
    frmMaze.tmrTime.Enabled = True  'start the timer on the form
    frmMaze.Enabled = True
    frmQuestions.Hide
        
    lblTime.Caption = 60
    opt(1).Value = False    'make sure that it doesn't automatically
    opt(2).Value = False    'choose another answer when it opens again
    opt(3).Value = False
    opt(4).Value = False
  Else  'if they answered wrong
    MsgBox "Incorrect sorry. The answer was " & txtFields(5).Text & ".", vbOKOnly, "Incorrect Answer" 'simple message box
    frmMaze.lblScore.Caption = frmMaze.lblScore.Caption - 25  'heavy penalty
    If frmMaze.lblScore.Caption < 0 Then    'if the score is less than 0 then gameover
      frmMaze.lblScore.Caption = 0    'don't let them go into the negative numbers
    End If
    
    Maths   'if the score wasn't under 0 set up for next question
    frmMaze.lblCover.BackStyle = 0
    frmMaze.tmrTime.Enabled = True
    frmMaze.Show
    frmMaze.Enabled = True
    frmQuestions.Hide
    
    lblTime.Caption = 60
    opt(1).Value = False
    opt(2).Value = False
    opt(3).Value = False
    opt(4).Value = False
  End If
End Sub

Private Sub tmrImage_Timer()
  Static Frame As Integer     'create an animated gif. cos i don't know why
                              'but vb doesn't like animated gifs working for themselves
  imgQuestions.Picture = imgQuestion(Frame).Picture 'make the empty picture have another picture
  Frame = Frame + 1     'make it so on the next time through it is a different picture
  If Frame = 8 Then     'if it is at the end of the cycle
    Frame = 0           'go back to the beggining
  End If
End Sub

Private Sub tmrTime_Timer()
  lblTime.Caption = lblTime.Caption - 1   'countdown to doom
  If lblTime.Caption < 0 Then     'if they run out of time
    tmrTime.Enabled = False
    Gameover         'there seems to be alot of ways to lose doesn't there
  End If
End Sub
