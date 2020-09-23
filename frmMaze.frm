VERSION 5.00
Begin VB.Form frmMaze 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Maze"
   ClientHeight    =   8265
   ClientLeft      =   585
   ClientTop       =   630
   ClientWidth     =   11370
   DrawWidth       =   5
   Icon            =   "frmMaze.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   758
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10080
      Top             =   7560
   End
   Begin VB.Timer tmrBonuses 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   7080
   End
   Begin VB.PictureBox picPucca 
      Height          =   255
      Left            =   8040
      Picture         =   "frmMaze.frx":08CA
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picMisha 
      Height          =   255
      Left            =   8040
      Picture         =   "frmMaze.frx":0E0C
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picGaru 
      Height          =   255
      Left            =   8040
      Picture         =   "frmMaze.frx":134E
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   7560
   End
   Begin VB.PictureBox pic 
      Height          =   255
      Left            =   8040
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCover 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblFinish 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblReturn 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image imgBanana 
      Height          =   195
      Left            =   9000
      Picture         =   "frmMaze.frx":1890
      Stretch         =   -1  'True
      Top             =   6960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgApple 
      Height          =   195
      Left            =   9000
      Picture         =   "frmMaze.frx":1992
      Stretch         =   -1  'True
      Top             =   7920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgBeforeQu 
      Height          =   1530
      Left            =   8880
      Picture         =   "frmMaze.frx":1A94
      Top             =   360
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image imgQuestion 
      Height          =   1530
      Left            =   8880
      Picture         =   "frmMaze.frx":273D
      Top             =   360
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "300"
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image imgMaSi 
      Height          =   6840
      Left            =   10680
      Picture         =   "frmMaze.frx":3321
      Top             =   4800
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgMaBo 
      Height          =   510
      Left            =   10200
      Picture         =   "frmMaze.frx":39D3
      Top             =   5280
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Image imgMaTo 
      Height          =   510
      Left            =   10200
      Picture         =   "frmMaze.frx":4B35
      Top             =   4800
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Image imgPuSi 
      Height          =   6840
      Left            =   10680
      Picture         =   "frmMaze.frx":58E5
      Top             =   3480
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgPuBo 
      Height          =   510
      Left            =   10080
      Picture         =   "frmMaze.frx":6066
      Top             =   3960
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Image imgPuTo 
      Height          =   510
      Left            =   10080
      Picture         =   "frmMaze.frx":7159
      Top             =   3480
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Image imgGaSi 
      Height          =   6825
      Left            =   10680
      Picture         =   "frmMaze.frx":7ECE
      Top             =   2160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgGaTo 
      Height          =   525
      Left            =   10080
      Picture         =   "frmMaze.frx":858D
      Top             =   2160
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Image imgGaBo 
      Height          =   525
      Left            =   10080
      Picture         =   "frmMaze.frx":93E8
      Top             =   2640
      Visible         =   0   'False
      Width           =   10500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Timeleft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image imgRight 
      Height          =   7200
      Left            =   7680
      Picture         =   "frmMaze.frx":A090
      Stretch         =   -1  'True
      Top             =   510
      Width           =   120
   End
   Begin VB.Image imgBottom 
      Height          =   525
      Left            =   0
      Picture         =   "frmMaze.frx":A74F
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   7800
   End
   Begin VB.Image imgLeft 
      Height          =   7200
      Left            =   0
      Picture         =   "frmMaze.frx":B3F7
      Stretch         =   -1  'True
      Top             =   510
      Width           =   120
   End
   Begin VB.Image imgTop 
      Height          =   525
      Left            =   0
      Picture         =   "frmMaze.frx":BAB6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7800
   End
   Begin VB.Label lblsec2 
      BackColor       =   &H80000009&
      Caption         =   "0"
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
      Left            =   9660
      TabIndex        =   5
      Top             =   4320
      Width           =   195
   End
   Begin VB.Label lblSec1 
      BackColor       =   &H80000009&
      Caption         =   "0"
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
      Left            =   9480
      TabIndex        =   4
      Top             =   4320
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   ":"
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
      Left            =   9360
      TabIndex        =   3
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Image imgPucca 
      Height          =   735
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image imgRabbit 
      Height          =   150
      Left            =   9240
      Picture         =   "frmMaze.frx":C911
      Stretch         =   -1  'True
      Top             =   7920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgEvil 
      Height          =   150
      Left            =   9240
      Picture         =   "frmMaze.frx":CB5B
      Stretch         =   -1  'True
      Top             =   7680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgBoy 
      Height          =   150
      Left            =   9240
      Picture         =   "frmMaze.frx":CD71
      Stretch         =   -1  'True
      Top             =   7440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgGirl 
      Height          =   150
      Left            =   9240
      Picture         =   "frmMaze.frx":CFBB
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblMin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   9000
      TabIndex        =   1
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgChar 
      Height          =   150
      Left            =   9240
      Picture         =   "frmMaze.frx":D205
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   150
   End
End
Attribute VB_Name = "frmMaze"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode   'simple way to do key presses
    Case vbKeyUp        'which key
      MoveUp            'calls the procedure
    Case vbKeyDown
      MoveDown
    Case vbKeyLeft
      MoveLeft
    Case vbKeyRight
      MoveRight
  End Select
End Sub

Private Sub Form_Load()
  Organize      'just a procedure to make this part neater
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = True 'stops unloading
  Quit          'Selecting whether to quit or not...
End Sub

Private Sub lblHelp_Click()
  tmrTime.Enabled = False   'So they aren't penalized for looking at the help form
  frmHelp.Show              'simple show the help form
End Sub

Private Sub lblNew_Click()
  MakeGrid      'call procedure
  tmrTime.Enabled = False
End Sub

Private Sub lblPause_Click()
  If lblPause.Caption = "Pause" Then
    lblPause.Caption = "Resume"
    tmrTime.Enabled = False
    lblCover.BackStyle = 1
  ElseIf lblPause.Caption = "Resume" Then
    lblPause.Caption = "Pause"
    tmrTime.Enabled = True
    lblCover.BackStyle = 0
  End If
End Sub

Private Sub lblReturn_Click()
  Response = MsgBox("Are you sure you wish to go back...doing so will end your current game!?", vbYesNo, "Warning")
  If Response = vbYes Then
    Me.Hide       'don't want to close it cos not needed
    frmMenu.Show
    Organize
  Else
    Exit Sub
  End If
End Sub

Private Sub tmrBonuses_Timer()
  Select Case Images      'didn't get this part working was supposed to make bonuses
    Case 1                'pop up randomly but didn't get around to it
      imgApple.Visible = True
    Case 2
      imgBanana.Visible = True
  End Select
  Randomize Images
  If imgApple.Visible = True Then
    imgBanana.Visible = False
  ElseIf imgBanana.Visible = True Then
    imgApple.Visible = False
  End If
End Sub

Private Sub tmrStart_Timer()  'To show the start and finish places
  lblFinish.Visible = False
  lblStart.Visible = False
  tmrStart.Enabled = False
End Sub

Private Sub tmrTime_Timer()
  If lblSec1.Caption = 0 And lblsec2.Caption = 0 Then 'count down the minutes
    lblsec2.Caption = 10   'make sure that these go down with it
    lblSec1.Caption = 5
    lblMin.Caption = lblMin.Caption - 1
  End If
  lblsec2.Caption = lblsec2.Caption - 1
  If lblsec2.Caption = -1 And lblSec1.Caption <> 0 Then   'Just for the newbs <> means does not equal
    lblSec1.Caption = lblSec1.Caption - 1     'Make the ten seconds go down
    lblsec2.Caption = 9
  End If
  If lblMin.Caption < 0 Then
    tmrTime.Enabled = False 'same as above probably don't need this one too
    lblMin.Caption = 0
    lblSec1.Caption = 0
    lblsec2.Caption = 0
    Gameover
  End If
End Sub
