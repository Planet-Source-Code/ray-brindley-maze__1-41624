VERSION 5.00
Begin VB.Form frmGameOver 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Game Over"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   ControlBox      =   0   'False
   Icon            =   "frmGameOver.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmGameOver.frx":08CA
   ScaleHeight     =   6615
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrGameover 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   6120
   End
   Begin VB.Image imgGoodbye 
      Height          =   6690
      Left            =   -4560
      Picture         =   "frmGameOver.frx":44CB
      Top             =   5640
      Visible         =   0   'False
      Width           =   5910
   End
   Begin VB.Image imgGameover 
      Height          =   6690
      Left            =   4800
      Picture         =   "frmGameOver.frx":821F
      Top             =   5640
      Visible         =   0   'False
      Width           =   5910
   End
   Begin VB.Image imgCongrat 
      Height          =   6690
      Left            =   4320
      Picture         =   "frmGameOver.frx":BE20
      Top             =   -5520
      Visible         =   0   'False
      Width           =   5910
   End
End
Attribute VB_Name = "frmGameOver"
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

Private Sub tmrGameover_Timer()
  If frmGameOver.Picture = imgGameover.Picture Then
    Me.Visible = False
    frmMenu.Show
    tmrGameover.Enabled = False
  ElseIf frmGameOver.Picture = imgCongrat.Picture Then
    Response = MsgBox("Do you want to play again?", vbYesNo, "Play Again?")
    If Response = vbYes Then
      frmMenu.Show
      Me.Visible = False
    Else
      Me.Hide
      frmMenu.Show
      Quit
    End If
    tmrGameover.Enabled = False
  ElseIf frmGameOver.Picture = imgGoodbye.Picture Then
    End
  End If
End Sub
