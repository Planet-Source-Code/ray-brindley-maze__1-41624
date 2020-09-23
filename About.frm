VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   7500
   ClientLeft      =   1785
   ClientTop       =   2805
   ClientWidth     =   6750
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   3735
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Image imgAbout 
      Height          =   7500
      Left            =   0
      Picture         =   "About.frx":08CA
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmAbout"
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
'i did this form this way because it is a simpler way to read and check
'what you have written in a file rather than in that small caption box
'it is also easy to understand for begginers on how to read a tile
'my other files that are read are quite detailed and have extra things
'that aren't used as often. this is the complete basics of reading a file

Private Sub Form_Load()
  AboutFile = FreeFile  'you must declare the something (in this case AboutFile) to be a freefile

  Open "at.at" For Input As #AboutFile    'Open the required file
    Input #AboutFile, Cap                 'Read the file and put in the Text (represented as Cap)
      Label1.Caption = Cap                'Put it into the label caption.
  Close #AboutFile                        'don't forget to close it otherwise it will have problems
End Sub
