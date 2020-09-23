VERSION 5.00
Begin VB.Form frmChar 
   BorderStyle     =   0  'None
   Caption         =   "Character Select"
   ClientHeight    =   4785
   ClientLeft      =   4710
   ClientTop       =   3315
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   6000
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optRabbit 
      Height          =   1335
      Left            =   3600
      Picture         =   "frmChar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton optEvil 
      Height          =   1335
      Left            =   720
      Picture         =   "frmChar.frx":5ED2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton optBoy 
      Height          =   1335
      Left            =   3600
      Picture         =   "frmChar.frx":BB28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton optGirl 
      Height          =   1335
      Left            =   720
      Picture         =   "frmChar.frx":1139A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  optBoy.Value = False
  optGirl.Value = False
  optEvil.Value = False
  optRabbit.Value = False
End Sub

Private Sub optBoy_Click()
  frmMaze.imgChar.Picture = frmMaze.imgBoy.Picture
  frmMaze.Show
  Unload Me
End Sub

Private Sub optEvil_Click()
  frmMaze.imgChar.Picture = frmMaze.imgEvil.Picture
  frmMaze.Show
  Unload Me
End Sub

Private Sub optGirl_Click()
  frmMaze.imgChar.Picture = frmMaze.imgGirl.Picture
  frmMaze.Show
  Unload Me
End Sub

Private Sub optRabbit_Click()
  frmMaze.imgChar.Picture = frmMaze.imgGirl.Picture
  frmMaze.Show
  Unload Me
End Sub
