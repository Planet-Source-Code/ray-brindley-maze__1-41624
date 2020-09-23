VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   7860
   ClientLeft      =   2505
   ClientTop       =   2025
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   10500
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   7080
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6720
      TabIndex        =   3
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label lblAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label lblScores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label lblPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblAbout_Click()
  frmMenu.Hide
  frmAbout.Show
End Sub

Private Sub lblExit_Click()
  End
End Sub

Private Sub lblInfo_Click()
  frmMenu.Hide
  frmInfo.Show
End Sub

Private Sub lblPlay_Click()
  frmMenu.Hide
  frmChar.Show
End Sub

Private Sub lblScores_Click()
  frmMenu.Hide
  frmHighScores.Show
End Sub
