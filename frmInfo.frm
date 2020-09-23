VERSION 5.00
Begin VB.Form frmInfo 
   Caption         =   "Information"
   ClientHeight    =   3780
   ClientLeft      =   4950
   ClientTop       =   3945
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5235
   Begin VB.Label lblReturn 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmInfo.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblReturn_Click()
  frmInfo.Hide
  frmMaze.Show
End Sub
