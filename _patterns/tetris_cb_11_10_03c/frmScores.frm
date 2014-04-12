VERSION 5.00
Begin VB.Form frmScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patterns Hall of Fame"
   ClientHeight    =   3030
   ClientLeft      =   4680
   ClientTop       =   4590
   ClientWidth     =   3855
   Icon            =   "frmScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3855
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   2700
      TabIndex        =   13
      Top             =   2580
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   1680
      Max             =   3
      Min             =   1
      TabIndex        =   12
      Top             =   180
      Value           =   1
      Width           =   1335
   End
   Begin VB.TextBox txtPlace 
      Height          =   315
      Left            =   960
      TabIndex        =   11
      Text            =   "1"
      Top             =   180
      Width           =   615
   End
   Begin VB.TextBox txtPatterns 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox txtRows 
      Height          =   345
      Left            =   960
      TabIndex        =   8
      Top             =   1740
      Width           =   915
   End
   Begin VB.TextBox txtLevel 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   1380
      Width           =   375
   End
   Begin VB.TextBox txtScore 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   660
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Place:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Patterns:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Rows:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Level:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Score:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   555
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  HScroll1.Value = 1
  txtPlace.Text = HScroll1.Value
  txtName.Text = win(HScroll1.Value).Name
  txtScore.Text = win(HScroll1.Value).Score
  txtLevel.Text = win(HScroll1.Value).Level
  txtRows.Text = win(HScroll1.Value).Rows
  txtPatterns.Text = win(HScroll1.Value).Patterns

End Sub

Private Sub HScroll1_Change()
  txtPlace.Text = HScroll1.Value
  txtName.Text = win(HScroll1.Value).Name
  txtScore.Text = win(HScroll1.Value).Score
  txtLevel.Text = win(HScroll1.Value).Level
  txtRows.Text = win(HScroll1.Value).Rows
  txtPatterns.Text = win(HScroll1.Value).Patterns
End Sub

Private Sub HScroll1_Scroll()
  HScroll1_Change
End Sub
