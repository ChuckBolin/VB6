VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6420
      Max             =   11
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Phase"
      Height          =   375
      Left            =   6420
      TabIndex        =   6
      Top             =   2460
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   5820
      Max             =   60
      Min             =   1
      TabIndex        =   4
      Top             =   900
      Value           =   1
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   7140
      TabIndex        =   2
      Top             =   1500
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   1500
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Top             =   1020
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  g.StartGameClock
End Sub

Private Sub Command2_Click()
  g.StopGameClock
End Sub

Private Sub Command3_Click()
  
  g.SetPhase HScroll1.Value
End Sub

Private Sub Timer1_Timer()
  g.UpdateGameClock
  Label1.Caption = g.GameClock
  Label2.Caption = g.PhaseText
End Sub

Private Sub VScroll1_Change()
  g.SetTimeFactor VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
  VScroll1_Change
End Sub
