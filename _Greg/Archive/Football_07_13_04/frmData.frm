VERSION 5.00
Begin VB.Form frmData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Field Data Manipulator"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.HScrollBar hsbFirstDownLine 
      Height          =   195
      LargeChange     =   10
      Left            =   840
      Max             =   60
      Min             =   -60
      TabIndex        =   13
      Top             =   540
      Width           =   1395
   End
   Begin VB.CheckBox chkFirstDownLine 
      Caption         =   "1st Line On"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   540
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkDefLine 
      Caption         =   "Def Line On"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   300
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkOffLine 
      Caption         =   "Off Line On"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      DrawWidth       =   2
      Height          =   600
      Left            =   840
      ScaleHeight     =   -60
      ScaleLeft       =   -60
      ScaleMode       =   0  'User
      ScaleTop        =   30
      ScaleWidth      =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1200
   End
   Begin VB.HScrollBar hsbDefLine 
      Height          =   195
      LargeChange     =   10
      Left            =   840
      Max             =   60
      Min             =   -60
      TabIndex        =   3
      Top             =   300
      Width           =   1395
   End
   Begin VB.HScrollBar hsbOffLine 
      Height          =   195
      LargeChange     =   10
      Left            =   840
      Max             =   60
      Min             =   -60
      TabIndex        =   1
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label lblFirstDownLine 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "1st Line:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   675
   End
   Begin VB.Label lblBallY 
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label lblBallX 
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Ball:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Def Line:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   300
      Width           =   675
   End
   Begin VB.Label lblDefLine 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   300
      Width           =   375
   End
   Begin VB.Label lblOffLine 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Off Line:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDefLine_Click()
  If chkDefLine.Value = vbChecked Then
    f.DefLineOn = True
  Else
    f.DefLineOn = False
  End If
End Sub

Private Sub chkFirstDownLine_Click()
  If chkFirstDownLine.Value = vbChecked Then
    f.FirstDownLineOn = True
  Else
    f.FirstDownLineOn = False
  End If
End Sub

Private Sub chkOffLine_Click()
  If chkOffLine.Value = vbChecked Then
    f.OffLineOn = True
  Else
    f.OffLineOn = False
  End If
End Sub

Private Sub Form_Load()
  
  'auto positions this form to right of frmField
  frmData.Left = frmField.Width
  frmData.Top = frmField.Top
  frmData.Height = frmField.Height
  frmData.Width = frmField.Width \ 2
  
  'loads data
  lblOffLine.Caption = f.OffLine
  lblDefLine.Caption = f.DefLine
  lblFirstDownLine.Caption = f.FirstDownLine
  picField.PSet (f.BallX, f.BallY)
  lblBallX.Caption = Format(f.BallX, "###.#")
  lblBallY.Caption = Format(f.BallY, "##.#")
End Sub

Private Sub hsbDefLine_Change()
  f.DefLine = hsbDefLine.Value
  lblDefLine.Caption = f.DefLine
End Sub

Private Sub hsbDefLine_Scroll()
  hsbDefLine_Change
End Sub

Private Sub hsbFirstDownLine_Change()
  f.FirstDownLine = hsbFirstDownLine.Value
  lblFirstDownLine.Caption = f.FirstDownLine
End Sub

Private Sub hsbFirstDownLine_Scroll()
  hsbFirstDownLine_Change
End Sub

Private Sub hsbOffLine_Change()
  f.OffLine = hsbOffLine.Value
  lblOffLine.Caption = f.OffLine
End Sub

Private Sub hsbOffLine_Scroll()
  hsbOffLine_Change
End Sub

Private Sub picField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    picField.Cls
    picField.PSet (X, Y)
    f.BallX = X
    f.BallY = Y
    lblBallX.Caption = Format(f.BallX, "###.#")
    lblBallY.Caption = Format(f.BallY, "##.#")
  End If
End Sub

Private Sub picField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    picField.Cls
    picField.PSet (X, Y)
    f.BallX = X
    f.BallY = Y
    lblBallX.Caption = Format(f.BallX, "###.#")
    lblBallY.Caption = Format(f.BallY, "##.#")
  End If
End Sub
