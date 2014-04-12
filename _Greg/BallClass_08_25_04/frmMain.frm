VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ball Class Demo"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6120
      Top             =   6120
   End
   Begin VB.PictureBox pic 
      Height          =   6000
      Left            =   0
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Shape shpBall 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   100
         Index           =   0
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   2820
         Width           =   100
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim ret As Boolean
  Dim i As Integer
  
  ret = SetPicColor(255, 255, 255)
  If ret = False Then MsgBox "Illegal parameters!"
  
  'create more shapes
  For i = 1 To MAX_BALLS - 1
    Load shpBall(i)
    shpBall(i).Visible = True
  Next i
  
  For i = 0 To MAX_BALLS - 1
    b(i).BallWidth = 0.5 + (MAX_BALLS - i) * 0.2
    b(i).BallLength = 0.5 + (MAX_BALLS - i) * 0.2
    shpBall(i).FillColor = RGB(0, 100 + i * 4, 0)
  Next i
  tmrUpdate_Timer
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  For i = 0 To MAX_BALLS - 1
    b(i).AssignTarget X, Y   'assigns target for ball
    b(i).Velocity = 0.5 + i * 0.2
  Next i
  tmrUpdate.Enabled = True

End Sub

Private Sub tmrUpdate_Timer()
  Dim i As Integer
  
  For i = 0 To MAX_BALLS - 1
    b(i).Update 'updates position of ball
    DoEvents
  Next i
  For i = 0 To MAX_BALLS - 1
    shpBall(i).Width = b(i).BallWidth
    shpBall(i).Height = b(i).BallLength
    shpBall(i).Left = b(i).X - shpBall(i).Width / 2  'draws ball
    shpBall(i).Top = b(i).Y + shpBall(i).Height / 2
    DoEvents
  Next i
End Sub


Private Function SetPicColor(r As Integer, g As Integer, b As Integer) As Boolean
  SetPicColor = False
  If r < 0 Or r > 255 Then Exit Function
  If g < 0 Or g > 255 Then Exit Function
  If b < 0 Or b > 255 Then Exit Function
  pic.BackColor = RGB(r, g, b)
  SetPicColor = True
End Function
