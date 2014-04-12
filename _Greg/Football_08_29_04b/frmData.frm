VERSION 5.00
Begin VB.Form frmData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Field Data Manipulator"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   9375
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   60
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Home Dir"
      Height          =   915
      Left            =   4920
      TabIndex        =   17
      Top             =   60
      Width           =   1335
      Begin VB.OptionButton optWest 
         Caption         =   "West"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton optEast 
         Caption         =   "East"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdHuddle 
      Caption         =   "Huddle"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   60
      Width           =   855
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   840
   End
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


Private Sub cmdHuddle_Click()
  Dim k As Single
  Dim bRet As Boolean
  Dim i As Integer
  
  If optEast.Value = True Then
    k = 1
  Else
    k = -1
  End If
  
  For i = 1 To MAX_BOTS
    bRet = b.SetVelocity(i, b.GetMaxVelocity(i))
  Next i
  
  'sample play - home team
  bRet = b.SetTargetX(1, f.OffLine - 9 * k)
  bRet = b.SetTargetY(1, f.BallY)
  bRet = b.SetTargetX(2, f.OffLine - 10 * k)
  bRet = b.SetTargetY(2, f.BallY + 1)
  bRet = b.SetTargetX(3, f.OffLine - 11 * k)
  bRet = b.SetTargetY(3, f.BallY + 2)
  bRet = b.SetTargetX(4, f.OffLine - 12 * k)
  bRet = b.SetTargetY(4, f.BallY + 3)
  bRet = b.SetTargetX(5, f.OffLine - 13 * k)
  bRet = b.SetTargetY(5, f.BallY + 2)
  bRet = b.SetTargetX(6, f.OffLine - 14 * k)
  bRet = b.SetTargetY(6, f.BallY + 1)
  bRet = b.SetTargetX(7, f.OffLine - 14 * k)
  bRet = b.SetTargetY(7, f.BallY + 1)
  bRet = b.SetTargetX(8, f.OffLine - 13 * k)
  bRet = b.SetTargetY(8, f.BallY - 2)
  bRet = b.SetTargetX(9, f.OffLine - 12 * k)
  bRet = b.SetTargetY(9, f.BallY - 3)
  bRet = b.SetTargetX(10, f.OffLine - 11 * k)
  bRet = b.SetTargetY(10, f.BallY - 2)
  bRet = b.SetTargetX(11, f.OffLine - 10 * k)
  bRet = b.SetTargetY(11, f.BallY - 2)
  
  'coaches
  bRet = b.SetTargetX(12, f.OffLine) 'home
  bRet = b.SetTargetY(12, 28.5)
  bRet = b.SetTargetX(24, f.OffLine) 'visitor
  bRet = b.SetTargetY(24, -27.5)
  
  'sample play - visitor team
  bRet = b.SetTargetX(13, f.OffLine + 1 * k)
  bRet = b.SetTargetY(13, f.BallY)
  bRet = b.SetTargetX(14, f.OffLine + 1 * k)
  bRet = b.SetTargetY(14, f.BallY + 1)
  bRet = b.SetTargetX(15, f.OffLine + 1 * k)
  bRet = b.SetTargetY(15, f.BallY + 2)
  bRet = b.SetTargetX(16, f.OffLine + 1 * k)
  bRet = b.SetTargetY(16, f.BallY + 3)
  bRet = b.SetTargetX(17, f.OffLine + 1 * k)
  bRet = b.SetTargetY(17, f.BallY + 4)
  bRet = b.SetTargetX(18, f.OffLine + 1 * k)
  bRet = b.SetTargetY(18, f.BallY - 1)
  bRet = b.SetTargetX(19, f.OffLine + 1 * k)
  bRet = b.SetTargetY(19, f.BallY - 2)
  bRet = b.SetTargetX(20, f.OffLine + 1 * k)
  bRet = b.SetTargetY(20, f.BallY - 3)
  bRet = b.SetTargetX(21, f.OffLine + 1 * k)
  bRet = b.SetTargetY(21, f.BallY - 4)
  bRet = b.SetTargetX(22, f.OffLine + 5 * k)
  bRet = b.SetTargetY(22, f.BallY + 3)
  bRet = b.SetTargetX(23, f.OffLine + 5 * k)
  bRet = b.SetTargetY(23, f.BallY - 4)
  
  tmrUpdate.Enabled = True
End Sub

Private Sub cmdStop_Click()
  tmrUpdate.Enabled = False
  GeneratePositions
End Sub

Private Sub Form_Load()
  
  'auto positions this form to right of frmField
  frmData.Left = 0 'frmField.Width
  frmData.Top = frmField.Height
  'frmData.Height = frmField.Height
  'frmData.Width = frmField.Width / 2.5
  
  picField.BackColor = RGB(0, 150, 0)
  
  'loads data
  lblOffLine.Caption = f.OffLine
  lblDefLine.Caption = f.DefLine
  lblFirstDownLine.Caption = f.FirstDownLine
  hsbOffLine.Value = f.OffLine
  hsbDefLine.Value = f.DefLine
  hsbFirstDownLine.Value = f.FirstDownLine
  picField.PSet (f.BallX, f.BallY)
  lblBallX.Caption = Format(f.BallX, "###.#")
  lblBallY.Caption = Format(f.BallY, "##.#")
  
  If optEast.Value = True Then
    optEast_Click
  Else
    optWest_Click
  End If
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

Private Sub optEast_Click()
  f.FirstDownLine = f.BallX + 10
  f.DefLine = f.BallX + 1
  f.OffLine = f.BallX
  If tmrUpdate.Enabled = True Then
    cmdHuddle_Click
  End If
End Sub

Private Sub optWest_Click()
  f.FirstDownLine = f.BallX - 10
  f.DefLine = f.BallX - 1
  f.OffLine = f.BallX
  If tmrUpdate.Enabled = True Then
    cmdHuddle_Click
  End If
End Sub

Private Sub picField_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  Dim nYTF As Single
  Dim bRet As Boolean
  
  If Button = 1 Then
    picField.Cls
    picField.PSet (x, y)
    f.BallX = x
    If y > 4 Then f.BallY = 4
    If y < -4 Then f.BallY = -4
        
    'position lines based upon ball and direction of home team
    nYTF = 10
    If optEast.Value = True Then  'facing EAST
      f.OffLine = f.BallX - 0.15
      f.DefLine = f.BallX + 0.15
      If f.BallX > 40 Then nYTF = 50 - f.BallX
      f.FirstDownLine = f.BallX + nYTF
    Else    'facing WEST
      f.OffLine = f.BallX + 0.15
      f.DefLine = f.BallX - 0.15
      If f.BallX < -40 Then nYTF = f.BallX + 50
      f.FirstDownLine = f.BallX - nYTF
    End If
    
    'f.QBX = x
    'f.QBY = Y
    lblBallX.Caption = Format(f.BallX, "###.#")
    lblBallY.Caption = Format(f.BallY, "##.#")
    tmrUpdate.Enabled = False
  End If
  
  If Button = 2 Then
     For i = 1 To MAX_BOTS
       bRet = b.SetTargetX(i, x)
       bRet = b.SetTargetY(i, y)
     Next i
     bRet = b.SetTargetX(12, f.OffLine)
     bRet = b.SetTargetY(12, 26)
     bRet = b.SetTargetX(24, f.OffLine)
     bRet = b.SetTargetY(24, -26)
     
    'f.RecX = x    'define position of receiver
    'f.RecY = Y
    'f.MoveBall    'command f object to move ball
    tmrUpdate.Enabled = True  'start updating position of the ball
  End If
  
End Sub

Private Sub picField_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  Dim bRet As Boolean
  
  If Button = 1 Then
    picField.Cls
    picField.PSet (x, y)
    f.BallX = x
    f.BallY = y
    lblBallX.Caption = Format(f.BallX, "###.#")
    lblBallY.Caption = Format(f.BallY, "##.#")
  End If

  For i = 1 To MAX_BOTS
    bRet = b.SetTargetX(i, x)
    bRet = b.SetTargetY(i, y)
  Next i

End Sub

Private Sub tmrUpdate_Timer()
  f.Update
  b.UpdateBots
End Sub
