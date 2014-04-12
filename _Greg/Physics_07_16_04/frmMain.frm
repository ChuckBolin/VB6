VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Physics Demo"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Target"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtData 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6960
      Top             =   3360
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.HScrollBar hsbVel 
      Height          =   255
      LargeChange     =   100
      Left            =   6720
      Max             =   200
      TabIndex        =   8
      Top             =   840
      Value           =   100
      Width           =   1335
   End
   Begin VB.HScrollBar hsbElev 
      Height          =   255
      LargeChange     =   15
      Left            =   6720
      Max             =   90
      TabIndex        =   5
      Top             =   480
      Value           =   30
      Width           =   1335
   End
   Begin VB.HScrollBar hsbDir 
      Height          =   255
      LargeChange     =   15
      Left            =   6720
      Max             =   360
      TabIndex        =   2
      Top             =   120
      Value           =   90
      Width           =   1335
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   3000
      Left            =   0
      ScaleHeight     =   -60
      ScaleLeft       =   -60
      ScaleMode       =   0  'User
      ScaleTop        =   30
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Line Line1 
         X1              =   -60
         X2              =   61.212
         Y1              =   -1.837
         Y2              =   -1.837
      End
      Begin VB.Shape shpZ 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   960
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape shpBall 
         FillStyle       =   0  'Solid
         Height          =   49
         Left            =   960
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   50
      End
      Begin VB.Shape shpTarget 
         Height          =   98
         Left            =   3840
         Top             =   2280
         Width           =   99
      End
   End
   Begin VB.Label lblVel 
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Vel:"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblElev 
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Elev:"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblDir 
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Dir:"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'creates private variables
Private BallX As Single  'current ball position
Private BallY As Single
Private BallZ As Single
Private nVT As Single 'total velocity
Private nVV As Single 'vertical velocity
Private nVH As Single 'horizontal velocity
Private nVX As Single 'velocity horizontal X
Private nVY As Single 'velocity horizontal Y
Private nElev As Single 'in radians
Private nDir As Single 'direction in radians
Private nG As Single 'gravitational value 32.2Ft/sec*sec
Private nTX As Single 'target X, Y, Z
Private nTY As Single
Private nTZ As Single
Private nStartX As Single 'beginning location of ball
Private nStartY As Single
Private nDistance As Single 'distance traveled

Private Sub cmdStart_Click()
  If cmdStart.Caption = "Start" Then
    nElev = CtoR(DtoC(hsbElev.Value))
    nDir = CtoR(DtoC(hsbDir.Value))
    nVT = hsbVel.Value * 0.01
    
    Dim nFactor As Integer
    nFactor = ((nVT / 0.25) ^ 2) * 6.25
    
    'distance algorithm in terms of elevation and velocity
    '((nVT / 0.25) ^ 2) * 6.25
    '(Sin(2 * nElev) * nFactor)
    
    txtData.Text = txtData.Text & "Calc Distance: " & Format((Sin(2 * nElev) * nFactor), "###.##") & vbCrLf
    txtData.Text = txtData.Text & "Elev: " & RtoD(nElev) & vbCrLf
    txtData.Text = txtData.Text & "Dir: " & RtoD(nDir) & vbCrLf
    txtData.Text = txtData.Text & "Vel: " & nVT & vbCrLf
    
    cmdStart.Caption = "Stop"
    tmrUpdate.Enabled = True
    nStartX = BallX
    nStartY = BallY
  Else
    cmdStart.Caption = "Start"
    tmrUpdate.Enabled = False
    BallX = -50
    BallY = 0
    BallZ = 0
    nG = 10.73 'Yards per sec per sec
    shpBall.Left = BallX
    shpBall.Top = BallY
    shpZ.Left = BallX
    shpZ.Top = BallZ
    hsbDir_Change
    hsbElev_Change
    hsbVel_Change
  End If
End Sub

Private Sub cmdTarget_Click()
  Dim d As Single  'distance
  Dim v As Single  'velocity
  Dim e As Single 'elevation in radians
  
  d = Sqr((nTX - BallX) ^ 2 + (nTY - BallY) ^ 2)
  e = 0.785
  v = (Sqr(d / (Sin(2 * e) * 6.25))) / 4
  
  nElev = e
  nDir = GetTargetDirection(BallX, BallY, nTX, nTY)
  nVT = v
  Dim nFactor As Integer
  nFactor = ((nVT / 0.25) ^ 2) * 6.25
    
  'distance algorithm in terms of elevation and velocity
  '((nVT / 0.25) ^ 2) * 6.25
  '(Sin(2 * nElev) * nFactor)
    
  txtData.Text = txtData.Text & "Calc Distance: " & Format((Sin(2 * nElev) * nFactor), "###.##") & vbCrLf
  txtData.Text = txtData.Text & "Elev: " & RtoD(nElev) & vbCrLf
  txtData.Text = txtData.Text & "Dir: " & RtoD(nDir) & vbCrLf
  txtData.Text = txtData.Text & "Vel: " & nVT & vbCrLf
    
  tmrUpdate.Enabled = True
  nStartX = BallX
  nStartY = BallY
End Sub

Private Sub Command1_Click()
txtData.Text = ""
End Sub

Private Sub Form_Load()
  
  'initialization
  BallX = -50
  BallY = 0
  BallZ = 0
  nG = 10.73 'Yards per sec per sec
  shpBall.Left = BallX
  shpBall.Top = BallY
  shpZ.Left = BallX
  shpZ.Top = BallZ
  hsbDir_Change
  hsbElev_Change
  hsbVel_Change
  nTX = shpTarget.Left
  nTY = shpTarget.Top
End Sub

Private Sub hsbDir_Change()
  lblDir.Caption = hsbDir.Value
End Sub

Private Sub hsbDir_Scroll()
  hsbDir_Change
End Sub

Private Sub hsbElev_Change()
  lblElev.Caption = hsbElev.Value
End Sub

Private Sub hsbElev_Scroll()
  hsbElev_Change
End Sub

Private Sub hsbVel_Change()
  lblVel.Caption = Format(hsbVel.Value * 0.01, "##.##0")
End Sub

Private Sub hsbVel_Scroll()
  hsbVel_Change
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    BallX = x
    BallY = y
    shpBall.Left = BallX
    shpBall.Top = BallY
  End If
  
  If Button = 2 Then
    shpTarget.Left = x
    shpTarget.Top = y
    nTX = shpTarget.Left
    nTY = shpTarget.Top
  End If
End Sub


Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  frmMain.Caption = "X: " & x & "  Y: " & y
End Sub

'every time interval this routine updates the trajectory and position of the ball
'to create the illusion of movement..updates
Private Sub tmrUpdate_Timer()
    
  'get current velocity total and elevation
  nVH = Cos(nElev) * nVT: nVV = (Sin(nElev) * nVT) - 0.01
  nVX = nVH * Sin(nDir): nVY = nVH * Cos(nDir)
  BallX = BallX + nVX: BallY = BallY + nVY: BallZ = BallZ + nVV
  nVT = Sqr((nVH ^ 2) + (nVV ^ 2))
  If nVH = 0 Then nVH = 0.01
  nElev = Atn(nVV / nVH)
  
  'comment out GoTo if you want to see these values
  GoTo Skip
  'txtData.Text = ""
  txtData.Text = txtData.Text & "VH: " & nVH & vbCrLf
  txtData.Text = txtData.Text & "VV: " & nVV & vbCrLf
  txtData.Text = txtData.Text & "VT: " & nVT & vbCrLf
  txtData.Text = txtData.Text & "VX: " & nVX & vbCrLf
  txtData.Text = txtData.Text & "VY: " & nVY & vbCrLf
  txtData.Text = txtData.Text & "Elev: " & nElev & vbCrLf
  txtData.Text = txtData.Text & "Dir: " & nDir & vbCrLf
  txtData.Text = txtData.Text & "BallX: " & BallX & vbCrLf
  txtData.Text = txtData.Text & "BallY: " & BallY & vbCrLf
  txtData.Text = txtData.Text & "BallZ: " & BallZ & vbCrLf
Skip:
  
  shpBall.Left = BallX
  shpBall.Top = BallY
  shpZ.Left = BallX
  shpZ.Top = BallZ
  
  'ball has landed on ground
  If BallZ <= 0 Or BallX <= -60 Or BallX >= 60 Or BallY > 30 Or BallY < -30 Then
    tmrUpdate.Enabled = False
    cmdStart.Enabled = True
    BallZ = 0
    nDistance = Sqr((BallX - nStartX) ^ 2 + (BallY - nStartY) ^ 2)
    txtData.Text = txtData.Text & Format(nDistance, "###.##") & vbCrLf
  End If
End Sub
