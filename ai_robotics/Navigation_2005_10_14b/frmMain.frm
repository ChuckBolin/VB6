VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Robot Navigation v0.12 October 14, 2005 - By Chuck Bolin"
   ClientHeight    =   10365
   ClientLeft      =   525
   ClientTop       =   480
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   3600
      TabIndex        =   62
      Top             =   8760
      Width           =   975
   End
   Begin VB.Timer tmrLR 
      Interval        =   250
      Left            =   6120
      Top             =   8880
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Left            =   1680
      TabIndex        =   59
      Text            =   "Time Factor"
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   8760
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Laser Rangefinder"
      Height          =   8895
      Left            =   10440
      TabIndex        =   55
      Top             =   0
      Width           =   2295
      Begin VB.ListBox lstLR 
         Height          =   7080
         ItemData        =   "frmMain.frx":0000
         Left            =   480
         List            =   "frmMain.frx":0002
         TabIndex        =   60
         Top             =   1560
         Width           =   1335
      End
      Begin VB.PictureBox picLR 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   1000
         Left            =   120
         ScaleHeight     =   -1000
         ScaleMode       =   0  'User
         ScaleTop        =   1000
         ScaleWidth      =   2000
         TabIndex        =   56
         Top             =   240
         Width           =   2000
      End
      Begin VB.Line Line1 
         BorderWidth     =   4
         X1              =   240
         X2              =   2040
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label17 
         Caption         =   "Range Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "GPS Position Information"
      Height          =   1815
      Left            =   8160
      TabIndex        =   42
      Top             =   5160
      Width           =   2175
      Begin VB.CheckBox chkGPS 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   660
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Dir:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "Vel:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblGPSX 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblGPSY 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGPSVel 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblGPSDir 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Autonomous Mode Source"
      Height          =   1815
      Left            =   8160
      TabIndex        =   39
      Top             =   7080
      Width           =   2175
      Begin VB.OptionButton optAutoProg 
         Caption         =   "Programmed"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton optAutoGPS 
         Caption         =   "GPS"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optAutoDR 
         Caption         =   "Dead Reckoning"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optAutoActual 
         Caption         =   "Actual Bot Info"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Waypoint Data"
      Height          =   615
      Left            =   2640
      TabIndex        =   28
      Top             =   0
      Width           =   5415
      Begin VB.Label lblWPNum 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Waypoint:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblWPDir 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblWPDist 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Direction:"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Distance:"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frame3 
      Caption         =   "Drive Mode:"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optAuto 
         Caption         =   "Autonomous"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dead Reckoning"
      Height          =   1815
      Left            =   8160
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
      Begin VB.CheckBox chkDR 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblDRDir 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblDRVel 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblDRY 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblDRX 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Vel:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Dir:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   660
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actual Bot Info"
      Height          =   3135
      Left            =   8160
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      Begin VB.CheckBox chkBot 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   1800
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblBotDist 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Dist:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblBotTurn 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Turn:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblBotFuel 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   660
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Fuel:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   375
      End
      Begin VB.Shape shpBotFuel 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   560
         Top             =   1800
         Width           =   1000
      End
      Begin VB.Label lblBotDir 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblBotVel 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblBotY 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblBotX 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Dir:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Vel:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   7320
      Top             =   8880
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   8000
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   -8000
      ScaleMode       =   0  'User
      ScaleTop        =   8000
      ScaleWidth      =   8000
      TabIndex        =   0
      Top             =   720
      Width           =   8000
      Begin VB.Label lblMousePos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Shape shpBot 
         BackStyle       =   1  'Opaque
         Height          =   127
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   3240
         Visible         =   0   'False
         Width           =   127
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Time:"
      Height          =   255
      Left            =   1200
      TabIndex        =   58
      Top             =   8760
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bAutoMode As Boolean

Private Sub cboTime_Click()
  tmrUpdate.Enabled = False
  'tmrUpdate.Interval = 100
  
  Select Case cboTime.ListIndex
    Case 0:
        tmrUpdate.Interval = 400 'x 1/4
    Case 1:
        tmrUpdate.Interval = 200 'x 1/2
    Case 2:
        tmrUpdate.Interval = 100 'x 1
    Case 3:
        tmrUpdate.Interval = 50 'x 2
    Case 4:
        tmrUpdate.Interval = 25 'x 4
    Case 5:
        tmrUpdate.Interval = 12 'x 8
  End Select
  tmrUpdate.Enabled = True
End Sub

Private Sub cmdPause_Click()
  If tmrUpdate.Enabled = True Then
    tmrUpdate.Enabled = False
    tmrLR.Enabled = False
  Else
    tmrUpdate.Enabled = True
    tmrLR.Enabled = True
  End If
End Sub

Private Sub cmdReset_Click()
  bot.Energy = 100000
  dr.Turn = 0
  dr.Velocity = 0
  g_nLegNum = 1
  bot.X = leg(g_nLegNum).X1
  bot.Y = leg(g_nLegNum).Y1
  bot.Turn = 0: bot.Velocity = 0
  dr.X = bot.X
  dr.Y = bot.Y
  bot.Direction = PI / 2
  dr.Direction = bot.Direction
  u_GPS.Direction = bot.Direction
  g_nOdometer = 0
  g_nLastLegNum = 8
  optManual.Value = True
  
End Sub

Private Sub Form_Load()
  LoadVariables
  m_bAutoMode = False

  cboTime.AddItem "x 1/4"
  cboTime.AddItem "x 1/2"
  cboTime.AddItem "x 1"
  cboTime.AddItem "x 2"
  cboTime.AddItem "x 4"
  cboTime.AddItem "x 8"
  cboTime.ListIndex = 2
End Sub

Private Sub optAuto_Click()
  m_bAutoMode = True    'operate in autonomous mode
  pic.SetFocus
End Sub

Private Sub optManual_Click()
  m_bAutoMode = False   'disable autonomous mode
  pic.SetFocus
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft And m_bAutoMode = False Then
    bot.Turn = bot.Turn + 0.001
  ElseIf KeyCode = vbKeyRight And m_bAutoMode = False Then
    bot.Turn = bot.Turn - 0.001
  ElseIf KeyCode = vbKeyUp And m_bAutoMode = False Then
    bot.Velocity = bot.Velocity + 1
    If bot.Velocity > bot.MaxVel Then bot.Velocity = bot.MaxVel
  ElseIf KeyCode = vbKeyDown And m_bAutoMode = False Then
    bot.Velocity = bot.Velocity - 1
    If bot.Velocity < bot.MinVel Then bot.Velocity = bot.MinVel
  ElseIf KeyCode = 27 Then 'escape key
    Unload Me
  ElseIf KeyCode = 82 Then 'R key, reset
    'bot.X = 10000: bot.Y = 10000:  bot.Direction = 1.57
    bot.Energy = 100000
    dr.Turn = 0
    dr.Velocity = 0
    g_nLegNum = 1
    bot.X = leg(g_nLegNum).X1
    bot.Y = leg(g_nLegNum).Y1
    bot.Turn = 0: bot.Velocity = 0
    dr.X = bot.X
    dr.Y = bot.Y
    dr.Direction = bot.Direction
    g_nOdometer = 0
    g_nLastLegNum = 5
  ElseIf KeyCode = 32 Then 'space - toggle timer on/off
    If tmrUpdate.Enabled = True Then
      tmrUpdate.Enabled = False
      
    Else
      tmrUpdate.Enabled = True
    End If
  'ElseIf KeyCode = 65 Then  'A
  Else
    'frmMain.Caption = KeyCode
  End If
  
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'frmMain.Caption = X & " " & Y
  lblMousePos = "X: " & FormatNumber(bot.X - (4000 - X), 0) & "  Y: " & FormatNumber(bot.Y - (4000 - Y), 0)
  lblMousePos.Left = X + 500
  lblMousePos.Top = Y - 500
  
End Sub

'this updates range data from Laser Rangerfinder
Private Sub tmrLR_Timer()
  Dim i As Integer
  
  lstLR.Clear
  
  For i = 1 To 36 ' To 1 Step -1
    lstLR.AddItem 90 - ((i - 1) * 5) & "(" & 37 - i & "): " & vbTab & CInt(g_uLR(i).Range)
  Next i
End Sub

Private Sub tmrUpdate_Timer()
  Dim DX, DY As Single 'difference calculations
  Dim DX2, DY2 As Single
  Dim cx, cy As Single
  Dim i As Integer
  
  'calls sub to generate 'turn' and 'velocity' variables
  If m_bAutoMode = True Then Autonomous
 
  'redraw everything
  pic.Cls
  
  'UPDATE actual bot information...everything based upon this
  bot.Energy = bot.Energy - bot.Velocity
  
  'calc actual bot data..verify there is enough fuel to continue
  If bot.Energy > 0 Then
    bot.Direction = bot.Direction + bot.Turn
    If bot.Direction > 6.28 Then bot.Direction = 0
    If bot.Direction < 0 Then bot.Direction = 6.28
    bot.VX = bot.Velocity * Cos(bot.Direction)
    bot.VY = bot.Velocity * Sin(bot.Direction)
    bot.X = bot.X + bot.VX
    bot.Y = bot.Y + bot.VY
    g_nOdometer = g_nOdometer + bot.Velocity
  Else
    bot.Velocity = 0
    bot.VX = 0
    bot.VY = 0
    bot.Energy = 0
  End If
  
  'UPDATE DR information
  dr.Velocity = bot.Velocity / NAV_DR_VEL_FACTOR  'slight error
  dr.Direction = bot.Direction / NAV_DR_DIR_FACTOR 'slight error
  If dr.Direction > (2 * PI) Then dr.Direction = dr.Direction - (2 * PI)
  If dr.Direction < 0 Then dr.Direction = dr.Direction + (2 * PI)
  dr.VX = Cos(dr.Direction) * dr.Velocity
  dr.VY = Sin(dr.Direction) * dr.Velocity
  dr.X = dr.X + dr.VX
  dr.Y = dr.Y + dr.VY
    
  UpdateLaserRangeFinder
  UpdateRadioBeacons
  UpdateGPS 'draws GPS symbol
  DrawRoute
  DrawObstacles
  DrawBotSymbol
  DrawDRSymbol
  
  'update displays
  'actual information
  lblBotX = FormatNumber(bot.X, 1)
  lblBotY = FormatNumber(bot.Y, 1)
  lblBotVel = FormatNumber(bot.Velocity, 1)
  lblBotDir = FormatNumber(RtoC(bot.Direction), 1)
  lblBotFuel = bot.Energy
  If bot.Energy > 50000 Then
    shpBotFuel.BackColor = vbGreen
  ElseIf bot.Energy > 10000 Then
    shpBotFuel.BackColor = vbYellow
  Else
    shpBotFuel.BackColor = vbRed
  End If
  shpBotFuel.Width = (bot.Energy / 100000) * 1000
  lblBotTurn = FormatNumber(bot.Turn, 3)
  lblBotDist = g_nOdometer
  
  'MsgBox UBound(leg) 'leg(g_nLegNum).X2 & "  " & leg(g_nLegNum).Y2
  
  'update waypoint information
  lblWPNum = g_nLegNum
  lblWPDist = FormatNumber(GetTargetDistance2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2), 0)
  lblWPDir = FormatNumber(RtoC(GetTargetDirection2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)), 0)
  
  
  'dead-reckoning information
  lblDRX = FormatNumber(dr.X, 1)
  lblDRY = FormatNumber(dr.Y, 1)
  lblDRVel = FormatNumber(dr.Velocity, 1)
  lblDRDir = FormatNumber(RtoC(dr.Direction), 1)

End Sub

'*******************************************************************
' DrawBotSymbol
'*******************************************************************
Public Sub DrawBotSymbol()
  
  'draws bot
  If chkBot.Value = 1 Then
    pic.FillStyle = 0
    pic.ForeColor = vbBlack
    pic.FillColor = vbRed
    pic.Circle (4000, 4000), 64, 0  'draws bot
    If bot.Velocity > 0 Then
      pic.ForeColor = vbBlack
    ElseIf bot.Velocity < 0 Then
      pic.ForeColor = vbRed
    End If
    pic.Line (4000, 4000)-(4000 + (25 * bot.VX), 4000 + (25 * bot.VY)) 'direction line
  Else
  
  End If
End Sub

'*******************************************************************
' DrawDRSymbol
'*******************************************************************
Public Sub DrawDRSymbol()
  Dim DX, DY As Single
  
  'draws Dead Reckoning on screen
  If chkDR.Value = 1 Then
    pic.DrawStyle = 0
    pic.FillStyle = 0
    pic.FillColor = vbMagenta
    DX = bot.X - dr.X
    DY = bot.Y - dr.Y

    pic.Circle (4000 - DX, 4000 - DY), 96, 0
    pic.ForeColor = vbBlack
    If dr.Velocity > 0 Then
        pic.ForeColor = vbBlack
    ElseIf dr.Velocity < 0 Then
        pic.ForeColor = vbRed
    End If
    pic.Line (4000 - DX, 4000 - DY)-(4000 - DX + (25 * dr.VX), 4000 - DY + (25 * dr.VY))
  End If
  
End Sub

'*******************************************************************
' DrawObstacles
'*******************************************************************
Public Sub DrawObstacles()
  Dim i As Integer
  Dim DX, DY As Single
  
  If g_nMaxObstacles = 0 Then Exit Sub
  
  'draw obstacles
  pic.DrawStyle = 0
  pic.FillStyle = 0
  
   'move obstacles
   'For i = 0 To 25
   '  g_uOb(i).X = g_uOb(i).X + 5
   '  g_uOb(i).Y = g_uOb(i).Y + 5
   'Next i
   'For i = 26 To 50
   '  g_uOb(i).X = g_uOb(i).X - 5
   '  g_uOb(i).Y = g_uOb(i).Y + 5
   'Next i
   'For i = 51 To 75
   '  g_uOb(i).X = g_uOb(i).X + 5
   '  g_uOb(i).Y = g_uOb(i).Y - 5
   'Next i
   'For i = 76 To 100
   '  g_uOb(i).X = g_uOb(i).X - 5
   '  g_uOb(i).Y = g_uOb(i).Y - 5
   'Next i
   'For i = 0 To 100
   '  If g_uOb(i).X < 8000 Or g_uOb(i).X > 27000 Or g_uOb(i).Y < 8000 Or g_uOb(i).Y > 18000 Then
   '    g_uOb(i).X = 8000 + GetRandomSingle(0, 20000)
   '    g_uOb(i).Y = 8000 + GetRandomSingle(0, 10000)
   '  End If

   '
   'Next i
  
  For i = 1 To g_nMaxObstacles
    pic.FillColor = g_uOb(i).Color
    DX = bot.X - g_uOb(i).X
    DY = bot.Y - g_uOb(i).Y
    pic.Circle (4000 - DX, 4000 - DY), g_uOb(i).Radius, 0
  Next i

End Sub

'*******************************************************************
' DrawRoute
'*******************************************************************
Public Sub DrawRoute()
  Dim i  As Integer
  
  If g_nLastLegNum = 0 Then Exit Sub
  'draws waypoints
  pic.ForeColor = vbBlack
  pic.DrawWidth = 1
  pic.FillStyle = 1
  
  For i = 1 To g_nLastLegNum
    pic.ForeColor = vbCyan
    pic.DrawStyle = 2
    pic.Line (4000 + leg(i).X1 - bot.X, 4000 + leg(i).Y1 - bot.Y)-(4000 + leg(i).X2 - bot.X, 4000 + leg(i).Y2 - bot.Y)
    pic.ForeColor = vbBlack
    pic.DrawStyle = 0
    If leg(i).Orientation = 1 Then 'North
      pic.Line (4000 + leg(i).X2 - bot.X - leg(i).Width, 4000 + leg(i).Y2 - bot.Y + leg(i).Width)-(4000 + leg(i).X1 - bot.X + leg(i).Width, 4000 + leg(i).Y1 - bot.Y - leg(i).Width), , B
    ElseIf leg(i).Orientation = 2 Then 'East
      pic.Line (4000 + leg(i).X1 - bot.X - leg(i).Width, 4000 + leg(i).Y1 - bot.Y + leg(i).Width)-(4000 + leg(i).X2 - bot.X + leg(i).Width, 4000 + leg(i).Y2 - bot.Y - leg(i).Width), , B
    ElseIf leg(i).Orientation = 3 Then 'South
      pic.Line (4000 + leg(i).X1 - bot.X - leg(i).Width, 4000 + leg(i).Y1 - bot.Y + leg(i).Width)-(4000 + leg(i).X2 - bot.X + leg(i).Width, 4000 + leg(i).Y2 - bot.Y - leg(i).Width), , B
    ElseIf leg(i).Orientation = 4 Then 'West
      pic.Line (4000 + leg(i).X2 - bot.X - leg(i).Width, 4000 + leg(i).Y2 - bot.Y + leg(i).Width)-(4000 + leg(i).X1 - bot.X + leg(i).Width, 4000 + leg(i).Y1 - bot.Y - leg(i).Width), , B
    End If
  Next i

End Sub

'*******************************************************************
' UpdateGPS
'*******************************************************************
Public Sub UpdateGPS()
  Dim DX, DY As Single 'difference calculations
  Dim DX2, DY2 As Single
  Dim nSatOffset As Single
  Static nGPSOffsetX As Single, nGPSOffsetY As Single
  Dim i As Integer
 
  If g_nMaxGPS = 0 Then Exit Sub
  
  'gps information
  g_nNumGPSSat = 2 'default value...two sats
  g_bGPSStatus = True
  For i = 1 To g_nMaxGPS
    If bot.X > GPS(i).A.X And bot.X < GPS(i).B.X And bot.Y < GPS(i).A.Y And bot.Y > GPS(i).B.Y Then
      g_nNumGPSSat = GPS(i).Num 'in this GPS coverage area
      If chkGPS.Value = 1 Then  'draws box
        DX = bot.X - GPS(i).A.X
        DY = bot.Y - GPS(i).A.Y
        DX2 = bot.X - GPS(i).B.X
        DY2 = bot.Y - GPS(i).B.Y
        If g_nNumGPSSat < 1 Then
          pic.FillColor = vbRed
          pic.ForeColor = vbRed
        Else
          pic.FillColor = vbGreen
          pic.ForeColor = vbGreen
        End If
        pic.FillStyle = 4
        pic.Line (4000 - DX, 4000 - DY)-(4000 - DX2, 4000 - DY2), , B
      End If
      Exit For
    End If
  Next i
  If g_nNumGPSSat < 1 Then
    g_bGPSStatus = False
    lblGPSX = "- - - -"
    lblGPSY = "- - - -"
    lblGPSVel = "- - - -"
    lblGPSDir = "- - - -"
  Else
    nSatOffset = 110 - (g_nNumGPSSat * 20) 'more sats...greater confidence in position
    g_nGPSOffsetX = GetRandomSingle(-nSatOffset, nSatOffset)
    g_nGPSOffsetY = GetRandomSingle(-nSatOffset, nSatOffset)
    nGPSOffsetX = ((nGPSOffsetX * 10 - (nGPSOffsetX / 10)) + g_nGPSOffsetX) / 10 ' / 100
    nGPSOffsetY = ((nGPSOffsetY * 10 - (nGPSOffsetY / 10)) + g_nGPSOffsetY) / 10 ' / 100
    u_GPS.X = bot.X + nGPSOffsetX
    u_GPS.Y = bot.Y + nGPSOffsetY
    u_GPS.Velocity = bot.Velocity
    u_GPS.Direction = bot.Direction
    lblGPSX = u_GPS.X
    lblGPSY = u_GPS.Y
    lblGPSVel = u_GPS.Velocity
    lblGPSDir = u_GPS.Direction
    If chkGPS.Value = 1 Then
        pic.DrawStyle = 0
        pic.FillStyle = 0
        pic.FillColor = vbBlue
        DX = bot.X - u_GPS.X
        DY = bot.Y - u_GPS.Y
    
        pic.Circle (4000 - DX, 4000 - DY), 96, 0
        pic.ForeColor = vbBlack
        If dr.Velocity > 0 Then
            pic.ForeColor = vbBlack
        ElseIf dr.Velocity < 0 Then
            pic.ForeColor = vbRed
        End If
        pic.Line (4000 - DX, 4000 - DY)-(4000 - DX + (25 * bot.VX), 4000 - DY + (25 * bot.VY))
    End If
  End If
End Sub
 

'*******************************************************************
' UpdateLaserRangeFinder
'*******************************************************************
Public Sub UpdateLaserRangeFinder()
  Dim DX, DY As Single 'difference calculations
  Dim i, j As Integer
  Dim nRng As Single 'range and bearing for Laser Rangefinder
  Dim nBrg As Single
  Dim nAng As Single 'angular difference
  Dim nStep As Single
  Dim nM, nD, nU, nL As Single 'used with g_uLR( )
  
  'draws picLR data
  picLR.Cls
  picLR.FillColor = vbBlack
  picLR.FillStyle = 0
  picLR.Circle (1000, 0), 1000
  picLR.ForeColor = vbGreen
  pic.FillColor = vbGreen
  picLR.FillStyle = 1
  
  'reset range values to 2000
  nStep = PI / 36
  For i = 1 To 36
    g_uLR(i).Range = 2000
  Next i
  
  
  'determines range and bearing to obstacles
  For i = 1 To g_nMaxObstacles
    nRng = GetTargetDistance2D(bot.X, bot.Y, g_uOb(i).X, g_uOb(i).Y)
    If nRng < 2000 Then
      nBrg = GetTargetDirection2D(bot.X, bot.Y, g_uOb(i).X, g_uOb(i).Y)
      nAng = GetAngularDifference(bot.Direction, nBrg)
      If nAng < 1.57 And nAng > -1.57 Then
        DX = Cos(nAng - 1.57) * nRng / 2
        DY = Sin(nAng + 1.57) * nRng / 2
        picLR.Circle (1000 + DX, DY), g_uOb(i).Radius / 2, vbGreen
        
        'need to normalize this data to go into array g_uLR( ).Range
        nM = PI / 2 + nAng
        nD = Atn((g_uOb(i).Radius / 1.5) / nRng) 'was 2
        nU = nM + nD
        nL = nM - nD
        
        For j = CInt(nL / nStep) To CInt(nU / nStep)
          If j > 0 And j < 36 Then
            If nRng - g_uOb(i).Radius < g_uLR(j).Range Then g_uLR(j).Range = nRng - g_uOb(i).Radius
          End If
        Next j
        
      End If
    End If
  Next i
  
End Sub

'*******************************************************************
' UpdateRadioBeacons
'*******************************************************************
Public Sub UpdateRadioBeacons()
  Dim i As Integer
  Dim DX, DY As Single
  Dim u_tri As RECT_COORD
  
  If g_nMaxBeacons = 0 Then Exit Sub
  
  'draws beacon
  For i = 1 To g_nMaxBeacons
    DX = bot.X - nav(i).X
    DY = bot.Y - nav(i).Y
    pic.FillColor = vbCyan
    pic.Circle (4000 - DX, 4000 - DY), 64, 0
  Next i
  
  'derives triangulation position
  u_tri = GetTriangulationPosition()
  DX = bot.X - u_tri.X
  DY = bot.Y - u_tri.Y
  pic.FillColor = vbGreen
  pic.Circle (4000 - DX, 4000 - DY), 96, 0
  
End Sub
  
