VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Robot Navigation"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Dead Reckoning"
      Height          =   1815
      Left            =   8160
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
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
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Vel:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   495
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
      Height          =   2175
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
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
      Left            =   8880
      Top             =   5160
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   8000
      Left            =   120
      ScaleHeight     =   -8000
      ScaleMode       =   0  'User
      ScaleTop        =   8000
      ScaleWidth      =   8000
      TabIndex        =   0
      Top             =   120
      Width           =   8000
      Begin VB.Label lblDRPos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DR Pos."
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bAutoMode As Boolean

Private Sub Form_Load()
  LoadVariables


End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
  'frmMain.Caption = KeyCode
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
    
    g_nLastLegNum = 5
  ElseIf KeyCode = 32 Then 'space - toggle timer on/off
    If tmrUpdate.Enabled = True Then
      tmrUpdate.Enabled = False
    Else
      tmrUpdate.Enabled = True
    End If
  ElseIf KeyCode = 65 Then  'A
    If m_bAutoMode = True Then
      m_bAutoMode = False   'disable autonomous mode
    Else
      m_bAutoMode = True    'operate in autonomous mode
    End If
  Else
    'frmMain.Caption = KeyCode
  End If
  
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'frmMain.Caption = X & " " & Y
  
End Sub

Private Sub tmrUpdate_Timer()
  Dim DX, DY As Single 'difference calculations
  Dim cx, cy As Single
  Dim i As Integer
  Dim u_tri As RECT_COORD
  
  'calls sub to generate 'turn' and 'velocity' variables
  If m_bAutoMode = True Then Autonomous
  'Autonomous
  
  'redraw everything
  pic.Cls
  
  'update bot data...actual stuff
  bot.Energy = bot.Energy - bot.Velocity
  
  'calc actual bot data
  If bot.Energy > 0 Then
    bot.Direction = bot.Direction + bot.Turn
    If bot.Direction > 6.28 Then bot.Direction = 0
    If bot.Direction < 0 Then bot.Direction = 6.28
    bot.VX = bot.Velocity * Cos(bot.Direction)
    bot.VY = bot.Velocity * Sin(bot.Direction)
    bot.X = bot.X + bot.VX
    bot.Y = bot.Y + bot.VY
  Else
    bot.Velocity = 0
    bot.VX = 0
    bot.VY = 0
    bot.Energy = 0
  End If
  
  'calculate DR information
  dr.Velocity = bot.Velocity / NAV_DR_VEL_FACTOR
  dr.Direction = bot.Direction / NAV_DR_DIR_FACTOR
  If dr.Direction > 6.28 Then dr.Direction = 0
  If dr.Direction < 0 Then dr.Direction = 6.28
  dr.VX = Cos(dr.Direction) * dr.Velocity
  dr.VY = Sin(dr.Direction) * dr.Velocity
  dr.X = dr.X + dr.VX
  dr.Y = dr.Y + dr.VY
  
  'draws Dead Reckoning on screen
  pic.DrawStyle = 0
  pic.FillColor = vbMagenta
  DX = dr.X - bot.X
  DY = dr.Y - bot.Y
  'pic.Circle (4000 - DX, 4000 - DY), 96, 0
  'pic.ForeColor = vbBlack
  'If dr.Velocity > 0 Then
  '  pic.ForeColor = vbBlack
  'ElseIf dr.Velocity < 0 Then
  '  pic.ForeColor = vbRed
  'End If
  'pic.Line (4000 - DX, 4000 - DY)-(4000 - DX + (25 * dr.VX), 4000 - DY + (25 * dr.VY))
  lblDRPos.Left = 4050 - DX: lblDRPos.Top = 4050 - DY
  
  'draws beacon
  'NAV_TRIANGULATION_FACTOR
  For i = 1 To 3
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
      
  'draws bot
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
  
  'dead-reckoning information
  lblDRX = FormatNumber(dr.X, 1)
  lblDRY = FormatNumber(dr.Y, 1)
  lblDRVel = FormatNumber(dr.Velocity, 1)
  lblDRDir = FormatNumber(RtoC(dr.Direction), 1)
    
End Sub

