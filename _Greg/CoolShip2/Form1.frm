VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1320
      Top             =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2640
      X2              =   2340
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_KeyDown(255) As Boolean

Private Sub Form_Activate()
  
  'scales vertical axis
  Form1.ScaleLeft = 0
  Form1.ScaleWidth = Form1.Width
  Form1.ScaleTop = Form1.Height
  Form1.ScaleHeight = -Form1.Height
      
  'initialize ship
  S.X = Form1.Width / 2
  S.Y = Form1.Height / 2
  S.Heading = PI / 4
  S.Angle = S.Heading
  S.Velocity = 0
  S.Thrust = S.Velocity
      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  m_KeyDown(KeyCode) = True
    
  If KeyCode = vbKeyEscape Then End
  
  'change physical orientation of ship's direction
  If KeyCode = vbKeyLeft Then
    S.Angle = S.Angle + 0.1
    If S.Angle > 2 * PI Then S.Angle = 0
  End If
  
  If KeyCode = vbKeyRight Then
    S.Angle = S.Angle - 0.1
    If S.Angle < 0 Then S.Angle = 2 * PI
  End If
  
  If KeyCode = vbKeyUp Then
    S.Thrust = S.Thrust + STEP_VELOCITY
    If S.Thrust >= MAX_VELOCITY Then S.Thrust = MAX_VELOCITY
  End If
  If KeyCode = vbKeyDown Then
    S.Thrust = S.Thrust - STEP_VELOCITY
    If S.Thrust <= MIN_VELOCITY Then S.Thrust = MIN_VELOCITY
  End If
 
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  m_KeyDown(KeyCode) = False
End Sub

'update graphic display
Private Sub Timer1_Timer()

  Dim TX, TY As Single
  Dim vx, vy As Single
  
  'calculate components of thrust to velocity
  If m_KeyDown(vbKeyUp) = True Or m_KeyDown(vbKeyDown) = True Then
    TX = Cos(S.Angle) * S.Thrust
    TY = Sin(S.Angle) * S.Thrust
  End If

  'add velocity vector and thrust vector
  vx = Cos(S.Heading) * S.Velocity + TX
  vy = Sin(S.Heading) * S.Velocity + TY
  
  'calculate ships new velocity as a result of thrusting
  S.Velocity = Sqr(vx ^ 2 + vy ^ 2)
  S.Heading = GetAngle(vx, vy)

  'update ship position
  S.X = S.X + Cos(S.Heading) * S.Velocity
  S.Y = S.Y + Sin(S.Heading) * S.Velocity
  
  'boundary controls for x,y of ship
  If S.X < 0 Then S.X = Form1.ScaleWidth
  If S.X > Form1.ScaleWidth Then S.X = 0
  If S.Y < 0 Then S.Y = Form1.ScaleTop
  If S.Y > Form1.ScaleTop Then S.Y = 0
  
  'update line end points
  Line1.X1 = S.X
  Line1.Y1 = S.Y
  Line1.X2 = Line1.X1 + Cos(S.Angle) * SHIP_LENGTH
  Line1.Y2 = Line1.Y1 + Sin(S.Angle) * SHIP_LENGTH
  
End Sub
