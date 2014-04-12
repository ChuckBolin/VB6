VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Robot v0.11 - Written by Team 342, Chuck Bolin (Mentor)"
   ClientHeight    =   7365
   ClientLeft      =   1155
   ClientTop       =   765
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   9390
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   6870
   End
   Begin VB.Timer tmrAutoUpdate 
      Interval        =   26
      Left            =   510
      Top             =   6870
   End
   Begin VB.TextBox txtVC 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5955
      Left            =   6330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   240
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Robot Mode"
      Height          =   1215
      Left            =   3120
      TabIndex        =   16
      Top             =   2970
      Width           =   1965
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal (Manual)"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   870
         Width           =   1575
      End
      Begin VB.OptionButton optDisabled 
         Caption         =   "Disabled"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   570
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Autonomous"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.HScrollBar hsbDirection 
      Height          =   225
      LargeChange     =   31
      Left            =   1020
      Max             =   628
      TabIndex        =   6
      Top             =   6600
      Width           =   1995
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8100
      TabIndex        =   5
      Top             =   6930
      Width           =   1245
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   6870
   End
   Begin VB.PictureBox picJoy1 
      BackColor       =   &H00000000&
      Height          =   2000
      Left            =   3060
      ScaleHeight     =   -255
      ScaleLeft       =   255
      ScaleMode       =   0  'User
      ScaleTop        =   255
      ScaleWidth      =   -255
      TabIndex        =   1
      Top             =   810
      Width           =   2000
      Begin VB.Line linJoy1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   7
         X1              =   124.535
         X2              =   104.767
         Y1              =   124.535
         Y2              =   152.209
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         X1              =   127.039
         X2              =   127.039
         Y1              =   255
         Y2              =   -1.977
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         X1              =   255
         X2              =   -1.977
         Y1              =   127.039
         Y2              =   127.039
      End
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00808080&
      Height          =   6000
      Left            =   30
      ScaleHeight     =   -54
      ScaleLeft       =   -13.5
      ScaleMode       =   0  'User
      ScaleTop        =   27
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   570
      Width           =   3000
      Begin VB.Line linRef 
         BorderColor     =   &H00FFFF00&
         X1              =   -0.827
         X2              =   -1.929
         Y1              =   2.727
         Y2              =   0
      End
      Begin VB.Shape shpCenter 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   55
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   2940
         Width           =   54
      End
      Begin VB.Line linFront 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -5.235
         X2              =   3.857
         Y1              =   3.545
         Y2              =   1.909
      End
      Begin VB.Line linLeft 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   -8.541
         X2              =   -10.745
         Y1              =   4.636
         Y2              =   -2.182
      End
      Begin VB.Line linBack 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -6.337
         X2              =   0.827
         Y1              =   -3.273
         Y2              =   -5.455
      End
      Begin VB.Line linRight 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         X1              =   3.857
         X2              =   1.929
         Y1              =   1.909
         Y2              =   -4.636
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Digital Display"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   585
      Left            =   5340
      TabIndex        =   22
      Top             =   810
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Autonomous Virtual Code"
      Height          =   225
      Left            =   6840
      TabIndex        =   21
      Top             =   0
      Width           =   1905
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Manual Control Joystick"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3150
      TabIndex        =   15
      Top             =   570
      Width           =   1785
   End
   Begin VB.Label lblX 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3900
      TabIndex        =   14
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblLeft 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   990
      TabIndex        =   13
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Initial Angle:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   60
      TabIndex        =   12
      Top             =   6600
      Width           =   915
   End
   Begin VB.Label label6 
      BackColor       =   &H00000000&
      Caption         =   "Position X:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   0
      Left            =   3030
      TabIndex        =   11
      Top             =   0
      Width           =   825
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Position Y:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   4410
      TabIndex        =   10
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lblY 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   5280
      TabIndex        =   9
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblDir 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   990
      TabIndex        =   8
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Direction:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   30
      TabIndex        =   7
      Top             =   240
      Width           =   915
   End
   Begin VB.Label lblRight 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Right Speed:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1500
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Left Speed:"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'draws robot
Public Sub DrawRobot()
      
  'gets new values
  'VR.Direction = CtoR(hsbDir.Value)
  shpCenter.Left = VR.Center.X - 0.25
  shpCenter.Top = VR.Center.Y + 0.25
  'frmMain.Caption = VR.Center.X & "  " & VR.Center.Y
  
  
  'draws reference line
  linRef.X1 = VR.Center.X
  linRef.Y1 = VR.Center.Y
  linRef.X2 = VR.Reference.X
  linRef.Y2 = VR.Reference.Y
  
  'draws left side
  linLeft.X1 = VR.LeftFront.X
  linLeft.Y1 = VR.LeftFront.Y
  linLeft.X2 = VR.LeftBack.X
  linLeft.Y2 = VR.LeftBack.Y
    
  'draws right side
  linRight.X1 = VR.RightFront.X
  linRight.Y1 = VR.RightFront.Y
  linRight.X2 = VR.RightBack.X
  linRight.Y2 = VR.RightBack.Y
    
  'draws front side
  linFront.X1 = VR.LeftFront.X
  linFront.Y1 = VR.LeftFront.Y
  linFront.X2 = VR.RightFront.X
  linFront.Y2 = VR.RightFront.Y
  
  'draws back side
  linBack.X1 = VR.LeftBack.X
  linBack.Y1 = VR.LeftBack.Y
  linBack.X2 = VR.RightBack.X
  linBack.Y2 = VR.RightBack.Y
  
End Sub

Private Sub cmdExit_Click()
  End
End Sub



'initialize form and robot when starting program
Private Sub Form_Activate()
  Dim bReturn As Boolean
  
  linJoy1.X1 = 127
  linJoy1.Y1 = 127
  linJoy1.X2 = 127
  linJoy1.Y2 = 127
  
  LoadRobotVariables
  UpdateRobot
  DrawRobot
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
  hsbDirection_Change
  lblLeft = Format(GetWheelSpeed(VR.LeftMotor), "##.#")
  lblRight = Format(GetWheelSpeed(VR.RightMotor), "##.#")
  txtVC = LoadSampleVirtualCode
  bReturn = LoadVirtualCodeIntoArray(txtVC)
  If bReturn = False Then
    MsgBox "Error loading Virtual Code into program array!"
    End
  End If
End Sub

'manually orienting the robot
Private Sub hsbDirection_Scroll()
  hsbDirection_Change
End Sub

'manually orienting the robot
Private Sub hsbDirection_Change()
  'If tmrUpdate.Enabled = True Then Exit Sub
  If optDisabled.Value = False Then Exit Sub
  VR.Direction = hsbDirection.Value * 0.01
  UpdateRobot
  DrawRobot
  lblDir = Format(RtoC(VR.Direction), "###.#")
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
End Sub

Private Sub optAuto_Click()
  tmrUpdate.Enabled = False    'manual mode
  tmrCountdown.Enabled = True  '15 second countdown
  tmrAutoUpdate.Enabled = True 'drives 26.2mSec code
  g_nCounter = 15
  lblTime = g_nCounter
  hsbDirection.Enabled = False
  lblDir = Format(RtoC(VR.Direction), "###.#")
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
  lblLeft = Format(GetWheelSpeed(VR.LeftMotor), "##.#")
  lblRight = Format(GetWheelSpeed(VR.RightMotor), "##.#")

End Sub

Private Sub optDisabled_Click()
  tmrUpdate.Enabled = False
  hsbDirection.Enabled = True
  lblDir = Format(RtoC(VR.Direction), "###.#")
  hsbDirection.Value = CInt(RtoC(VR.Direction))
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
  
  'reset values
  VR.Joy_X = 127
  VR.Joy_Y = 127
  linJoy1.X2 = VR.Joy_X
  linJoy1.Y2 = VR.Joy_Y
  VR.RightMotor = Limit_Mix(2000 + VR.Joy_Y + VR.Joy_X - 127)
  VR.LeftMotor = Limit_Mix(2000 + VR.Joy_Y - VR.Joy_X + 127)
  lblLeft = Format(GetWheelSpeed(VR.LeftMotor), "##.#")
  lblRight = Format(GetWheelSpeed(VR.RightMotor), "##.#")
  
End Sub

Private Sub optNormal_Click()
  tmrUpdate.Enabled = True
  hsbDirection.Enabled = False
  lblDir = Format(RtoC(VR.Direction), "###.#")
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
  lblLeft = Format(GetWheelSpeed(VR.LeftMotor), "##.#")
  lblRight = Format(GetWheelSpeed(VR.RightMotor), "##.#")

End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If tmrUpdate.Enabled = True Then Exit Sub
  If optDisabled.Value = False Then Exit Sub
  VR.Center.X = X
  VR.Center.Y = Y
  UpdateRobot
  DrawRobot
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
End Sub

Private Sub picJoy1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  VR.Joy_X = CByte(X)
  VR.Joy_Y = CByte(Y)
  linJoy1.X2 = VR.Joy_X
  linJoy1.Y2 = VR.Joy_Y
  VR.RightMotor = Limit_Mix(2000 + VR.Joy_Y + VR.Joy_X - 127)
  VR.LeftMotor = Limit_Mix(2000 + VR.Joy_Y - VR.Joy_X + 127)
  lblLeft = Format(GetWheelSpeed(VR.LeftMotor), "##.#")
  lblRight = Format(GetWheelSpeed(VR.RightMotor), "##.#")
End Sub

'this calls VM processor to move robot
Private Sub tmrAutoUpdate_Timer()
  ProcessVirtualCode
End Sub

Private Sub tmrCountdown_Timer()
  g_nCounter = g_nCounter - 1
  lblTime = g_nCounter
  If g_nCounter <= 0 Then
    tmrCountdown.Enabled = False
    tmrAutoUpdate.Enabled = False
    optNormal.Value = True
  End If
End Sub

Private Sub tmrUpdate_Timer()
  UpdateRobot
  DrawRobot
  lblDir = Format(RtoC(VR.Direction), "###.#")
  lblX = Format(VR.Center.X, "##.#")
  lblY = Format(VR.Center.Y, "##.#")
End Sub
