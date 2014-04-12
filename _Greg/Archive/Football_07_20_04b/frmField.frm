VERSION 5.00
Begin VB.Form frmField 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Field"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   3660
      Top             =   1800
   End
   Begin VB.Shape shpRec 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   150
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   480
      Width           =   150
   End
   Begin VB.Shape shpQB 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   150
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   720
      Width           =   150
   End
   Begin VB.Line lnFirstDown 
      BorderColor     =   &H00FF0000&
      X1              =   1860
      X2              =   2580
      Y1              =   1560
      Y2              =   2820
   End
   Begin VB.Shape shpBall 
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   2820
      Shape           =   3  'Circle
      Top             =   720
      Width           =   75
   End
   Begin VB.Line lnDef 
      BorderColor     =   &H000000C0&
      X1              =   1740
      X2              =   2700
      Y1              =   900
      Y2              =   2340
   End
   Begin VB.Line lnOff 
      BorderColor     =   &H0000FFFF&
      X1              =   1140
      X2              =   1800
      Y1              =   1140
      Y2              =   2580
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  'sizes and scales frmField
  frmField.Width = f.Width
  frmField.Height = f.Height
  frmField.ScaleMode = 0  'user defined mode
  frmField.ScaleTop = f.ScaleTop
  frmField.ScaleLeft = f.ScaleLeft
  frmField.ScaleHeight = f.ScaleHeight
  frmField.ScaleWidth = f.ScaleWidth
  
  'draws football field
  'set autoredraw property for form to TRUE
  frmField.ForeColor = RGB(0, 150, 0)
  frmField.Line (-50, 25)-(50, -25), , BF
  frmField.ForeColor = vbWhite
  frmField.Line (-60, 30)-(60, 25), , BF
  frmField.Line (-60, -25)-(60, -30), , BF
  
  'draw markers and ball
  lnOff.X1 = f.OffLine
  lnOff.X2 = f.OffLine
  lnOff.Y1 = 25
  lnOff.Y2 = -25
  lnDef.X1 = f.DefLine
  lnDef.X2 = f.DefLine
  lnDef.Y1 = 25
  lnDef.Y2 = -25
  lnFirstDown.X1 = f.FirstDownLine
  lnFirstDown.X2 = f.FirstDownLine
  lnFirstDown.Y1 = 25
  lnFirstDown.Y2 = -25
  shpBall.Left = f.BallX - 0.5
  shpBall.Top = f.BallY
  shpQB.Left = f.BallX - 1
  shpQB.Top = f.BallY + 0.5
  If f.DefLineOn = True Then
    lnDef.Visible = True
  Else
    lnDef.Visible = False
  End If
  If f.OffLineOn = True Then
    lnOff.Visible = True
  Else
    lnOff.Visible = False
  End If
  If f.FirstDownLineOn = True Then
    lnFirstDown.Visible = True
  Else
    lnFirstDown.Visible = False
  End If
End Sub

Private Sub tmrUpdate_Timer()
 'draw markers and ball
  lnOff.X1 = f.OffLine
  lnOff.X2 = f.OffLine
  lnOff.Y1 = 25
  lnOff.Y2 = -25
  lnDef.X1 = f.DefLine
  lnDef.X2 = f.DefLine
  lnDef.Y1 = 25
  lnDef.Y2 = -25
  lnFirstDown.X1 = f.FirstDownLine
  lnFirstDown.X2 = f.FirstDownLine
  lnFirstDown.Y1 = 25
  lnFirstDown.Y2 = -25
  
  shpBall.Left = f.BallX - 0.5
  shpBall.Top = f.BallY
  shpQB.Left = f.QBX - 1
  shpQB.Top = f.QBY + 0.5
  shpRec.Left = f.RecX - 1
  shpRec.Top = f.RecY + 0.5
  
  If f.DefLineOn = True Then
    lnDef.Visible = True
  Else
    lnDef.Visible = False
  End If
  If f.OffLineOn = True Then
    lnOff.Visible = True
  Else
    lnOff.Visible = False
  End If
  If f.FirstDownLineOn = True Then
    lnFirstDown.Visible = True
  Else
    lnFirstDown.Visible = False
  End If


End Sub
