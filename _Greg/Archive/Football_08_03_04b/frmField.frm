VERSION 5.00
Begin VB.Form frmField 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Field"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   3660
      Top             =   1800
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   11
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Line lnDef 
      BorderColor     =   &H000000C0&
      X1              =   2400
      X2              =   3360
      Y1              =   960
      Y2              =   2400
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   10
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   9
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   8
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   7
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   6
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   5
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   4
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   10
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   9
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   8
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   6
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   5
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape shpVisitor 
      BorderWidth     =   2
      Height          =   135
      Index           =   0
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape shpHome 
      BorderWidth     =   2
      Height          =   135
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1860
      Width           =   135
   End
   Begin VB.Line lnFirstDown 
      BorderColor     =   &H0000FFFF&
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
  Dim i As Integer
  Dim j As Integer
  
  'sizes and scales frmField
  f.Width = frmMain.Width - 200
  f.Height = f.Width / 2.16
  frmField.Width = f.Width
  frmField.Height = f.Height
  frmField.ScaleMode = 0  'user defined mode
  frmField.ScaleTop = f.ScaleTop
  frmField.ScaleLeft = f.ScaleLeft
  frmField.ScaleHeight = f.ScaleHeight
  frmField.ScaleWidth = f.ScaleWidth
  
  'draws football field
  'set autoredraw property for form to TRUE
  frmField.BackColor = RGB(0, 100, 0)         'whole field green
  frmField.ForeColor = RGB(225, 225, 225)
  frmField.Line (-62, 28.67)-(62, -28.67), , BF  'white border
  
  frmField.ForeColor = RGB(0, 120, 0)
  frmField.Line (-50, 26.67)-(50, -26.67), , BF  'green field
  frmField.ForeColor = RGB(0, 120, 0)        'end zones
  frmField.Line (-60, 26.67)-(-50, -26.67), , BF
  frmField.Line (50, 26.67)-(60, -26.67), , BF
  
  'draws hash marks on sidelines and inbound lines
  frmField.ForeColor = vbWhite
  For i = 1 To 99
    frmField.Line (-50 + i, 26.67)-(-50 + i, 25.67)
    frmField.Line (-50 + i, 4.08)-(-50 + i, 3.08)
    frmField.Line (-50 + i, -4.08)-(-50 + i, -3.08)
    frmField.Line (-50 + i, -26.67)-(-50 + i, -25.67)
  Next i
  
  'white end lines
  frmField.Line (-50, 26.67)-(-50, -26.67)
  frmField.Line (50, 26.67)-(50, -26.67)
  
  'goal posts
  frmField.ForeColor = RGB(225, 225, 0)
  frmField.Line (-60, 3.1)-(-59.7, -3.1), , BF
  frmField.Line (60, 3.1)-(59.7, -3.1), , BF
  frmField.ForeColor = RGB(255, 0, 0)
  frmField.Circle (-59.85, 3.1), 0.3
  frmField.Circle (-59.85, -3.1), 0.3
  frmField.Circle (59.85, 3.1), 0.3
  frmField.Circle (59.85, -3.1), 0.3
  
  'yard lines
  frmField.ForeColor = RGB(255, 255, 255)
  For i = 0 To 95 Step 5
    frmField.Line (-50 + i, 26.67)-(-50 + i, -26.67)
  Next i
  frmField.Line (-0.1, 26.67)-(-0.1, -26.67)  '50 yard line
  frmField.Line (0.1, 26.67)-(0.1, -26.67)
  
  'yard values on south side of field
  '10 to 50 yard
  frmField.FontSize = 18
  frmField.ForeColor = RGB(225, 225, 225)
  For i = 1 To 5
     frmField.CurrentX = -50 + (i * 10) - 2.8  '1st digit
     frmField.CurrentY = -15
     frmField.Print i
     frmField.CurrentX = -50 + (i * 10) '2nd digit
     frmField.CurrentY = -15
     frmField.Print 0
  Next i
  
  '40 to 10 yard
  For i = 4 To 1 Step -1
     frmField.CurrentX = (i * 10) - 2.8  '1st digit
     frmField.CurrentY = -15
     frmField.Print 5 - i
     frmField.CurrentX = (i * 10)  '2nd digit
     frmField.CurrentY = -15
     frmField.Print 0
  Next i
  
  'labels endzones
  frmField.FontSize = 30
  frmField.ForeColor = RGB(105, 0, 0)
  frmField.CurrentX = -57
  frmField.CurrentY = 20
  frmField.Print "W"
  frmField.CurrentX = -57
  frmField.CurrentY = 10
  frmField.Print "E"
  frmField.CurrentX = -57
  frmField.CurrentY = 0
  frmField.Print "S"
  frmField.CurrentX = -57
  frmField.CurrentY = -10
  frmField.Print "T"
  frmField.ForeColor = RGB(0, 0, 105)
  frmField.CurrentX = 53
  frmField.CurrentY = 20
  frmField.Print "E"
  frmField.CurrentX = 53
  frmField.CurrentY = 10
  frmField.Print "A"
  frmField.CurrentX = 53
  frmField.CurrentY = 0
  frmField.Print "S"
  frmField.CurrentX = 53
  frmField.CurrentY = -10
  frmField.Print "T"

  
  'generate player positions
  GeneratePositions
  For i = 0 To 11
    shpHome(i).Shape = 3
    shpVisitor(i).Shape = 3
  Next i
  
  'draw markers and ball
  lnOff.X1 = f.OffLine
  lnOff.X2 = f.OffLine
  lnOff.Y1 = 26.67
  lnOff.Y2 = -26.67
  lnDef.X1 = f.DefLine
  lnDef.X2 = f.DefLine
  lnDef.Y1 = 26.67
  lnDef.Y2 = -26.67
  lnFirstDown.X1 = f.FirstDownLine
  lnFirstDown.X2 = f.FirstDownLine
  lnFirstDown.Y1 = 26.67
  lnFirstDown.Y2 = -26.67
  shpBall.Left = f.BallX - 0.5
  shpBall.Top = f.BallY
  'shpQB.Left = f.BallX - 1
  'shpQB.Top = f.BallY + 0.5
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

'redraws marker lines and players
Private Sub tmrUpdate_Timer()
  Dim i As Integer
    
  'draw markers And ball
  lnOff.X1 = f.OffLine
  lnOff.X2 = f.OffLine
  lnOff.Y1 = 26.67
  lnOff.Y2 = -26.67
  lnDef.X1 = f.DefLine
  lnDef.X2 = f.DefLine
  lnDef.Y1 = 26.67
  lnDef.Y2 = -26.67
  lnFirstDown.X1 = f.FirstDownLine
  lnFirstDown.X2 = f.FirstDownLine
  lnFirstDown.Y1 = 26.67
  lnFirstDown.Y2 = -26.67
  
  'position players
  For i = 0 To 11
    shpHome(i).Left = p(i).X
    shpVisitor(i).Left = p(i + 12).X
    shpHome(i).Top = p(i).Y
    shpVisitor(i).Top = p(i + 12).Y
    shpHome(i).Width = 0.75
    shpHome(i).Height = 0.75
    shpVisitor(i).Width = 0.75
    shpVisitor(i).Height = 0.75
    shpHome(i).BorderColor = vbRed
    shpVisitor(i).BorderColor = vbBlue
  Next i
  
  'shpBall.Left = f.BallX - 0.5
  'shpBall.Top = f.BallY
  'shpQB.Left = f.QBX - 1
  'shpQB.Top = f.QBY + 0.5
  'shpRec.Left = f.RecX - 1
  'shpRec.Top = f.RecY + 0.5
  
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
