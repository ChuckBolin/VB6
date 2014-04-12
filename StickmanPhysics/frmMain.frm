VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stickman Physics v0.001"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   683
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   974
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   6360
      Top             =   8040
   End
   Begin VB.Label lblAngle 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDistance 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblXY 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReset_Click()
  j(0).X = 400
  j(0).Y = 300
  j(0).VX = 0
  j(0).VY = 0
  j(1).X = 400
  j(1).Y = 400
  j(1).VX = 0
  j(1).VY = 0
  
End Sub

Private Sub Form_Activate()
  j(0).X = 400
  j(0).Y = 300
  j(0).VX = 0
  j(0).VY = 0
  j(1).X = 400
  j(1).Y = 400
  j(1).VX = 0
  j(1).VY = 0
  g_segment = 100
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim distance As Single
  Dim angle As Single
  
  distance = GetDistance(X, Y, j(0).X, j(0).Y)
  lblDistance.Caption = "Distance: " & CInt(distance)
  angle = GetAngle(X, Y, j(0).X, j(0).Y)
  lblAngle.Caption = "Angle: " & angle
  g_X = X
  g_Y = Y
  If j(0).X > X Then
    j(0).VX = (j(0).X - X) / 10
  Else
    j(0).VX = -(X - j(0).X) / 10
  End If
  
  If j(0).Y > Y Then
    j(0).VY = -(Y - j(0).Y) / 10
  Else
    j(0).VY = (j(0).Y - Y) / 10
  End If
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblXY.Caption = "X: " & CInt(X) & "  Y: " & CInt(Y)
  frmMain.Line (j(0).X, j(0).Y)-(j(0).X + (j(0).X - X), j(0).Y + (j(0).Y - Y))
End Sub

Private Function GetDistance(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
  GetDistance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Private Function GetAngle(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
  Dim angle As Single
  
  If x2 - x1 = 0 Then
    angle = Atn((y2 - y1) / 0.001)
  Else
    angle = Atn((y2 - y1) / (x2 - x1))
  End If
  
  If y1 > y2 Then 'bottom
    If x2 > x1 Then
      angle = angle * -1
    ElseIf x2 < x1 Then
      angle = (angle - PI) * -1
    Else  'x1 = x2
    
    End If
  ElseIf y2 > y1 Then
    If x2 > x1 Then
      angle = 2 * PI + angle * -1
      'angle = angle + ((angle - PI) * -1)
    Else
      angle = PI + (angle * -1)
    End If
  
  End If
  
  GetAngle = angle
End Function

Private Sub Update()
  j(0).X = j(0).X + j(0).VX
  j(0).Y = j(0).Y + j(0).VY
  Dim angle As Single
  
  angle = GetAngle(j(1).X, j(1).Y, j(0).X, j(0).Y)
  Dim dx As Single
  Dim dy As Single
  dx = angle * g_segment
  dy = angle * g_segment
  frmMain.Caption = dx & ":" & dy
  j(1).X = j(0).X - dx / 2
  j(1).Y = j(0).Y + dy / 2
  
End Sub

Private Sub tmrUpdate_Timer()
  Update
  frmMain.Cls
  DrawForm
End Sub

Private Sub DrawForm()
  frmMain.Circle (j(0).X, j(0).Y), 5
  frmMain.Circle (j(1).X, j(1).Y), 5
  frmMain.ForeColor = vbGreen
  frmMain.Line (j(0).X, j(0).Y)-(j(1).X, j(1).Y)
  frmMain.ForeColor = vbRed
  frmMain.Line (j(0).X, j(0).Y)-(j(0).X + (j(0).X - g_X), j(0).Y + (j(0).Y - g_Y))
  frmMain.ForeColor = vbBlack

End Sub
