VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avoidance v0.1"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   -73.636
   ScaleLeft       =   -60
   ScaleMode       =   0  'User
   ScaleTop        =   30
   ScaleWidth      =   120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   1440
      Top             =   6000
   End
   Begin VB.CommandButton cmdDist 
      Caption         =   "&Distribute"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   29
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   28
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   27
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   26
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   25
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   24
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   23
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   22
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   21
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   20
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   19
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   18
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   17
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   16
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   270
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stopped!"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   15
      Left            =   600
      Shape           =   3  'Circle
      Top             =   240
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   14
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   13
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   12
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   11
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   10
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   9
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   8
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   7
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   6
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   266
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   5
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   4
      Left            =   840
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   3
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   960
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   2
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   1
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   270
   End
   Begin VB.Shape shpOb 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   352
      Index           =   0
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   266
   End
   Begin VB.Shape shpQB 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   352
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   266
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BOTS = 30
Private b(BOTS) As BOT
Private q As BOT  'quarterback moves around obstacles
Private t As BOT  'target position
Private bFind As Boolean 'true if q is searching for t
Private Const BOTWIDTH = 8

'redistributes all 16 obstacles
Private Sub cmdDist_Click()
 GeneratePositions
End Sub

'starts and stops q movement towards destination
Private Sub cmdFind_Click()
  If bFind = True Then
    bFind = False
    cmdFind.Caption = "&Find"
    lblMsg.Caption = "Stopped!!!"
        
    'antistuck initialization
    q.rx = q.x
    q.ry = q.y
    q.Count = 0
  Else
    bFind = True
    cmdFind.Caption = "&Stop"
    lblMsg.Caption = "Finding..."
    
  End If
End Sub

'initialization stuff
Private Sub Form_Load()
  Dim i As Integer
  For i = 0 To 29
    shpOb(i).Height = BOTWIDTH
    shpOb(i).WIDTH = BOTWIDTH
  Next i
  
  bFind = False
  t.x = 30
  t.y = 0
  q.x = -30
  q.y = 0
  q.Vel = 1.8
  q.Dir = 0
  q.Range = BOTWIDTH '* 4
  q.tx = t.x
  q.ty = t.y
  q.Mode = BOT_FINDING
  shpQB.Left = q.x - 2
  shpQB.Top = q.y + 2
  shpQB.WIDTH = BOTWIDTH
  shpQB.Height = BOTWIDTH
  Randomize Timer
  GeneratePositions
End Sub

'generates positions for 16 objects and attaches shapes
Private Sub GeneratePositions()
  Dim i As Integer
  
  For i = 0 To BOTS - 1
    b(i).x = -58 + Rnd() * 116
    b(i).y = -28 + Rnd() * 56
    shpOb(i).Left = b(i).x - 2
    shpOb(i).Top = b(i).y + 2
  Next i
  
End Sub

'positions red circle (quarterback) and black circle (target)
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    q.x = x
    q.y = y
    frmMain.Circle (t.x, t.y), 1
  End If
  If Button = 2 Then
    t.x = x
    t.y = y
    q.tx = t.x
    q.ty = t.y
    frmMain.Cls
    frmMain.Circle (t.x, t.y), 1
  End If
End Sub

'initial drawing of target circle
Private Sub Form_Resize()
    frmMain.Circle (t.x, t.y), 1
End Sub

'animates movement
Private Sub tmrUpdate_Timer()
  Dim dx As Single  'holds incremental change in q position x,y
  Dim dy As Single
  
  'move quarterback
  If bFind = True And q.Mode = BOT_FINDING Then
    q.Dir = GetBestDirection  'avoids obstacles
    dx = q.Vel * Sin(q.Dir)
    dy = q.Vel * Cos(q.Dir)
    q.x = q.x + dx
    q.y = q.y + dy
    
    'antistuck code
    q.Count = q.Count + 1
    If q.Count > 5 Then
      q.Count = 0
      
      'haven't gotten anywhere in 20 timer clicks
      'so go somewhere else
      If GetTargetDistance(q.x, q.y, q.rx, q.ry) < q.Vel * 1 Then
        q.rx = -58 + Rnd() * 116
        q.ry = -28 + Rnd() * 56
        q.Stuck = True 'Not q.Stuck
        lblMsg.Caption = "Stuck"
      Else
        q.rx = q.x
        q.ry = q.y
        q.Stuck = False 'Not q.Stuck
        lblMsg.Caption = "Finding..."
      End If
    
    End If
    
  End If
  If GetTargetDistance(q.x, q.y, t.x, t.y) < q.Vel Then
    q.Mode = BOT_FOUND
  Else
    q.Mode = BOT_FINDING
  End If
  shpQB.Left = q.x - 2
  shpQB.Top = q.y + 2
End Sub

'algorithm to determine best direction to travel to avoid
'obstacles and to get to destination
Private Function GetBestDirection()
  Dim i As Integer
  Dim nRange As Single 'stores temp range
  Dim nDir As Single 'stores temp direction
  Dim nAngle As Single 'stores +/- angle offset to obstacle
  Dim nLo As Single  'holds two possible directions to move
  Dim nHi As Single
  Dim nBestDir As Single 'direction to target destination
  
  'default direction if there are no obstacles
  If q.Stuck = True Then 'stuck...get out
    nBestDir = GetTargetDirection(q.x, q.y, q.rx, q.ry)
  Else
    nBestDir = GetTargetDirection(q.x, q.y, q.tx, q.ty)
  End If
  nLo = nBestDir
  nHi = nBestDir
  
  'scan through all objects and evaluate them
  'if they are within q.range
  For i = 0 To BOTS - 1
    nRange = GetTargetDistance(q.x, q.y, b(i).x, b(i).y)
    
    'this is, evaluate direction of obstacle
    If nRange < q.Range And q.Range <> 0 Then 'prevents divide by zero
      nDir = GetTargetDirection(q.x, q.y, b(i).x, b(i).y)
      
      'look only at those in front +/- 90 degrees
      If nDir > q.Dir - 1.57 Or nDir < q.Dir + 1.57 Then
        nAngle = Atn(1.5 * BOTWIDTH / nRange)
        
        'nLo needs decreased
        If nLo > nDir - nAngle And nLo < nDir + nAngle Then
          nLo = nDir - nAngle
        End If
        
        'nHi needs increased
        If nHi > nDir - nAngle And nHi < nDir + nAngle Then
          nHi = nDir + nAngle
        End If
        
        'this means that there is an obstacle in the way
        'If nBestDir > nDir - nAngle And nBestDir < nDir + nAngle Then
        '  nLo = nDir - nAngle
        '  nHi = nDir + nAngle
        'End If
      End If
    End If
  Next i

  'choose best direction from options
  
  If nLo = nBestDir Or nHi = nBestDir Then
    GetBestDirection = nBestDir
    Exit Function
  
  'the low angle is closer to best dir then high angle
  ElseIf Abs(nLo - nBestDir) < Abs(nHi - nBestDir) Then
    GetBestDirection = nLo
    Exit Function
  
  'the high angle is closer to best dir then low angle
  ElseIf Abs(nLo - nBestDir) > Abs(nHi - nBestDir) Then
    GetBestDirection = nHi
    Exit Function
  
  End If
End Function
