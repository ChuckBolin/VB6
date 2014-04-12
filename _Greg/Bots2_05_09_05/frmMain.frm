VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bot Demo"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   8160
      TabIndex        =   10
      Top             =   6360
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   8160
      TabIndex        =   9
      Top             =   5820
      Width           =   915
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   8340
      Top             =   7740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   8100
      Width           =   7995
      Begin VB.CheckBox chkIntermediate 
         Caption         =   "Intermediate Pos."
         Height          =   255
         Left            =   1980
         TabIndex        =   13
         Top             =   1020
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkAvoid 
         Caption         =   "Avoidance Lines"
         Height          =   255
         Left            =   1980
         TabIndex        =   12
         Top             =   660
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkTarget 
         Caption         =   "Show Targets"
         Height          =   255
         Left            =   1980
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtData 
         Height          =   1215
         Left            =   6180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.VScrollBar vsbRange 
         Height          =   795
         Left            =   5640
         Max             =   5
         Min             =   100
         SmallChange     =   10
         TabIndex        =   7
         Top             =   600
         Value           =   15
         Width           =   315
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Left            =   5520
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkCircle 
         Caption         =   "Bot Circles"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1020
         Width           =   1815
      End
      Begin VB.CheckBox chkQuad 
         Caption         =   "Search Quad"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.CheckBox chkDir 
         Caption         =   "Direction Vector"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Search Range:"
         Height          =   255
         Left            =   4380
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8000
      Left            =   60
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   60
      Width           =   8000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If tmrUpdate.Interval = 50 Then
    tmrUpdate.Interval = 500
  Else
    tmrUpdate.Interval = 50
  End If
End Sub

Private Sub Command2_Click()
  tmrUpdate.Enabled = Not tmrUpdate.Enabled
End Sub

Private Sub Form_Click()
  If tmrUpdate.Interval = 50 Then
  tmrUpdate.Interval = 500
  Else
  tmrUpdate.Interval = 50
  End If
   'txtData.Text = b.GetVelocity(BOT_MAIN)
End Sub

'initialization
Private Sub Form_Load()
  Randomize Timer
  vsbRange_Change
  LoadBotData    'load all bot information
  DrawBots
End Sub

'***************************************** DrawBots
'draws all bots
Private Sub DrawBots()
  Dim i As Integer
  Dim j As Integer
  
  Dim nPair As PAIR
  
  pic.Cls
  
  'sets color points based upon bot( )
  pic.DrawWidth = 2
  pic.ForeColor = vbWhite
  For i = 1 To b.GetMaxBots
    pic.PSet (b.GetX(i), b.GetY(i))
  Next i
  
  'draws target X and Y
  If chkTarget.Value = vbChecked Then
    pic.ForeColor = vbYellow
    pic.DrawStyle = 0
    pic.DrawWidth = 1
    For i = 1 To b.GetMaxBots '1 To b.GetMaxBots
      DoEvents
      pic.Circle (b.GetTargetX(i), b.GetTargetY(i)), 1
    Next i
  End If
  
  'draws intermediate positions
  If chkIntermediate.Value = vbChecked Then
    pic.ForeColor = vbCyan
    pic.DrawStyle = 0
    pic.DrawWidth = 1
    
    For i = 1 To b.GetMaxBots
      If b.GetNumIntermediatePos(i) > 0 Then
        For j = 1 To b.GetNumIntermediatePos(i)
          DoEvents
          pic.Circle (b.GetIntermediateX(i, j), b.GetIntermediateY(i, j)), 1
        Next j
      End If
    Next i
  End If
  
  'draws line from X,Y to target
  pic.DrawStyle = 0
  pic.DrawWidth = 1
  If chkDir.Value = vbChecked Then
    pic.ForeColor = vbWhite
    pic.DrawStyle = 4
    For i = 1 To b.GetMaxBots '1 To b.GetMaxBots
      DoEvents
      pic.Line (b.GetX(i), b.GetY(i))-(b.GetTargetX(i), b.GetTargetY(i))
    Next i
  End If
  
  'draws search quad
  If chkQuad.Value = vbChecked Then
    pic.ForeColor = vbCyan
    pic.DrawWidth = 1
    pic.DrawStyle = 6
    g_nRange = GetSearchQuad(b.GetDirection(BOT_MAIN), b.GetVelocity(BOT_MAIN), vsbRange.Value)
    pic.Line (b.GetX(BOT_MAIN) + g_nRange.X_Min, b.GetY(BOT_MAIN) + g_nRange.Y_Max)-(b.GetX(BOT_MAIN) + g_nRange.X_Max, b.GetY(BOT_MAIN) + g_nRange.Y_Min), , B
  End If
  
  'draws collision avoidance direction
  If chkAvoid.Value = vbChecked Then
    pic.ForeColor = vbRed
    For i = 1 To b.GetMaxBots
      If b.GetAvoidStatus(i) = True Then
        pic.Line (b.GetX(i), b.GetY(i))-(b.GetCX(i), b.GetCY(i))
      End If
    Next i
  End If
  
  'draw circles
  If chkCircle.Value = vbChecked Then
    pic.DrawWidth = 1
    pic.DrawStyle = 0
    For i = 1 To b.GetMaxBots
      DoEvents
      If i = BOT_MAIN Then
        pic.ForeColor = vbMagenta

      Else
        pic.ForeColor = vbCyan
      End If
      pic.Circle (b.GetX(i), b.GetY(i)), b.GetDiameter(i) / 2
      pic.CurrentX = b.GetX(i)
      pic.CurrentY = b.GetY(i)
      pic.Print i
      
    Next i
  End If
  
  'draw thick circles of bots inside search quad
  If chkQuad.Value = vbChecked Then
    pic.ForeColor = vbBlue
    pic.DrawWidth = 2
    pic.DrawWidth = 1
    pic.ForeColor = vbWhite
    pic.ForeColor = vbRed
    pic.ForeColor = vbGreen
  End If
End Sub

'places target on picture box
Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim bRet As Boolean
  
  If Button = 1 Then
    bRet = b.SetTargetX(BOT_MAIN, x)
    bRet = b.SetTargetY(BOT_MAIN, y)
  End If
End Sub

Private Sub tmrUpdate_Timer()
  Dim i As Integer
  Dim bRet As Boolean
 
  b.UpdateBots
  For i = 1 To b.GetMaxBots
    DoEvents
    If b.AtTarget(i) = True Then
      bRet = b.SetTargetX(i, GetRandomSingle(10, 90))
      bRet = b.SetTargetY(i, GetRandomSingle(10, 90))
      bRet = b.SetVelocity(i, 0.5)
    End If
  Next i
  
  'draw data
  DrawBots
  
End Sub

Private Sub vsbRange_Change()
  txtRange.Text = vsbRange.Value
End Sub

Private Sub vsbRange_Scroll()
  vsbRange_Change
End Sub
