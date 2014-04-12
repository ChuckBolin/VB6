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
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   8340
      Top             =   7800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   8160
      Width           =   7995
      Begin VB.TextBox txtData 
         Height          =   1215
         Left            =   4380
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.VScrollBar vsbRange 
         Height          =   795
         Left            =   3420
         Max             =   5
         Min             =   100
         SmallChange     =   10
         TabIndex        =   7
         Top             =   660
         Value           =   75
         Width           =   315
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Left            =   3300
         TabIndex        =   6
         Top             =   300
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
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkDir 
         Caption         =   "Direction Vector"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Search Range:"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   300
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

Private Sub Form_Click()
  If tmrUpdate.Interval = 50 Then
  tmrUpdate.Interval = 500
  Else
  tmrUpdate.Interval = 50
  End If
  
End Sub

'initialization
Private Sub Form_Load()
  Randomize Timer
  vsbRange_Change
  LoadBotData 'load all bot information
  DrawBots
End Sub

'draws all bots
Private Sub DrawBots()
  Dim i As Integer
  Dim nPair As PAIR
  
  pic.Cls
  
  'sets color points based upon bot( )
  pic.DrawWidth = 2
  For i = 0 To MAX_BOTS
    If i = BOT_MAIN Then
      pic.ForeColor = vbYellow
    Else
      pic.ForeColor = vbGreen
    End If
    pic.PSet (bot(i).x, bot(i).y)
  Next i
  
  'draws target X and Y
  pic.ForeColor = vbRed
  pic.DrawStyle = 0
  pic.DrawWidth = 1
  pic.Circle (bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY), 1
  
  'draws line from X,Y to target
  If chkDir.Value = vbChecked Then
    pic.ForeColor = vbWhite
    pic.DrawStyle = 4
    pic.Line (bot(BOT_MAIN).x, bot(BOT_MAIN).y)-(bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)
  End If
  
  'draws search quad
  If chkQuad.Value = vbChecked Then
    pic.ForeColor = vbCyan
    pic.DrawWidth = 1
    pic.DrawStyle = 6
    nRange = GetSearchQuad(bot(BOT_MAIN).dir, bot(BOT_MAIN).vel, vsbRange.Value)
    pic.Line (bot(BOT_MAIN).x + nRange.X_Min, bot(BOT_MAIN).y + nRange.Y_Max)-(bot(BOT_MAIN).x + nRange.X_Max, bot(BOT_MAIN).y + nRange.Y_Min), , B
  End If
  
  'draw circles
  If chkCircle.Value = vbChecked Then
    pic.ForeColor = vbGreen
    pic.DrawWidth = 1
    pic.DrawStyle = 0
    For i = 0 To MAX_BOTS
      pic.Circle (bot(i).x, bot(i).y), bot(i).Diameter / 2
    Next i
  End If
  
  'draw thick circles of bots inside search quad
  If chkQuad.Value = vbChecked Then
    If bot(BOT_MAIN).CloseCount > 0 Then
      For i = 0 To 0 ' UBound(bot(BOT_MAIN).InRange) - 1
        If bot(BOT_MAIN).InRange(i) > 0 Then
          pic.ForeColor = vbBlue
          pic.DrawWidth = 2
          pic.Circle (bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y), bot(bot(BOT_MAIN).InRange(i)).Diameter / 2
          pic.DrawWidth = 1
          pic.ForeColor = vbWhite
          pic.CurrentX = bot(bot(BOT_MAIN).InRange(i)).x + 2
          pic.CurrentY = bot(bot(BOT_MAIN).InRange(i)).y + 2
          pic.Print bot(BOT_MAIN).InRange(i)
          If GetTargetDistance2D(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y) < GetTargetDistance2D(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY) Then
            'pic.ForeColor = vbYellow
            'pic.Line (bot(BOT_MAIN).x, bot(BOT_MAIN).y)-(bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y)
            nPair = GetCCW(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y)
            pic.ForeColor = vbRed
            pic.Line (bot(BOT_MAIN).x, bot(BOT_MAIN).y)-(nPair.x, nPair.y)
            pic.Line (bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)-(nPair.x, nPair.y)
            'pic.Line (bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y)-(nPair.x, nPair.y)
            nPair = GetCW(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y)
            pic.ForeColor = vbGreen
            pic.Line (bot(BOT_MAIN).x, bot(BOT_MAIN).y)-(nPair.x, nPair.y)
            pic.Line (bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)-(nPair.x, nPair.y)
            'pic.Line (bot(bot(BOT_MAIN).InRange(i)).x, bot(bot(BOT_MAIN).InRange(i)).y)-(nPair.x, nPair.y)
          End If
        End If
      Next i
    End If
  End If
End Sub

'


'places target on picture box
Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    bot(BOT_MAIN).TargetX = x: bot(BOT_MAIN).TargetY = y
    bot(BOT_MAIN).TargetFound = False
    bot(BOT_MAIN).Obstacle = False
       bot(BOT_MAIN).cx = bot(BOT_MAIN).TargetX
            bot(BOT_MAIN).cy = bot(BOT_MAIN).TargetY
    frmMain.Caption = x & " : " & y
    txtData.Text = ""
  txtData.Text = bot(BOT_MAIN).TargetFound & " " & bot(BOT_MAIN).Obstacle
   ' update bot data
  'UpdateBots
   
  'draw data
  'DrawBots
  End If
End Sub

Private Sub tmrUpdate_Timer()
  Dim i As Integer
  
  'update bot data
  UpdateBots
  txtData.Text = bot(BOT_MAIN).TargetFound & " " & bot(BOT_MAIN).Obstacle & " " & bot(BOT_MAIN).InRange(0)
  
  'draw data
  DrawBots
  
  'display data in txtData
  
  
  'If Not IsEmpty(UBound(bot(BOT_MAIN).InRange)) Then
  '
  '  For i = 1 To UBound(bot(BOT_MAIN).InRange)
  '    txtData.Text = txtData.Text & bot(i).ID & ", " & bot(i).X & ", " & bot(i).Y & vbCrLf
  '  Next i
  'End If
  
End Sub

Private Sub vsbRange_Change()
  txtRange.Text = vsbRange.Value
End Sub

Private Sub vsbRange_Scroll()
  vsbRange_Change
End Sub
