VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Designer"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbView 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8220
      TabIndex        =   7
      Top             =   660
      Width           =   1815
   End
   Begin VB.TextBox txtView 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   6
      Top             =   360
      Width           =   1755
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Play"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8100
      TabIndex        =   5
      Top             =   4500
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   8100
      TabIndex        =   4
      Top             =   6300
      Width           =   1335
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Keep Go To"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8100
      TabIndex        =   3
      Top             =   5100
      Width           =   1335
   End
   Begin VB.TextBox txtPlay 
      Height          =   1755
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   8040
      Width           =   7995
   End
   Begin VB.CommandButton cmdKeepLineup 
      Caption         =   "Keep Line Up"
      Height          =   315
      Left            =   8100
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   8000
      Left            =   0
      ScaleHeight     =   -60
      ScaleMode       =   0  'User
      ScaleTop        =   60
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8000
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   1
         Left            =   1860
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape shpPlayer 
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   0
         Left            =   2400
         Top             =   5880
         Width           =   195
      End
      Begin VB.Line LOS 
         BorderColor     =   &H0000C0C0&
         X1              =   1.361
         X2              =   57.618
         Y1              =   17.353
         Y2              =   17.353
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variable declaration
Private m_bSelect As Boolean 'true means player selected
Private m_nPlayer As Integer '0 to 10..player number selected
Private m_nLOS As Single 'line of scrimmage
Private m_nRelX As Single 'offset from top-left of player
Private m_nRelY As Single
Private m_nGotoCt As Integer 'counts goto's
Private m_nStepLimit As Single 'ensure no overtravel between waypoints
Private m_nStartX As Single
Private m_nStartY As Single
Private m_nPlayNum As Integer 'play number


Private Sub cmdAdd_Click()
  m_nPlayNum = m_nPlayNum + 1
  txtPlay.Text = "Play:" & CStr(m_nPlayNum) & vbCrLf & txtPlay.Text
  ReDim Preserve g_sPlay(m_nPlayNum)
  g_sPlay(m_nPlayNum) = txtPlay.Text
  
  cmdAdd.Enabled = False
  cmdGoto.Enabled = False
  txtView.Enabled = True
  hsbView.Enabled = True
  hsbView.Max = m_nPlayNum
  hsbView.Min = 1
  hsbView_Change
End Sub

Private Sub cmdClear_Click()
  txtPlay.Text = ""
  pic.Cls
  LineUpPlayers
  cmdKeepLineup.Enabled = True
  cmdGoto.Enabled = False
  cmdAdd.Enabled = False
  txtView.Text = ""
  txtView.Enabled = False
  hsbView.Enabled = False
End Sub

Private Sub cmdGoto_Click()
  If Len(txtPlay.Text) < 1 Then Exit Sub
  
  Dim i As Integer
  Dim sGoto As String
  
  m_nGotoCt = m_nGotoCt + 1
  
  sGoto = "Goto:" & CStr(m_nGotoCt) & ","
  
  For i = 0 To 10
    sGoto = sGoto & CStr(Format(shpPlayer(i).Left, "#.#")) & "," & CStr(Format(shpPlayer(i).Top, "#.#")) & ","
    pic.ForeColor = shpPlayer(i).BackColor
    pic.Circle (shpPlayer(i).Left, shpPlayer(i).Top), 0.5
    
  Next i
  sGoto = Left(sGoto, Len(sGoto) - 1)
  txtPlay.Text = txtPlay.Text & vbCrLf & sGoto
  
  'cmdKeepLineup.Enabled = True
  'cmdGoto.Enabled = true
  cmdAdd.Enabled = True

End Sub

Private Sub cmdKeepLineup_Click()
  If Len(txtPlay.Text) > 0 Then Exit Sub
  Dim i As Integer
  Dim sLineUp As String
  Dim sPositions As String
    
  sLineUp = "LineUp:"
  sPositions = "Positions:" & "WR1,TE,QB,LT,LG,C,RG,RT,WR2,RB1,RB2"
  For i = 0 To 10
    sPositions = sPositions & ""
    sLineUp = sLineUp & CStr(Format(shpPlayer(i).Left, "#.#")) & "," & CStr(Format(shpPlayer(i).Top, "#.#")) & ","
    pic.ForeColor = shpPlayer(i).BackColor
  pic.Line (shpPlayer(i).Left, shpPlayer(i).Top)-(shpPlayer(i).Left + shpPlayer(i).Width, shpPlayer(i).Top - shpPlayer(i).Height), , B
  Next i
  sLineUp = Left(sLineUp, Len(sLineUp) - 1)
  txtPlay.Text = sPositions & vbCrLf & sLineUp
  m_nGotoCt = 0
  
  cmdKeepLineup.Enabled = False
  cmdGoto.Enabled = True
  'cmdAdd.Enabled = False

End Sub



Private Sub Form_Load()
  Dim i As Integer
    
  'initialize graphics
  LOS.X1 = 0: LOS.X2 = pic.Width
  LOS.Y1 = 20: LOS.Y2 = LOS.Y1
  m_nLOS = 20
  LineUpPlayers
  ReDim g_sPlay(0)
  
End Sub

Private Sub LineUpPlayers()
  Dim i As Integer
  
  'position players
  For i = 0 To 10
    shpPlayer(i).Width = 1
    shpPlayer(i).Height = 1
    Select Case i
      Case P_LT To P_RT
        shpPlayer(i).BackColor = vbBlack
        shpPlayer(i).Left = 22.5 + i * 1.5
        shpPlayer(i).Top = m_nLOS - 1
      Case P_WR1 To P_TE, P_WR2
        shpPlayer(i).BackColor = vbRed
        If i = P_WR1 Then shpPlayer(i).Left = 20
        If i = P_WR2 Then shpPlayer(i).Left = 40
        If i = P_TE Then shpPlayer(i).Left = 24
        shpPlayer(i).Top = m_nLOS - 1
      Case P_RB1 To P_RB2
        shpPlayer(i).BackColor = vbCyan
        If i = P_RB1 Then shpPlayer(i).Left = 29
        If i = P_RB2 Then shpPlayer(i).Left = 31
        shpPlayer(i).Top = m_nLOS - 4
      Case P_QB
        shpPlayer(i).BackColor = vbYellow
        shpPlayer(i).Left = 30
        shpPlayer(i).Top = m_nLOS - 2
    End Select
    
  Next i

End Sub

Private Sub DrawPlay()
  Dim sIn As String
  Dim i, j As Integer
  Dim sLines() As String
  Dim sCoord() As String
  
  pic.Cls
    
  sIn = g_sPlay(hsbView.Value)
  sLines() = Split(sIn, vbCrLf)
  For i = 0 To UBound(sLines)
    If Left(sLines(i), 7) = "LineUp:" Then
      sCoord() = Split(sLines(i), ",")
      For j = 1 To UBound(sCoord) - 1 Step 2
        shpPlayer(j \ 2).Left = CSng(sCoord(j))
        shpPlayer(j \ 2).Top = CSng(sCoord(j + 1))
        'MsgBox j \ 2
      Next j
    End If
  Next i
  
  
  
End Sub

Private Sub hsbView_Change()
  txtPlay.Text = g_sPlay(hsbView.Value)
  txtView.Text = hsbView.Value
  DrawPlay
  
End Sub

Private Sub hsbView_Scroll()
  hsbView_Change
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  
  'select player
  If Button = vbLeftButton Then
    m_bSelect = False

    For i = 0 To 10
      If X > shpPlayer(i).Left And X < shpPlayer(i).Left + shpPlayer(i).Width Then
        If Y < shpPlayer(i).Top And Y > shpPlayer(i).Top - shpPlayer(i).Height Then
          m_bSelect = True
          m_nPlayer = i
          m_nRelX = shpPlayer(i).Left - X
          m_nRelY = shpPlayer(i).Top - Y
          
          'determines max distance to travel
          If shpPlayer(i).BackColor = vbRed Then
            m_nStepLimit = 20
          ElseIf shpPlayer(i).BackColor = vbBlack Then
            m_nStepLimit = 5
          ElseIf shpPlayer(i).BackColor = vbCyan Then
            m_nStepLimit = 15
          ElseIf shpPlayer(i).BackColor = vbYellow Then
            m_nStepLimit = 10
          End If
          m_nStartX = X
          m_nStartY = Y
          Exit For
        End If
      End If
    Next i
  End If
  
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim nDist As Single
  
  frmMain.Caption = Format(X, "##.#") & " " & Format(Y, "##.#")
  
 'move player
  If Button = vbLeftButton And m_bSelect = True Then
    If cmdKeepLineup.Enabled = True Then  'setting LOS lineup
        shpPlayer(m_nPlayer).Left = X + m_nRelX
        shpPlayer(m_nPlayer).Top = Y + m_nRelY
    Else
      nDist = Sqr((X - m_nStartX) ^ 2 + (Y - m_nStartY) ^ 2)
      If nDist < m_nStepLimit Then
        shpPlayer(m_nPlayer).Left = X + m_nRelX
        shpPlayer(m_nPlayer).Top = Y + m_nRelY
        pic.PSet (X + m_nRelX, Y + m_nRelY), shpPlayer(m_nPlayer).BackColor
      End If
    End If
    
  End If
  
End Sub

