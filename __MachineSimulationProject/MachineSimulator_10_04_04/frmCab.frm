VERSION 5.00
Begin VB.Form frmCab 
   Caption         =   "Electrical Cabinet (Internal)"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   8100
   ScaleWidth      =   10485
   Begin VB.HScrollBar hsb 
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   7830
      Width           =   10185
   End
   Begin VB.VScrollBar vsb 
      Height          =   7845
      LargeChange     =   1000
      Left            =   10230
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   7845
      Left            =   30
      MouseIcon       =   "frmCab.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   7845
      ScaleWidth      =   10245
      TabIndex        =   0
      Top             =   0
      Width           =   10245
      Begin VB.PictureBox picMeter 
         Height          =   795
         Left            =   4770
         ScaleHeight     =   735
         ScaleWidth      =   2085
         TabIndex        =   15
         Top             =   2400
         Width           =   2145
         Begin VB.Label lblMeter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Digital Display"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   2085
         End
      End
      Begin VB.Timer tmrHotspot 
         Interval        =   50
         Left            =   6180
         Top             =   6300
      End
      Begin VB.PictureBox picF1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3030
         Left            =   2430
         Picture         =   "frmCab.frx":0152
         ScaleHeight     =   3030
         ScaleWidth      =   7725
         TabIndex        =   4
         Top             =   870
         Width           =   7725
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "F1C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   3630
            TabIndex        =   14
            Top             =   1920
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "F1B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   3630
            TabIndex        =   13
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "F1A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   3630
            TabIndex        =   12
            Top             =   150
            Width           =   315
         End
      End
      Begin VB.PictureBox pic0Lx 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   1290
         Picture         =   "frmCab.frx":4C70C
         ScaleHeight     =   2505
         ScaleWidth      =   1110
         TabIndex        =   3
         Top             =   870
         Width           =   1110
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0L3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   11
            Top             =   570
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0L2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   10
            Top             =   570
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0L1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   9
            Top             =   570
            Width           =   315
         End
      End
      Begin VB.PictureBox picLx 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   120
         Picture         =   "frmCab.frx":5596E
         ScaleHeight     =   2505
         ScaleWidth      =   1110
         TabIndex        =   2
         Top             =   870
         Width           =   1110
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "L3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   8
            Top             =   570
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "L2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   7
            Top             =   570
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "L1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   6
            Top             =   570
            Width           =   315
         End
      End
   End
End
Attribute VB_Name = "frmCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type COMPONENT
  X As Single
  Y As Single
End Type
Private m_nXOffset As Single
Private m_nYOffset As Single
Private m_uComp(10) As COMPONENT
Private m_nVolt As Single 'value of measured voltage
Private m_nOldX, m_nOldY As Single 'where is the meter when clicked relative to frmCab
Private m_bMoveMeter As Boolean 'move if true

'loads array to store initial positions of all picture boxes (components)
Private Sub Form_Load()
  m_uComp(0).X = picLx.Left
  m_uComp(0).Y = picLx.Top
  m_uComp(1).X = pic0Lx.Left
  m_uComp(1).Y = pic0Lx.Top
  m_uComp(2).X = picF1.Left
  m_uComp(2).Y = picF1.Top
  
End Sub

Private Sub Form_Resize()
  vsb.Left = frmCab.Width - vsb.Width - 120
  If frmCab.Height > 660 Then vsb.Height = frmCab.Height - 660
  hsb.Top = frmCab.Height - hsb.Height - 400
  hsb.Width = frmCab.Width - 400
  pic.Width = vsb.Left
  If hsb.Top > 0 Then pic.Height = hsb.Top
End Sub

'moves inside electrical cabinet horizontally
Private Sub hsb_Change()
  m_nXOffset = hsb.Value
  
  picLx.Left = m_uComp(0).X - m_nXOffset
  pic0Lx.Left = m_uComp(1).X - m_nXOffset
  picF1.Left = m_uComp(2).X - m_nXOffset
End Sub

Private Sub hsb_Scroll()
  hsb_Change
End Sub

Private Sub lblMeter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'frmCab.Caption = frmCab.Left + X & " " & frmCab.Top + Y
  
  
  If Button = 1 Then
    If m_bMoveMeter = False Then
      m_nOldX = X
      m_nOldY = Y
      m_bMoveMeter = True
      pic.MouseIcon = LoadPicture(App.Path & "\images\movemeter.cur")
      pic.MousePointer = 99
      
    Else
      m_bMoveMeter = False
      pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
      pic.MousePointer = 99
      
    End If
  End If
  
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_nVolt = 0
  If m_bMoveMeter = True Then
      'm_nOldX = X
      'm_nOldY = Y
      picMeter.Left = X - m_nOldX: picMeter.Top = Y - m_nOldY
      m_bMoveMeter = False
     pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
      pic.MousePointer = 99
        
  End If
      
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'lblMeter.Left = X + 2000
  'lblMeter.Top = Y + 500
  
  
End Sub

Private Sub pic0Lx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(X, Y, 165, 900) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(X, Y, 165, 1560) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(X, Y, 540, 915) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(X, Y, 540, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(X, Y, 930, 930) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
    If Hotspot(X, Y, 930, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
  Else
    picMeter.Left = pic0Lx.Left + X - m_nOldX
    picMeter.Top = pic0Lx.Top + Y - m_nOldY
    m_bMoveMeter = False
     pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
      pic.MousePointer = 99
      
  End If


End Sub

Private Sub picF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(X, Y, 1725, 675) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(X, Y, 1725, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(X, Y, 1725, 2430) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
    If Hotspot(X, Y, 5790, 675) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L1))
    If Hotspot(X, Y, 5790, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L2))
    If Hotspot(X, Y, 5790, 2430) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L3))
  Else
    picMeter.Left = picF1.Left + X - m_nOldX
    picMeter.Top = picF1.Top + Y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
      
  End If

End Sub

Private Sub picLx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_nVolt = 0
  
  If m_bMoveMeter = False Then
    If Hotspot(X, Y, 165, 900) Then m_nVolt = g_nPhaseV * -CInt(v(V_L1))
    If Hotspot(X, Y, 165, 1560) Then m_nVolt = g_nPhaseV * -CInt(v(V_L1))
    If Hotspot(X, Y, 540, 915) Then m_nVolt = g_nPhaseV * -CInt(v(V_L2))
    If Hotspot(X, Y, 540, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_L2))
    If Hotspot(X, Y, 930, 930) Then m_nVolt = g_nPhaseV * -CInt(v(V_L3))
    If Hotspot(X, Y, 930, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_L3))
  Else
    picMeter.Left = picLx.Left + X - m_nOldX
    picMeter.Top = picLx.Top + Y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
  
  End If
End Sub

Private Sub tmrHotspot_Timer()
  'frmCab.Caption = m_nVolt
  lblMeter.Caption = CStr(m_nVolt) & " V"
  lblMeter.ZOrder
  
End Sub

'moves inside electrical cabinet vertically
Private Sub vsb_Change()
  
  m_nYOffset = vsb.Value
  
  picLx.Top = m_uComp(0).Y - m_nYOffset
  pic0Lx.Top = m_uComp(1).Y - m_nYOffset
  picF1.Top = m_uComp(2).Y - m_nYOffset
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub

'returns true if mouseclick is inside circle of diameter 160
Private Function Hotspot(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Boolean
  Hotspot = False
  If Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2) < 120 Then Hotspot = True
End Function
