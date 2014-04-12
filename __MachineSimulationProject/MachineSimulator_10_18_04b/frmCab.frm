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
      LargeChange     =   1000
      Left            =   210
      TabIndex        =   5
      Top             =   7860
      Width           =   10185
   End
   Begin VB.VScrollBar vsb 
      Height          =   7845
      LargeChange     =   1000
      Left            =   10500
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   60
      MouseIcon       =   "frmCab.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   8055
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.PictureBox picTrans 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4110
         Left            =   7770
         Picture         =   "frmCab.frx":0152
         ScaleHeight     =   4110
         ScaleWidth      =   4125
         TabIndex        =   24
         Top             =   4080
         Width           =   4125
      End
      Begin VB.PictureBox picMeter 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   960
         ScaleHeight     =   2775
         ScaleWidth      =   2745
         TabIndex        =   15
         Top             =   3660
         Width           =   2775
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VDC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1830
            TabIndex        =   20
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VAC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1050
            TabIndex        =   19
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   1200
            Width           =   585
         End
         Begin VB.Shape shpDot 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   840
            Shape           =   3  'Circle
            Top             =   1470
            Width           =   225
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   1245
            Left            =   600
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label lblMode 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            Caption         =   "VAC"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2040
            TabIndex        =   17
            Top             =   90
            Width           =   495
         End
         Begin VB.Label lblMeter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            TabIndex        =   16
            Top             =   90
            Width           =   1875
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00008080&
            BackStyle       =   1  'Opaque
            Height          =   855
            Left            =   60
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   2595
         End
      End
      Begin VB.Timer tmrHotspot 
         Interval        =   50
         Left            =   30
         Top             =   30
      End
      Begin VB.PictureBox picF1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3030
         Left            =   2430
         Picture         =   "frmCab.frx":377CC
         ScaleHeight     =   3030
         ScaleWidth      =   7725
         TabIndex        =   4
         Top             =   840
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
         Picture         =   "frmCab.frx":83D86
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
         Picture         =   "frmCab.frx":8CFE8
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
      Begin VB.PictureBox picF2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   5190
         Picture         =   "frmCab.frx":9624A
         ScaleHeight     =   3915
         ScaleWidth      =   1845
         TabIndex        =   21
         Top             =   3870
         Width           =   1845
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "F2B"
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
            Index           =   10
            Left            =   1170
            TabIndex        =   23
            Top             =   1500
            Width           =   315
         End
         Begin VB.Label lblTag 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "F2A"
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
            Index           =   9
            Left            =   240
            TabIndex        =   22
            Top             =   1500
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
  x As Single
  y As Single
End Type

Private Enum METER_MODE
  OFF = 0
  VAC = 1
  VDC = 2
End Enum

Private m_nXOffset As Single
Private m_nYOffset As Single
Private m_uComp(10) As COMPONENT
Private m_nVolt As Single 'value of measured voltage
Private m_nOldX, m_nOldY As Single 'where is the meter when clicked relative to frmCab
Private m_bMoveMeter As Boolean 'move if true
Private m_nMeterMode As METER_MODE
Private m_eVoltType As METER_MODE  'AC or DC

'loads array to store initial positions of all picture boxes (components)
Private Sub Form_Load()
  m_uComp(0).x = picLx.Left
  m_uComp(0).y = picLx.Top
  m_uComp(1).x = pic0Lx.Left
  m_uComp(1).y = pic0Lx.Top
  m_uComp(2).x = picF1.Left
  m_uComp(2).y = picF1.Top
  m_uComp(3).x = picF2.Left
  m_uComp(3).y = picF2.Top
  m_uComp(4).x = picTrans.Left
  m_uComp(4).y = picTrans.Top
  
  m_nMeterMode = OFF
  lblMode.Caption = ""
      
End Sub

Private Sub Form_Resize()
  vsb.Left = frmCab.Width - vsb.Width - 120
  If frmCab.Height > 660 Then vsb.Height = frmCab.Height - 660
  hsb.Top = frmCab.Height - hsb.Height - 400
  hsb.Width = frmCab.Width - 400
  pic.Width = vsb.Left
  If hsb.Top > 0 Then pic.Height = hsb.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.mnuViewElectricalCabinet.Checked = True Then frmMain.mnuViewElectricalCabinet.Checked = False
End Sub

'moves inside electrical cabinet horizontally
Private Sub hsb_Change()
  m_nXOffset = hsb.Value
  
  picLx.Left = m_uComp(0).x - m_nXOffset
  pic0Lx.Left = m_uComp(1).x - m_nXOffset
  picF1.Left = m_uComp(2).x - m_nXOffset
  picF2.Left = m_uComp(3).x - m_nXOffset
  picTrans.Left = m_uComp(4).x - m_nXOffset
End Sub

Private Sub hsb_Scroll()
  hsb_Change
End Sub

Private Sub lblMeter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'frmCab.Caption = frmCab.Left + X & " " & frmCab.Top + Y
  
  
  If Button = 1 Then
    If m_bMoveMeter = False Then
      m_nOldX = x
      m_nOldY = y
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

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  If m_bMoveMeter = True Then
      'm_nOldX = X
      'm_nOldY = Y
      picMeter.Left = x - m_nOldX: picMeter.Top = y - m_nOldY
      m_bMoveMeter = False
     pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
      pic.MousePointer = 99
        
  End If
      
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'lblMeter.Left = X + 2000
  'lblMeter.Top = Y + 500
  
  
End Sub

Private Sub pic0Lx_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(x, y, 165, 900) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(x, y, 165, 1560) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(x, y, 540, 915) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(x, y, 540, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(x, y, 930, 930) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
    If Hotspot(x, y, 930, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
  Else
    picMeter.Left = pic0Lx.Left + x - m_nOldX
    picMeter.Top = pic0Lx.Top + y - m_nOldY
    m_bMoveMeter = False
     pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
      pic.MousePointer = 99
      
  End If
  m_eVoltType = VAC

End Sub

Private Sub pic0Lx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
End Sub

Private Sub picF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(x, y, 1725, 675) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L1))
    If Hotspot(x, y, 1725, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L2))
    If Hotspot(x, y, 1725, 2430) Then m_nVolt = g_nPhaseV * -CInt(v(V_0L3))
    If Hotspot(x, y, 5790, 675) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L1))
    If Hotspot(x, y, 5790, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L2))
    If Hotspot(x, y, 5790, 2430) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L3))
  Else
    picMeter.Left = picF1.Left + x - m_nOldX
    picMeter.Top = picF1.Top + y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
      
  End If
  m_eVoltType = VAC
End Sub

Private Sub picF1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
End Sub

'Fuses F2 to transformer primary
Private Sub picF2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(x, y, 460, 675) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L3))
    If Hotspot(x, y, 1335, 690) Then m_nVolt = g_nPhaseV * -CInt(v(V_1L2))
    If Hotspot(x, y, 460, 3435) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H1))
    If Hotspot(x, y, 1350, 3435) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H4))
  Else
    picMeter.Left = picF2.Left + x - m_nOldX
    picMeter.Top = picF2.Top + y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
  End If
  m_eVoltType = VAC
End Sub

Private Sub picF2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
End Sub

Private Sub picLx_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  If m_bMoveMeter = False Then
    If Hotspot(x, y, 165, 900) Then m_nVolt = g_nPhaseV * -CInt(v(V_L1))
    If Hotspot(x, y, 165, 1560) Then m_nVolt = g_nPhaseV * -CInt(v(V_L1))
    If Hotspot(x, y, 540, 915) Then m_nVolt = g_nPhaseV * -CInt(v(V_L2))
    If Hotspot(x, y, 540, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_L2))
    If Hotspot(x, y, 930, 930) Then m_nVolt = g_nPhaseV * -CInt(v(V_L3))
    If Hotspot(x, y, 930, 1575) Then m_nVolt = g_nPhaseV * -CInt(v(V_L3))
  Else
    picMeter.Left = picLx.Left + x - m_nOldX
    picMeter.Top = picLx.Top + y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
    
  End If
  m_eVoltType = VAC
End Sub

Private Sub picLx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
End Sub

Private Sub picMeter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  frmMain.Caption = x & " " & y
  If Sqr((x - 960) ^ 2 + (1590 - y) ^ 2) < 200 Then
    m_nMeterMode = OFF
    lblMode.Caption = ""
    shpDot.Left = 840
    shpDot.Top = 1470
    lblMeter.Caption = ""
    
  End If
  If Sqr((x - 1350) ^ 2 + (1395 - y) ^ 2) < 200 Then
    m_nMeterMode = VAC
    lblMode.Caption = "VAC"
    shpDot.Left = 1230
    shpDot.Top = 1275
    lblMeter.Caption = "0"
  
  End If
  If Sqr((x - 1725) ^ 2 + (1590 - y) ^ 2) < 200 Then
    m_nMeterMode = VDC
    lblMode.Caption = "VDC"
    shpDot.Left = 1605
    shpDot.Top = 1470
    lblMeter.Caption = "0"
  End If
  m_nVolt = 0
End Sub

Private Sub picTrans_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
  frmMain.Caption = x & "  " & y
  If m_bMoveMeter = False Then
    If HotspotBig(x, y, 570, 660) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H1))
    If HotspotBig(x, y, 1365, 660) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H2))
    If HotspotBig(x, y, 2775, 660) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H3))
    If HotspotBig(x, y, 3525, 660) Then m_nVolt = g_nPhaseV * -CInt(v(V_T1_H4))
    If HotspotBig(x, y, 570, 3450) Then m_nVolt = g_nHotV * -CInt(v(V_T1_X1))
    If HotspotBig(x, y, 3525, 3450) Then m_nVolt = 0 * -CInt(v(V_T1_X2))
  
  Else
    picMeter.Left = picTrans.Left + x - m_nOldX
    picMeter.Top = picTrans.Top + y - m_nOldY
    m_bMoveMeter = False
    pic.MouseIcon = LoadPicture(App.Path & "\images\meterprobered.cur")
    pic.MousePointer = 99
  End If
  m_eVoltType = VAC
End Sub

Private Sub picTrans_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  m_nVolt = 0
End Sub

Private Sub tmrHotspot_Timer()
  'frmCab.Caption = m_nVolt
  If m_nMeterMode = OFF Then
    lblMeter.Caption = ""
  ElseIf m_nMeterMode = VAC Then
    If m_eVoltType = VAC Then
      lblMeter.Caption = CStr(m_nVolt)
      
    End If
  ElseIf m_nMeterMode = VDC Then
    If m_eVoltType = VDC Then
      lblMeter.Caption = CStr(m_nVolt)
    End If
  End If
  lblMeter.ZOrder
  
End Sub

'moves inside electrical cabinet vertically
Private Sub vsb_Change()
  
  m_nYOffset = vsb.Value
  
  picLx.Top = m_uComp(0).y - m_nYOffset
  pic0Lx.Top = m_uComp(1).y - m_nYOffset
  picF1.Top = m_uComp(2).y - m_nYOffset
  picF2.Top = m_uComp(3).y - m_nYOffset
  picTrans.Top = m_uComp(4).y - m_nYOffset
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub

'returns true if mouseclick is inside circle of diameter 120
Private Function Hotspot(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Boolean
  Hotspot = False
  If Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2) < 120 Then Hotspot = True
End Function

'returns true if mouseclick is inside circle of diameter 240
Private Function HotspotBig(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Boolean
  HotspotBig = False
  If Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2) < 240 Then HotspotBig = True
End Function

