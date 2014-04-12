VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscope v0.1 - Written by C. Bolin, October 2004"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbTime 
      Height          =   375
      LargeChange     =   5
      Left            =   6900
      Max             =   22
      TabIndex        =   17
      Top             =   2700
      Value           =   18
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Probe "
      Height          =   1035
      Left            =   5880
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
      Begin VB.OptionButton opt10X 
         Caption         =   "10x"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   600
         Width           =   795
      End
      Begin VB.OptionButton opt1X 
         Caption         =   "1x"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.HScrollBar hsbFreq 
      Height          =   375
      Left            =   1560
      Max             =   70
      Min             =   1
      TabIndex        =   9
      Top             =   5940
      Value           =   1
      Width           =   1215
   End
   Begin VB.VScrollBar vsbVDiv 
      Height          =   1275
      Left            =   7260
      Max             =   0
      Min             =   9
      TabIndex        =   7
      Top             =   720
      Value           =   7
      Width           =   435
   End
   Begin VB.VScrollBar vsbVPk 
      Height          =   1155
      LargeChange     =   10
      Left            =   600
      Max             =   0
      Min             =   1000
      TabIndex        =   4
      Top             =   5880
      Value           =   20
      Width           =   435
   End
   Begin VB.CheckBox chkIlluminate 
      Caption         =   "Illuminate Grid"
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   1395
   End
   Begin VB.VScrollBar vsb 
      Height          =   1275
      LargeChange     =   100
      Left            =   6060
      Max             =   -400
      Min             =   400
      TabIndex        =   1
      Top             =   720
      Width           =   435
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   4000
      Left            =   360
      ScaleHeight     =   -8
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   4
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   240
      Width           =   5000
      Begin VB.Line linRef 
         BorderColor     =   &H0000C000&
         X1              =   -4.279
         X2              =   1.366
         Y1              =   -0.075
         Y2              =   -0.075
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   6960
      TabIndex        =   16
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Time/Div"
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
      Left            =   6900
      TabIndex        =   15
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Function Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5220
      Width           =   2835
   End
   Begin VB.Label lblFreq 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   1620
      TabIndex        =   10
      Top             =   5640
      Width           =   1035
   End
   Begin VB.Label lblVoltDiv 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   6960
      TabIndex        =   8
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Volts/Div"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label lblVoltPeak 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   5640
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   3660
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Vertical Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5760
      TabIndex        =   2
      Top             =   60
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   4515
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'module variables
Private m_nRef As Single 'position of reference line
Private m_nVoltDiv(9) As Single 'stores volts/div possible values
Private m_nTimeDiv(22) As Single 'stores time/div values
Private m_nProbe As Single

'************************************************************ Form_Load
'initializes program
Private Sub Form_Load()

  'load variables
  m_nVoltDiv(9) = 5
  m_nVoltDiv(8) = 2
  m_nVoltDiv(7) = 1
  m_nVoltDiv(6) = 0.5
  m_nVoltDiv(5) = 0.2
  m_nVoltDiv(4) = 0.1
  m_nVoltDiv(3) = 0.05
  m_nVoltDiv(2) = 0.02
  m_nVoltDiv(1) = 0.01
  m_nVoltDiv(0) = 0.005
  
  m_nTimeDiv(0) = 0.0000001
  m_nTimeDiv(1) = 0.0000002
  m_nTimeDiv(2) = 0.0000005
  m_nTimeDiv(3) = 0.000001
  m_nTimeDiv(4) = 0.000002
  m_nTimeDiv(5) = 0.000005
  m_nTimeDiv(6) = 0.00001
  m_nTimeDiv(7) = 0.00002
  m_nTimeDiv(8) = 0.00005
  m_nTimeDiv(9) = 0.0001
  m_nTimeDiv(10) = 0.0002
  m_nTimeDiv(11) = 0.0005
  m_nTimeDiv(12) = 0.001
  m_nTimeDiv(13) = 0.002
  m_nTimeDiv(14) = 0.005
  m_nTimeDiv(15) = 0.01
  m_nTimeDiv(16) = 0.02
  m_nTimeDiv(17) = 0.05
  m_nTimeDiv(18) = 0.1
  m_nTimeDiv(19) = 0.2
  m_nTimeDiv(20) = 0.5
  m_nTimeDiv(21) = 1
  m_nTimeDiv(22) = 2
  
  m_nProbe = 1    'scope probe
  m_nRef = 0
  
  vsbVPk_Change   'function generator
  hsbFreq_Change
  hsbTime_Change  'scope settings
  vsbVDiv_Change
  
  
  
End Sub

'*********************************************************** chkIlluminate_Click
'occurs when check box value is changed
Private Sub chkIlluminate_Click()
  UpdateScope
End Sub

'**********************************************************  UpdateScope
' Redraws everything
Private Sub UpdateScope()
  pic.Cls
  linRef.X1 = -5
  linRef.X2 = 5
  linRef.Y1 = vsb.Value / 100
  linRef.Y2 = vsb.Value / 100
  m_nRef = vsb.Value / 100

  DrawScopeGrid
  DrawWaveform CSng(vsbVPk.Value / 10), CSng(hsbFreq.Value)
  

End Sub

'*********************************************************** DrawScopeGrid
Private Sub DrawScopeGrid()
  Dim i As Integer
  
  If chkIlluminate.Value = vbChecked Then
    pic.ForeColor = RGB(200, 200, 0)
  Else
    pic.ForeColor = RGB(0, 100, 0)
  End If
  
  For i = -4 To 4
    pic.Line (i, 4)-(i, -4) 'vertical lines
  Next i
  
  For i = -3 To 3
    pic.Line (-5, i)-(5, i) 'horizontal lines
  Next i
  
End Sub


Private Sub hsbTime_Change()
  lblTime.Caption = CStr(m_nTimeDiv(hsbTime.Value)) & " S"
  UpdateScope
End Sub

Private Sub hsbTime_Scroll()
  hsbTime_Change
End Sub

Private Sub opt10X_Click()
  m_nProbe = 10
  UpdateScope
End Sub

Private Sub opt1X_Click()
  m_nProbe = 1
  UpdateScope
End Sub

'********************************************************** vsb_Change
'moves vertical reference for scope
Private Sub vsb_Change()
  linRef.X1 = -5
  linRef.X2 = 5
  linRef.Y1 = vsb.Value / 100
  linRef.Y2 = vsb.Value / 100
  m_nRef = vsb.Value / 100
  UpdateScope
End Sub

'********************************************************** DrawWaveform
'Requires amplitude of sine wave (Volts) and frequency (Hz)
Private Sub DrawWaveform(nAmp As Single, nFreq As Single)
  If nFreq = 0 Then Exit Sub
  Dim i As Long
  Dim x, y As Single
  Dim oldx, oldy As Single
  Dim bFirst As Boolean
  bFirst = True
  
  '* 10 / m_nTimeDiv(hsbTime.Value)
  
  pic.ForeColor = vbGreen
  'For i = 0 To (360 * CInt(nFreq)) - 1
  For i = 0 To (360 * CInt(nFreq)) - 1
  
    'x = CSng(i) * (m_nTimeDiv(hsbTime.Value) / 10) / (36 * nFreq)
     x = CSng(i) / (36 * nFreq)
    y = m_nRef + Sin(3.14159 / 180 * CSng(i)) * (nAmp / m_nVoltDiv(vsbVDiv.Value) / m_nProbe)
    If bFirst = False Then
      pic.Line (-5 + x, y)-(-5 + oldx, oldy)
    End If
    oldx = x: oldy = y
    bFirst = False
  Next i
  
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub

Private Sub vsbVDiv_Change()
  lblVoltDiv.Caption = CStr(m_nVoltDiv(vsbVDiv.Value)) & " V/Div"
  UpdateScope
End Sub

Private Sub vsbVDiv_Scroll()
  vsbVDiv_Change
End Sub

'****************************************** F U N C T I O N   G E N E R A T O R

'********************************************************** vsbVPk_Change
'Changes applied voltage
Private Sub vsbVPk_Change()
  lblVoltPeak.Caption = "V-PEAK: " & CStr(vsbVPk.Value / 10) & " V"
  UpdateScope
End Sub

Private Sub vsbVPk_Scroll()
  vsbVPk_Change
End Sub

'************************************************************ hsbFreq_Change
Private Sub hsbFreq_Change()
  lblFreq = "Freq: " & CStr(hsbFreq.Value) & " Hz"
  UpdateScope
End Sub

Private Sub hsbFreq_Scroll()
  hsbFreq_Change
End Sub

