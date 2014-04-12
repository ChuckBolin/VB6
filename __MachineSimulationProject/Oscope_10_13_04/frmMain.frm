VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscope v0.1 - Written by C. Bolin, October 2004"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Solution"
      Height          =   375
      Left            =   7530
      TabIndex        =   23
      Top             =   4740
      Width           =   1665
   End
   Begin VB.CheckBox chkPractice 
      Caption         =   "Generate Problem"
      Height          =   345
      Left            =   5610
      TabIndex        =   22
      Top             =   4770
      Width           =   1755
   End
   Begin VB.HScrollBar hsbIntensity 
      Height          =   315
      LargeChange     =   20
      Left            =   2070
      Max             =   127
      TabIndex        =   20
      Top             =   4800
      Value           =   63
      Width           =   1155
   End
   Begin VB.HScrollBar hsbXRef 
      Height          =   375
      LargeChange     =   100
      Left            =   5670
      Max             =   -500
      Min             =   500
      TabIndex        =   19
      Top             =   2700
      Value           =   -500
      Width           =   1245
   End
   Begin VB.HScrollBar hsbTime 
      Height          =   375
      LargeChange     =   5
      Left            =   7200
      Max             =   22
      TabIndex        =   17
      Top             =   2730
      Value           =   18
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Probe "
      Height          =   1035
      Left            =   5790
      TabIndex        =   12
      Top             =   3390
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
      Max             =   20
      Min             =   1
      TabIndex        =   9
      Top             =   6090
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
      Top             =   6030
      Value           =   20
      Width           =   435
   End
   Begin VB.CheckBox chkIlluminate 
      Caption         =   "Illuminate Grid"
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Top             =   4560
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
   Begin VB.Label lblVavg 
      Height          =   285
      Left            =   3630
      TabIndex        =   27
      Top             =   6660
      Width           =   1515
   End
   Begin VB.Label lblVrms 
      Height          =   285
      Left            =   3630
      TabIndex        =   26
      Top             =   6270
      Width           =   1515
   End
   Begin VB.Label lblVpkpk 
      Height          =   285
      Left            =   3630
      TabIndex        =   25
      Top             =   5880
      Width           =   1515
   End
   Begin VB.Label lblVpk 
      Height          =   285
      Left            =   3630
      TabIndex        =   24
      Top             =   5490
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "Beam Intensity"
      Height          =   255
      Left            =   2070
      TabIndex        =   21
      Top             =   4560
      Width           =   1155
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Horizontal Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   5670
      TabIndex        =   18
      Top             =   2070
      Width           =   1245
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   7260
      TabIndex        =   16
      Top             =   2490
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7200
      TabIndex        =   15
      Top             =   2070
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
      Top             =   5370
      Width           =   2835
   End
   Begin VB.Label lblFreq 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   1620
      TabIndex        =   10
      Top             =   5790
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
      ForeColor       =   &H00800000&
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
      Top             =   5790
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   9210
      Y1              =   5250
      Y2              =   5250
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
      ForeColor       =   &H00800000&
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
'*****************************************************************
' Written by C. Bolin, October 2004
' Allows user to practice oscope usage.
'
'*****************************************************************
Option Explicit

'module variables
Private m_nRef As Single 'position of reference line
Private m_nXRef As Single 'horizontal reference
Private m_nVoltDiv(9) As Single 'stores volts/div possible values
Private m_nTimeDiv(22) As Single 'stores time/div values
Private m_nTimeFactor As Single 'needed for Time/Div to display correctly
Private m_nProbe As Single
Private m_nColor As Integer 'color of waveform - green component

Private Sub chkPractice_Click()
  If chkPractice.Value = vbChecked Then 'create problem
    frmMain.Height = 5580
    CreateRandomProblem
  Else  'restore
    frmMain.Height = 7845
  
  End If
End Sub

'***********************************************  CreateRandomProblem
' Generates random voltage and frequency.
' Messes up Oscope controls.
Private Sub CreateRandomProblem()
  Dim i As Long
  
  'generate function generator values
  vsbVPk.Value = GetRandom(0, 500)
  hsbFreq.Value = GetRandom(1, 10)
  
  'mess up scope settings
  vsb.Value = GetRandom(-400, 400)
  vsbVDiv.Value = GetRandom(0, 9)
  hsbXRef.Value = GetRandom(0, 22)
  hsbIntensity.Value = GetRandom(0, 127)
  
  'Probe x1, x10
  i = GetRandom(0, 1)
  If i = 0 Then
    opt1X.Value = True
  Else
    opt10X.Value = True
  End If
  
  'illumination
  i = GetRandom(0, 1)
  If i = 0 Then
    chkIlluminate.Value = vbChecked
  Else
    chkIlluminate.Value = vbUnchecked
  End If
  

End Sub

'generates random number between range specified
Private Function GetRandom(min As Integer, max As Integer) As Integer
  Dim i, temp As Integer
  
  'in case min is passed greater than max
  If min > max Then
    temp = max
    max = min
    min = temp
  End If
  
  i = CInt(Rnd * (max - min)) + min
  If i < min Then i = min
  If i > max Then i = max
  
  GetRandom = i
End Function

Private Sub cmdShow_Click()
  chkPractice.Value = vbUnchecked
  
End Sub

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
  m_nTimeFactor = 1
  
  chkPractice_Click
  vsbVPk_Change   'function generator
  hsbFreq_Change
  hsbTime_Change  'scope settings
  vsbVDiv_Change
  hsbXRef_Change
  hsbIntensity_Change
  
  Randomize Timer  'used for problem generation
  
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

  'update solutions
  lblVpk.Caption = "V (PK) : " & vsbVPk.Value / 10
  lblVpkpk.Caption = "V (PK-PK): " & (vsbVPk.Value / 10) * 2
  lblVrms.Caption = "V (RMS): " & (vsbVPk.Value / 10) * 0.707
  lblVavg.Caption = "V (AVG): " & (vsbVPk.Value / 10) * 0.637

  DrawWaveform CSng(vsbVPk.Value / 10), CSng(hsbFreq.Value)
  DrawScopeGrid

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
  
  'draws tic marks on vertical grid line in center
  For i = 1 To 39
    pic.Line (-0.1, -4 + (i * 0.2))-(0.15, -4 + (i * 0.2))
  Next i
  
  'draws tic marks on horizontal grid line in center
  For i = 1 To 49
    pic.Line (-5 + (i * 0.2), 0.1)-(-5 + (i * 0.2), -0.1)
  Next i
  
End Sub

Private Sub hsbIntensity_Change()
  m_nColor = hsbIntensity.Value
  UpdateScope
End Sub

Private Sub hsbIntensity_Scroll()
  hsbIntensity_Change
End Sub

Private Sub hsbTime_Change()
  lblTime.Caption = CStr(m_nTimeDiv(hsbTime.Value)) & " S"
  m_nTimeFactor = m_nTimeDiv(hsbTime.Value) / 0.1
  UpdateScope
End Sub

Private Sub hsbTime_Scroll()
  hsbTime_Change
End Sub


Private Sub hsbXRef_Change()
  m_nXRef = hsbXRef.Value / 100
  UpdateScope
End Sub

Private Sub hsbXRef_Scroll()
  hsbXRef_Change
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
  
  'sets waveform color and thickness
  pic.DrawWidth = 2
  pic.ForeColor = RGB(0, 128 + m_nColor, 128)
  
  'draws waveform...definitely needs optimized
  For i = 0 To (360 * CInt(nFreq) * m_nTimeFactor) - 1
    
    'calcuates next x,y point
    x = CSng(i) / (36 * nFreq * m_nTimeFactor) ' + m_nXRef
    y = m_nRef + Sin(3.14159 / 180 * CSng(i + (32 * m_nXRef))) * (nAmp / m_nVoltDiv(vsbVDiv.Value) / m_nProbe)
    
    'draws line segment between this x,y and last x,y
    If bFirst = False Then
      pic.Line (-5 + x, y)-(-5 + oldx, oldy)
    End If
    oldx = x: oldy = y
    bFirst = False
  Next i
  
  'restores drawing width
  pic.DrawWidth = 1
  
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

