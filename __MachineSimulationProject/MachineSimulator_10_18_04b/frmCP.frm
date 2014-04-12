VERSION 5.00
Begin VB.Form frmCP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3945
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2580
      Top             =   3960
   End
   Begin VB.Line linAuto 
      BorderWidth     =   3
      X1              =   2280
      X2              =   2490
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   990
      TabIndex        =   11
      Top             =   3540
      Width           =   765
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   420
      Index           =   11
      Left            =   570
      Shape           =   3  'Circle
      Top             =   3570
      Width           =   420
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   2670
      TabIndex        =   10
      Top             =   3180
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   2670
      TabIndex        =   9
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   2670
      TabIndex        =   8
      Top             =   2070
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   2700
      TabIndex        =   7
      Top             =   1500
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   2700
      TabIndex        =   6
      Top             =   900
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   300
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   990
      TabIndex        =   4
      Top             =   2970
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   990
      TabIndex        =   3
      Top             =   2430
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   990
      TabIndex        =   2
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   990
      TabIndex        =   1
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label lblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conveyor On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   990
      TabIndex        =   0
      Top             =   750
      Width           =   765
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   225
      Index           =   10
      Left            =   2340
      Shape           =   3  'Circle
      Top             =   2700
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   9
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   8
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   7
      Left            =   2250
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   6
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   990
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   225
      Index           =   5
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   540
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   4
      Left            =   660
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   3
      Left            =   660
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   225
      Index           =   2
      Left            =   690
      Shape           =   3  'Circle
      Top             =   1860
      Width           =   225
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   1
      Left            =   660
      Shape           =   3  'Circle
      Top             =   1380
      Width           =   225
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   10
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2670
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   9
      Left            =   2190
      Shape           =   3  'Circle
      Top             =   2250
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   8
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1860
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   7
      Left            =   2190
      Shape           =   3  'Circle
      Top             =   1470
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   6
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   960
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   5
      Left            =   2100
      Shape           =   3  'Circle
      Top             =   450
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   4
      Left            =   630
      Shape           =   3  'Circle
      Top             =   2730
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   3
      Left            =   600
      Shape           =   3  'Circle
      Top             =   2250
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   2
      Left            =   630
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   315
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   1
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   315
   End
   Begin VB.Shape shpPush 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   0
      Left            =   660
      Shape           =   3  'Circle
      Top             =   840
      Width           =   225
   End
   Begin VB.Shape shpBase 
      Height          =   315
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   810
      Width           =   315
   End
   Begin VB.Shape shpCP 
      BackStyle       =   1  'Opaque
      Height          =   4155
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3705
   End
End
Attribute VB_Name = "frmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_nID As Integer

Private Sub Form_Load()
  Dim i As Integer
  m_nID = -1
  
  'draw control panel buttons and indicators
  shpCP.BackColor = RGB(249, 240, 196)
 
  'align left button and indicators
   For i = 0 To 4
    shpBase(i).Width = 340
    shpBase(i).Height = 340
    shpPush(i).Width = 260
    shpPush(i).Height = 260
    shpBase(i).Left = 420
    shpPush(i).Left = shpBase(i).Left + (shpBase(i).Width - shpPush(i).Width) \ 2
    shpBase(i).Top = 420 + i * 520
    shpPush(i).Top = shpBase(i).Top + (shpBase(i).Height - shpPush(i).Height) \ 2
    lblTag(i).Left = shpBase(i).Left + shpBase(i).Width + 100
    lblTag(i).Top = shpBase(i).Top - 100
  Next i
    
  'align right buttons and indicators
  For i = 5 To 10
    shpBase(i).Width = 340
    shpBase(i).Height = 340
    shpPush(i).Width = 260
    shpPush(i).Height = 260
    shpBase(i).Left = 2000
    shpPush(i).Left = shpBase(i).Left + (shpBase(i).Width - shpPush(i).Width) \ 2
    shpBase(i).Top = 420 + ((i - 5) * 520)
    shpPush(i).Top = shpBase(i).Top + (shpBase(i).Height - shpPush(i).Height) \ 2
    lblTag(i).Left = shpBase(i).Left + shpBase(i).Width + 100
    lblTag(i).Top = shpBase(i).Top - 100
  Next i

  'this is line for auto/semi select switch
    linAuto.x1 = shpPush(8).Left + (shpPush(8).Width \ 2)
    linAuto.y1 = shpPush(8).Top + (shpPush(8).Height \ 2)
    linAuto.x2 = linAuto.x1 + 200
    linAuto.y2 = linAuto.y1 - 100
    

  'estop
  shpPush(11).Left = 360
  shpPush(11).Top = 3150
  lblTag(11).Left = shpPush(11).Left + 500
  lblTag(11).Top = shpPush(11).Top

  'labels
  lblTag(0).Caption = "Conveyor On"
  lblTag(1).Caption = "Conveyor Off"
  lblTag(2).Caption = "Control Power"
  lblTag(3).Caption = "Control On"
  lblTag(4).Caption = "Control Off"
  lblTag(5).Caption = "Automatic"
  lblTag(6).Caption = "Auto Start"
  lblTag(7).Caption = "Auto Stop"
  lblTag(8).Caption = "Auto  Semi"
  lblTag(9).Caption = "Jog"
  lblTag(10).Caption = "Fault"
  lblTag(11).Caption = "E-Stop"
 
 'positions form relative to frmPLC if it is loaded
  'If frmPLC.Enabled = True Then
  '  frmCP.Left = frmPLC.Left
 '   frmCP.Top = frmPLC.Top + frmPLC.Height
 ' End If
  
End Sub

'decide if button is pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  'left button pressed
  If Button = 1 Then
    If X > shpBase(0).Left And X < shpBase(0).Left + shpBase(0).Width Then 'left column
      If Y > shpBase(0).Top And Y < shpBase(0).Top + shpBase(0).Height Then 'conveyor on
        m_nID = 0
      ElseIf Y > shpBase(1).Top And Y < shpBase(1).Top + shpBase(1).Height Then 'conveyor off
        m_nID = 1
      ElseIf Y > shpBase(2).Top And Y < shpBase(2).Top + shpBase(2).Height Then 'control power
        m_nID = 2
      ElseIf Y > shpBase(3).Top And Y < shpBase(3).Top + shpBase(3).Height Then 'control on
        m_nID = 3
      ElseIf Y > shpBase(4).Top And Y < shpBase(4).Top + shpBase(4).Height Then 'control off
        m_nID = 4
      Else
      
      End If
    ElseIf X > shpBase(5).Left And X < shpBase(5).Left + shpBase(5).Width Then 'right column
      If Y > shpBase(5).Top And Y < shpBase(5).Top + shpBase(5).Height Then 'auto indicator
        m_nID = 5
      ElseIf Y > shpBase(6).Top And Y < shpBase(6).Top + shpBase(6).Height Then 'auto start
        m_nID = 6
      ElseIf Y > shpBase(7).Top And Y < shpBase(7).Top + shpBase(7).Height Then 'auto stop
        m_nID = 7
      ElseIf Y > shpBase(8).Top And Y < shpBase(8).Top + shpBase(8).Height Then 'auto semi
        m_nID = 8
      ElseIf Y > shpBase(9).Top And Y < shpBase(9).Top + shpBase(9).Height Then 'jog
        m_nID = 9
      ElseIf Y > shpBase(10).Top And Y < shpBase(10).Top + shpBase(10).Height Then 'fault
        m_nID = 10
      End If
    ElseIf X > shpBase(11).Left And X < shpBase(11).Left + shpBase(11).Width Then 'left column
      If Y > shpBase(11).Top And Y < shpBase(11).Top + shpBase(11).Height Then 'conveyor on
        m_nID = 11
      End If
    End If
  End If
  
  'button or indicator clicked
  If m_nID > -1 Then
    Select Case m_nID
      Case 0:
        e(S4_CONVEYOR_ON) = True
        shpPush(0).Left = shpPush(0).Left + 25
        shpPush(0).Top = shpPush(0).Top + 25
      Case 1:
        e(S5_CONVEYOR_OFF) = True
        shpPush(1).Left = shpPush(1).Left + 25
        shpPush(1).Top = shpPush(1).Top + 25
      Case 2:
        
      Case 3:
        e(S3_CONT_ON) = True
        shpPush(3).Left = shpPush(3).Left + 25
        shpPush(3).Top = shpPush(3).Top + 25
      Case 4:
        e(S2_CONT_OFF) = True
        shpPush(4).Left = shpPush(4).Left + 25
        shpPush(4).Top = shpPush(4).Top + 25
      Case 5:
      
      Case 6:
        e(S7_AUTOSTART) = True
        shpPush(6).Left = shpPush(6).Left + 25
        shpPush(6).Top = shpPush(6).Top + 25
      Case 7:
        e(S8_AUTOSTOP) = True
        shpPush(7).Left = shpPush(7).Left + 25
        shpPush(7).Top = shpPush(7).Top + 25
      Case 8:
        e(S6_AUTO) = Not e(S6_AUTO)  'these must be opposite each other
        e(S6_SEMI) = Not e(S6_AUTO)
        If e(S6_AUTO) Then
          linAuto.x2 = linAuto.x1 + 200
          linAuto.y2 = linAuto.y1 - 100
        Else
          linAuto.x2 = linAuto.x1 + 200
          linAuto.y2 = linAuto.y1 + 100
        End If
        
      Case 9:
        e(S9_JOG) = True
        shpPush(9).Left = shpPush(9).Left + 25
        shpPush(9).Top = shpPush(9).Top + 25
      Case 10:
        'shpPush(10).Left = shpPush(10).Left + 25
        'shpPush(10).Top = shpPush(10).Top + 25
      Case 11:
        'e(S1_stop) = True
        'shpPush(11).Left = shpPush(9).Left + 25
        'shpPush(11).Top = shpPush(9).Top + 25

    
    End Select
    tmrAnimate.Enabled = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.mnuViewControlPanel.Checked = True Then frmMain.mnuViewControlPanel.Checked = False
End Sub

'completes release of pushbutton for effects
Private Sub tmrAnimate_Timer()
  Select Case m_nID
    Case 0:
      e(S4_CONVEYOR_ON) = False
      shpPush(0).Left = shpPush(0).Left - 25
      shpPush(0).Top = shpPush(0).Top - 25
    Case 1:
      e(S5_CONVEYOR_OFF) = False
      shpPush(1).Left = shpPush(1).Left - 25
      shpPush(1).Top = shpPush(1).Top - 25
    Case 2:
    
    Case 3:
      e(S3_CONT_ON) = False
      shpPush(3).Left = shpPush(3).Left - 25
      shpPush(3).Top = shpPush(3).Top - 25
    Case 4:
      e(S2_CONT_OFF) = False
      shpPush(4).Left = shpPush(4).Left - 25
      shpPush(4).Top = shpPush(4).Top - 25
    Case 5:
    
    Case 6:
      e(S7_AUTOSTART) = False
      shpPush(6).Left = shpPush(6).Left - 25
      shpPush(6).Top = shpPush(6).Top - 25
    Case 7:
      e(S8_AUTOSTOP) = False
      shpPush(7).Left = shpPush(7).Left - 25
      shpPush(7).Top = shpPush(7).Top - 25
    Case 8:
      'e(S6_AUTO) = Not e(S6_AUTO)  'these must be opposite each other
      'e(S6_SEMI) = Not e(S6_AUTO)
      'shpPush(8).Left = shpPush(8).Left - 25
      'shpPush(8).Top = shpPush(8).Top - 25
    Case 9:
      e(S9_JOG) = False
      shpPush(9).Left = shpPush(9).Left - 25
      shpPush(9).Top = shpPush(9).Top - 25
    Case 10:
      'shpPush(10).Left = shpPush(10).Left - 25
      'shpPush(10).Top = shpPush(10).Top - 25


  End Select
  tmrAnimate.Enabled = False
  m_nID = -1
End Sub
