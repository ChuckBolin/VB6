VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RL Circuit Solver v0.01 - Written by C. Bolin"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Problems:"
      Height          =   1995
      Left            =   5880
      TabIndex        =   8
      Top             =   30
      Width           =   4065
      Begin VB.CheckBox chkQ 
         Caption         =   "With Q Factor"
         Height          =   375
         Left            =   2550
         TabIndex        =   13
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton Command4 
         Caption         =   "RL Circuit"
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   1500
         Width           =   2265
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Parallel Inductors"
         Height          =   345
         Left            =   150
         TabIndex        =   11
         Top             =   1080
         Width           =   2265
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Series Inductors"
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   660
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Single Inductor"
         Height          =   345
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8850
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solution:"
      Height          =   825
      Left            =   6120
      TabIndex        =   2
      Top             =   3750
      Width           =   3375
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   405
         Left            =   2790
         TabIndex        =   6
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Height          =   405
         Left            =   1320
         TabIndex        =   5
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   915
      End
      Begin VB.Label lblStep 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   1860
         TabIndex        =   4
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.PictureBox picShow 
      BackColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   9885
      TabIndex        =   1
      Top             =   4650
      Width           =   9945
   End
   Begin VB.PictureBox picDraw 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   30
      ScaleHeight     =   4485
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdNext_Click()
  g_nCurrentStep = g_nCurrentStep + 1
  If g_nCurrentStep > g_nTotalSteps Then g_nCurrentStep = 1
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep

End Sub

Private Sub cmdPrev_Click()
  g_nCurrentStep = g_nCurrentStep - 1
  If g_nCurrentStep < 1 Then g_nCurrentStep = g_nTotalSteps
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep
End Sub

Private Sub cmdShow_Click()
  Frame1.Width = 3400
  If g_eMode = CM_InductorOnly Then
    g_nTotalSteps = 3
  ElseIf g_eMode = CM_InductorQ Then
    g_nTotalSteps = 5
  End If
    
  g_nCurrentStep = 1
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep
End Sub

Private Sub Command1_Click()
  picShow.Cls
  Frame1.Width = 1200
  cmdShow.Enabled = True
  
  If chkQ.Value = vbChecked Then
    g_eMode = CM_InductorQ
  Else
    g_eMode = CM_InductorOnly
  End If
    
  SelectValues
  CreateCircuit
End Sub

'************************************************** SelectValues
Private Sub SelectValues()
  
  g_uSource.Voltage = g_nVolt(Rnd * 6 Mod 6)
  g_uSource.Frequency = g_nFreq(Rnd * 6 Mod 6)
  
  frmMain.Caption = g_uSource.Voltage & "  " & g_uSource.Frequency
  
  
  
  If g_eMode = CM_InductorOnly Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = 0
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uSource.Impedance = g_uInductor(0).XL
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
  End If

  If g_eMode = CM_InductorQ Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = g_nIntResistance(Rnd * 6 Mod 6)
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uSource.Impedance = g_uInductor(0).XL
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
  
  End If

  If g_eMode = CM_SeriesOnly Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = 0
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uInductor(1).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(1).Resistance = 0
    g_uInductor(1).XL = 2 * PI * g_uSource.Frequency * g_uInductor(1).Inductance
    g_uInductor(2).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(2).Resistance = 0
    g_uInductor(2).XL = 2 * PI * g_uSource.Frequency * g_uInductor(2).Inductance
    
    
    g_uSource.Impedance = g_uInductor(0).XL + g_uInductor(1).XL + g_uInductor(2).XL
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(1).Current = g_uSource.Current
    g_uInductor(0).Current = g_uSource.Current
    
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
    g_uInductor(1).Voltage = g_uSource.Current * g_uInductor(1).XL
    g_uInductor(1).VAR = g_uInductor(1).Current ^ 2 * g_uInductor(1).XL
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
  End If

  If g_eMode = CM_SeriesQ Then
  
  
  End If

End Sub

'************************************************** CreateCircuit
'Draws and chooses values
Private Sub CreateCircuit()
  If g_eMode = CM_InductorOnly Or g_eMode = CM_InductorQ Then
    DrawInductorCircuit
  ElseIf g_eMode = CM_SeriesOnly Or g_eMode = CM_SeriesQ Then
    DrawSeriesCircuit
  End If
End Sub

'************************************************** DisplayStep
Private Sub DisplayStep()
  picShow.Cls
  
  If g_eMode = CM_InductorOnly Then
    Select Case g_nCurrentStep
      Case 1:
        picShow.FontBold = True
        ShowText 100, 100, "Step 1:  Solve for XL"
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL = 2 x 3.14 x Freq x Inductance"
        ShowText 100, 700, "b)  XL = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(0).Inductance) & "H"
        ShowText 100, 1000, "c)  XL = " & FormatNumber(g_uInductor(0).XL) & "Ohms"
      Case 2:
        picShow.FontBold = True
        ShowText 100, 100, "Step 2:  Solve for IT"
        picShow.FontBold = False
        ShowText 100, 400, "a)  IT = Va / XL"
        ShowText 100, 700, "b)  IT = " & FormatNumber(g_uSource.Voltage) & "V / " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1000, "c)  IT = " & FormatNumber(g_uSource.Current) & "A"
      Case 3:
        picShow.FontBold = True
        ShowText 100, 100, "Step 3:  Solve for VARs"
        picShow.FontBold = False
        ShowText 100, 400, "a)  VARs = I x I x XL"
        ShowText 100, 700, "b)  VARs = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1000, "c)  VARs = " & FormatNumber(g_uInductor(0).VAR) & "VARs"
        

    End Select
    Exit Sub
  End If

  If g_eMode = CM_InductorQ Then
    Select Case g_nCurrentStep
      Case 1:
      
      
      Case 2:
      
      
      Case 3:
    
    
    End Select
    Exit Sub
  End If



End Sub


Private Sub Command2_Click()
  picShow.Cls
  Frame1.Width = 1200
  cmdShow.Enabled = True
  
  If chkQ.Value = vbChecked Then
    g_eMode = CM_SeriesQ
  Else
    g_eMode = CM_SeriesOnly
  End If
    
  SelectValues
  CreateCircuit
End Sub

Private Sub Form_Load()
  Randomize Timer
  LoadVariables
  Frame1.Width = 1200
End Sub
