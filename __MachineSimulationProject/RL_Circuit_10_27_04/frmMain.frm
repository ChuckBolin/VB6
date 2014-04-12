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
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   1500
         Width           =   2265
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Parallel Inductors"
         Enabled         =   0   'False
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
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   9885
      TabIndex        =   1
      Top             =   4650
      Width           =   9945
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
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

Private Sub chkQ_Click()
  picShow.Cls
  picDraw.Cls
End Sub

'displays next solution step
Private Sub cmdNext_Click()
  g_nCurrentStep = g_nCurrentStep + 1
  If g_nCurrentStep > g_nTotalSteps Then g_nCurrentStep = 1
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep

End Sub

'displays previous solution step
Private Sub cmdPrev_Click()
  g_nCurrentStep = g_nCurrentStep - 1
  If g_nCurrentStep < 1 Then g_nCurrentStep = g_nTotalSteps
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep
End Sub

'displays first step of solution for problem
Private Sub cmdShow_Click()
  Frame1.Width = 3400
  If g_eMode = CM_InductorOnly Then
    g_nTotalSteps = 3
  ElseIf g_eMode = CM_InductorQ Then
    g_nTotalSteps = 5
  ElseIf g_eMode = CM_SeriesOnly Then
    g_nTotalSteps = 6
  ElseIf g_eMode = CM_SeriesQ Then
    g_nTotalSteps = 8
    
  End If
    
  g_nCurrentStep = 1
  lblStep.Caption = g_nCurrentStep & " of " & g_nTotalSteps
  DisplayStep
End Sub

'Single inductor with or without internal resistance
Private Sub Command1_Click()
  picShow.Cls
  Frame1.Width = 1200
  cmdShow.Enabled = True
  
  If chkQ.Value = vbChecked Then
    g_eMode = CM_InductorQ
  Else
    g_eMode = CM_InductorOnly
  End If
    
  SelectAndCalculateValues
  CreateCircuit
End Sub

'************************************************** SelectAndCalculateValues
'Depending upon user mode, values are selected randomly from
'arrays of standard values. All calculations are performed here.
Private Sub SelectAndCalculateValues()
  Dim nR, nXL As Single  'stores total resistance and inductance for solving Z
  
  'select voltage and frequency
  g_uSource.Voltage = g_nVolt(Rnd * 6 Mod 6)
  g_uSource.Frequency = g_nFreq(Rnd * 6 Mod 6)
  
  'frmMain.Caption = g_uSource.Voltage & "  " & g_uSource.Frequency
  
  If g_eMode = CM_InductorOnly Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = 0
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uSource.Impedance = g_uInductor(0).XL
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
  ElseIf g_eMode = CM_InductorQ Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = g_nIntResistance(Rnd * 6 Mod 6)
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uSource.Impedance = Sqr(g_uInductor(0).XL ^ 2 + g_uInductor(0).Resistance ^ 2)
    g_uInductor(0).Z = g_uSource.Impedance
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
  ElseIf g_eMode = CM_SeriesOnly Then
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
    g_uInductor(2).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).XL
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
    g_uInductor(1).Voltage = g_uSource.Current * g_uInductor(1).XL
    g_uInductor(1).VAR = g_uInductor(1).Current ^ 2 * g_uInductor(1).XL
    g_uInductor(2).Voltage = g_uSource.Current * g_uInductor(2).XL
    g_uInductor(2).VAR = g_uInductor(2).Current ^ 2 * g_uInductor(2).XL
  ElseIf g_eMode = CM_SeriesQ Then
    g_uInductor(0).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(0).Resistance = g_nIntResistance(Rnd * 6 Mod 6)
    g_uInductor(0).XL = 2 * PI * g_uSource.Frequency * g_uInductor(0).Inductance
    g_uInductor(0).Z = Sqr(g_uInductor(0).Resistance ^ 2 + g_uInductor(0).XL ^ 2)
    g_uInductor(1).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(1).Resistance = g_nIntResistance(Rnd * 6 Mod 6)
    g_uInductor(1).XL = 2 * PI * g_uSource.Frequency * g_uInductor(1).Inductance
    g_uInductor(1).Z = Sqr(g_uInductor(1).Resistance ^ 2 + g_uInductor(1).XL ^ 2)
    g_uInductor(2).Inductance = g_nInductor(Rnd * 6 Mod 6)
    g_uInductor(2).Resistance = g_nIntResistance(Rnd * 6 Mod 6)
    g_uInductor(2).XL = 2 * PI * g_uSource.Frequency * g_uInductor(2).Inductance
    g_uInductor(2).Z = Sqr(g_uInductor(2).Resistance ^ 2 + g_uInductor(2).XL ^ 2)
    nR = g_uInductor(0).Resistance + g_uInductor(1).Resistance + g_uInductor(2).Resistance
    nXL = g_uInductor(0).XL + g_uInductor(1).XL + g_uInductor(2).XL
    g_uSource.Impedance = Sqr(nR ^ 2 + nXL ^ 2)
    g_uSource.Current = g_uSource.Voltage / g_uSource.Impedance
    g_uInductor(0).Current = g_uSource.Current
    g_uInductor(1).Current = g_uSource.Current
    g_uInductor(2).Current = g_uSource.Current
    g_uInductor(0).Voltage = g_uSource.Current * g_uInductor(0).Z
    g_uInductor(0).VAR = g_uInductor(0).Current ^ 2 * g_uInductor(0).XL
    g_uInductor(0).Power = g_uInductor(0).Current ^ 2 * g_uInductor(0).Resistance
    g_uInductor(0).VA = g_uInductor(0).Current ^ 2 * g_uInductor(0).Z
    g_uInductor(1).Voltage = g_uSource.Current * g_uInductor(1).Z
    g_uInductor(1).VAR = g_uInductor(1).Current ^ 2 * g_uInductor(1).XL
    g_uInductor(1).Power = g_uInductor(1).Current ^ 2 * g_uInductor(1).Resistance
    g_uInductor(1).VA = g_uInductor(1).Current ^ 2 * g_uInductor(1).Z
    g_uInductor(2).Voltage = g_uSource.Current * g_uInductor(2).Z
    g_uInductor(2).VAR = g_uInductor(2).Current ^ 2 * g_uInductor(2).XL
    g_uInductor(2).Power = g_uInductor(2).Current ^ 2 * g_uInductor(2).Resistance
    g_uInductor(2).VA = g_uInductor(2).Current ^ 2 * g_uInductor(2).Z
  
  
  End If

End Sub

'************************************************** CreateCircuit
'Draws circuit based upon user mode and values selected.
Private Sub CreateCircuit()
  If g_eMode = CM_InductorOnly Or g_eMode = CM_InductorQ Then
    DrawInductorCircuit
  ElseIf g_eMode = CM_SeriesOnly Or g_eMode = CM_SeriesQ Then
    DrawSeriesCircuit
  End If
End Sub

'************************************************** DisplayStep
'Displays solution in picShow step by step.
Private Sub DisplayStep()
  picShow.Cls
  
  If g_eMode = CM_InductorOnly Then
    Select Case g_nCurrentStep
      Case 1:
        picShow.FontBold = True
        ShowText 100, 100, "Step 1:  Solve for XL (Purely inductive circuit, no resistance)"
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL = 2 x 3.14 x F x L"
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
  
  ElseIf g_eMode = CM_InductorQ Then
    Select Case g_nCurrentStep
      Case 1:
        picShow.FontBold = True
        ShowText 100, 100, "Step 1:  Solve for XL of the Coil"
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL = 2 x 3.14 x F x L"
        ShowText 100, 700, "b)  XL = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(0).Inductance) & "H"
        ShowText 100, 1000, "c)  XL = " & FormatNumber(g_uInductor(0).XL) & "Ohms"
      Case 2:
        picShow.FontBold = True
        ShowText 100, 100, "Step 2:  Solve for Impedance (Z) of the Coil"
        picShow.FontBold = False
        ShowText 100, 400, "a)  Z =   R   +   XL  "
        ShowSpecial SPECIAL_RADICAL, 800, 420
        ShowSpecial SPECIAL_SQUARE, 1065, 420
        ShowSpecial SPECIAL_SQUARE, 1905, 420
        ShowText 100, 700, "b)  Z =   " & Format(g_uInductor(0).Resistance ^ 2, "###.###") & "  +  " & Format(g_uInductor(0).XL ^ 2, "###.###")
        ShowSpecial SPECIAL_RADICAL, 795, 735
        ShowText 100, 1000, "c)  Z = " & FormatNumber(g_uInductor(0).Z) & "Ohms"
      Case 3:
        picShow.FontBold = True
        ShowText 100, 100, "Step 3:  Solve for IT"
        picShow.FontBold = False
        ShowText 100, 400, "a)  IT = Va / Z"
        ShowText 100, 700, "b)  IT = " & FormatNumber(g_uSource.Voltage) & "V / " & FormatNumber(g_uInductor(0).Z) & "Ohms"
        ShowText 100, 1000, "c)  IT = " & FormatNumber(g_uSource.Current) & "A"
      Case 4:
        picShow.FontBold = True
        ShowText 100, 100, "Step 4:  Solve for VARs"
        picShow.FontBold = False
        ShowText 100, 400, "a)  VARs = I x I x XL"
        ShowText 100, 700, "b)  VARs = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1000, "c)  VARs = " & FormatNumber(g_uInductor(0).VAR) & "VARs"
      Case 5:
        picShow.FontBold = True
        ShowText 100, 100, "Step 2:  Solve for Q of the Coil"
        picShow.FontBold = False
        ShowText 100, 400, "a)  Q = XL / R"
        ShowText 100, 700, "b)  Q =   " & Format(g_uInductor(0).Resistance ^ 2, "###.###") & "  +  " & Format(g_uInductor(0).XL ^ 2, "###.###")
        ShowSpecial SPECIAL_RADICAL, 795, 735
        ShowText 100, 1000, "c)  Z = " & FormatNumber(g_uInductor(0).Z) & "Ohms"
    End Select
    Exit Sub
  
  ElseIf g_eMode = CM_SeriesOnly Then
    Select Case g_nCurrentStep
      Case 1:
        picShow.FontBold = True
        ShowText 100, 100, "Step 1:  Solve for XL for each Coil"
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL = 2 x 3.14 x F x L"
        ShowText 100, 800, "b)  XL1 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(0).Inductance) & "H"
        ShowText 100, 1100, "c)  XL1 = " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1500, "d)  XL2 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(1).Inductance) & "H"
        ShowText 100, 1800, "e)  XL2 = " & FormatNumber(g_uInductor(1).XL) & "Ohms"
        ShowText 100, 2200, "f)  XL3 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(2).Inductance) & "H"
        ShowText 100, 2500, "g)  XL3 = " & FormatNumber(g_uInductor(2).XL) & "Ohms"
      Case 2:
        picShow.FontBold = True
        ShowText 100, 100, "Step 2:  Solve for total XL"
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL Total = XL1 + XL2 + XL3"
        ShowText 100, 800, "b)  XL Total = " & FormatNumber(g_uInductor(0).XL) & "Ohms  + " & FormatNumber(g_uInductor(1).XL) & "Ohms" & " + " & FormatNumber(g_uInductor(2).XL) & "Ohms"
        ShowText 100, 1100, "c)  XL Total = " & FormatNumber(g_uSource.Impedance) & "Ohms"
      Case 3:
        picShow.FontBold = True
        ShowText 100, 100, "Step 3:  Solve for IT"
        picShow.FontBold = False
        ShowText 100, 400, "a)  IT = Va / XL"
        ShowText 100, 800, "b)  IT = " & FormatNumber(g_uSource.Voltage) & "V / " & FormatNumber(g_uSource.Impedance) & "Ohms"
        ShowText 100, 1100, "c)  IT = " & FormatNumber(g_uSource.Current) & "A"
      Case 4:
        picShow.FontBold = True
        ShowText 100, 100, "Step 4:  Calculate voltage drops across each inductor."
        picShow.FontBold = False
        ShowText 100, 400, "a)  V L = XL x IT"
        ShowText 100, 800, "b)  VL1 = " & FormatNumber(g_uInductor(0).XL) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 1100, "c)  VL1 = " & FormatNumber(g_uInductor(0).Voltage) & "V"
        ShowText 100, 1500, "d)  VL2 = " & FormatNumber(g_uInductor(1).XL) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 1800, "e)  VL2 = " & FormatNumber(g_uInductor(1).Voltage) & "V"
        ShowText 100, 2200, "f)  VL3 = " & FormatNumber(g_uInductor(2).XL) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 2500, "g)  VL3 = " & FormatNumber(g_uInductor(2).Voltage) & "V"
      Case 5:
        picShow.FontBold = True
        ShowText 100, 100, "Step 5:  Check voltage drops against applied voltage (Va)."
        picShow.FontBold = False
        ShowText 100, 400, "a)  VA = VL1 + VL2 + VL3"
        ShowText 100, 800, "b)  " & FormatNumber(g_uSource.Voltage) & " = " & FormatNumber(g_uInductor(0).Voltage) & "V + " & FormatNumber(g_uInductor(1).Voltage) & "V + " & FormatNumber(g_uInductor(2).Voltage) & "V"
        ShowText 100, 1100, "c)  Compare voltages: " & FormatNumber(g_uSource.Voltage) & "V = " & FormatNumber(g_uInductor(0).Voltage + g_uInductor(1).Voltage + g_uInductor(2).Voltage) & "V"
      Case 6:
        picShow.FontBold = True
        ShowText 100, 100, "Step 6:  Solve for VARs"
        picShow.FontBold = False
        ShowText 100, 400, "a)  VARs = I x I x XL"
        ShowText 100, 800, "b)  VARs L1 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1100, "c)  VARs L1 = " & FormatNumber(g_uInductor(0).VAR) & "VARs"
        ShowText 100, 1500, "d)  VARs L2 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(1).XL) & "Ohms"
        ShowText 100, 1800, "e)  VARs L2 = " & FormatNumber(g_uInductor(1).VAR) & "VARs"
        ShowText 100, 2200, "f)  VARs L3 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(2).XL) & "Ohms"
        ShowText 100, 2500, "g)  VARs L3 = " & FormatNumber(g_uInductor(2).VAR) & "VARs"
    End Select
    Exit Sub
  ElseIf g_eMode = CM_SeriesQ Then
    Select Case g_nCurrentStep
      Case 1:
        picShow.FontBold = True
        ShowText 100, 100, "Step 1:  Solve for XL for each Coil."
        picShow.FontBold = False
        ShowText 100, 400, "a)  XL = 2 x 3.14 x F x L"
        ShowText 100, 800, "b)  XL1 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(0).Inductance) & "H"
        ShowText 100, 1100, "c)  XL1 = " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1500, "d)  XL2 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(1).Inductance) & "H"
        ShowText 100, 1800, "e)  XL2 = " & FormatNumber(g_uInductor(1).XL) & "Ohms"
        ShowText 100, 2200, "f)  XL3 = 6.28 x " & FormatNumber(g_uSource.Frequency) & "Hz x " & FormatNumber(g_uInductor(2).Inductance) & "H"
        ShowText 100, 2500, "g)  XL3 = " & FormatNumber(g_uInductor(2).XL) & "Ohms"
      Case 2:
        picShow.FontBold = True
        ShowText 100, 100, "Step 2:  Solve for Z for each Coil."
        picShow.FontBold = False
        ShowText 100, 400, "a)  Z =   R   +   XL  "
        ShowSpecial SPECIAL_RADICAL, 800, 420
        ShowSpecial SPECIAL_SQUARE, 1065, 420
        ShowSpecial SPECIAL_SQUARE, 1905, 420
        ShowText 100, 800, "b)  Z1 =   " & Format(g_uInductor(0).Resistance ^ 2, "###.###") & "  +  " & Format(g_uInductor(0).XL ^ 2, "###.###")
        ShowSpecial SPECIAL_RADICAL, 920, 785
        ShowText 100, 1100, "c)  Z1 = " & FormatNumber(g_uInductor(0).Z) & "Ohms"
        
        ShowText 100, 1500, "d)  Z2 =   " & Format(g_uInductor(1).Resistance ^ 2, "###.###") & "  +  " & Format(g_uInductor(1).XL ^ 2, "###.###")
        ShowSpecial SPECIAL_RADICAL, 920, 1485
        ShowText 100, 1800, "e)  Z2 = " & FormatNumber(g_uInductor(1).Z) & "Ohms"

        ShowText 100, 2200, "f)  Z3 =   " & Format(g_uInductor(2).Resistance ^ 2, "###.###") & "  +  " & Format(g_uInductor(2).XL ^ 2, "###.###")
        ShowSpecial SPECIAL_RADICAL, 920, 2185
        ShowText 100, 2500, "g)  Z3 = " & FormatNumber(g_uInductor(2).Z) & "Ohms"

       Case 3:
        picShow.FontBold = True
        ShowText 100, 100, "Step 3:  Solve for total Z."
        picShow.FontBold = False
        ShowText 100, 400, "a)  Z = Z1 + Z2 + Z3"
        ShowText 100, 800, "b)  Z = " & FormatNumber(g_uInductor(0).Z) & "Ohms  + " & FormatNumber(g_uInductor(1).Z) & "Ohms" & " + " & FormatNumber(g_uInductor(2).Z) & "Ohms"
        ShowText 100, 1100, "c)  Z Total = " & FormatNumber(g_uSource.Impedance) & "Ohms"
      Case 4:
        picShow.FontBold = True
        ShowText 100, 100, "Step 4:  Solve for IT"
        picShow.FontBold = False
        ShowText 100, 400, "a)  IT = Va / Z"
        ShowText 100, 800, "b)  IT = " & FormatNumber(g_uSource.Voltage) & "V / " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 1100, "c)  IT = " & FormatNumber(g_uSource.Current) & "A"
      Case 5:
        picShow.FontBold = True
        ShowText 100, 100, "Step 5:  Calculate voltage drops across each inductor."
        picShow.FontBold = False
        ShowText 100, 400, "a)  V L = Z x IT"
        ShowText 100, 800, "b)  VL1 = " & FormatNumber(g_uInductor(0).Z) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 1100, "c)  VL1 = " & FormatNumber(g_uInductor(0).Voltage) & "V"
        ShowText 100, 1500, "d)  VL2 = " & FormatNumber(g_uInductor(1).Z) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 1800, "e)  VL2 = " & FormatNumber(g_uInductor(1).Voltage) & "V"
        ShowText 100, 2200, "f)  VL3 = " & FormatNumber(g_uInductor(2).Z) & "Ohms x " & FormatNumber(g_uSource.Current) & "A"
        ShowText 100, 2500, "g)  VL3 = " & FormatNumber(g_uInductor(2).Voltage) & "V"
      Case 6:
        picShow.FontBold = True
        ShowText 100, 100, "Step 6:  Check voltage drops against applied voltage (Va)."
        picShow.FontBold = False
        ShowText 100, 400, "a)  VA = VL1 + VL2 + VL3"
        ShowText 100, 800, "b)  " & FormatNumber(g_uSource.Voltage) & " = " & FormatNumber(g_uInductor(0).Voltage) & "V + " & FormatNumber(g_uInductor(1).Voltage) & "V + " & FormatNumber(g_uInductor(2).Voltage) & "V"
        ShowText 100, 1100, "c)  Compare voltages: " & FormatNumber(g_uSource.Voltage) & "V = " & FormatNumber(g_uInductor(0).Voltage + g_uInductor(1).Voltage + g_uInductor(2).Voltage) & "V"
      Case 7:
        picShow.FontBold = True
        ShowText 100, 100, "Step 7:  Solve for VARs"
        picShow.FontBold = False
        ShowText 100, 400, "a)  VARs = I x I x XL"
        ShowText 100, 800, "b)  VARs L1 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).XL) & "Ohms"
        ShowText 100, 1100, "c)  VARs L1 = " & FormatNumber(g_uInductor(0).VAR) & "VARs"
        ShowText 100, 1500, "d)  VARs L2 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(1).XL) & "Ohms"
        ShowText 100, 1800, "e)  VARs L2 = " & FormatNumber(g_uInductor(1).VAR) & "VARs"
        ShowText 100, 2200, "f)  VARs L3 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(2).XL) & "Ohms"
        ShowText 100, 2500, "g)  VARs L3 = " & FormatNumber(g_uInductor(2).VAR) & "VARs"
      Case 8:
        picShow.FontBold = True
        ShowText 100, 100, "Step 8:  Solve for Power across each Coil."
        picShow.FontBold = False
        ShowText 100, 400, "a)  P = I x I x R"
        ShowText 100, 800, "b)  PL1 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).Resistance) & "Ohms"
        ShowText 100, 1100, "c)  PL1 = " & FormatNumber(g_uInductor(0).Power) & "W"
        ShowText 100, 1500, "d)  PL2 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(1).Resistance) & "Ohms"
        ShowText 100, 1800, "e)  PL2 = " & FormatNumber(g_uInductor(1).Power) & "W"
        ShowText 100, 2200, "f)  PL3 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(2).Resistance) & "Ohms"
        ShowText 100, 2500, "g)  PL3 = " & FormatNumber(g_uInductor(2).Power) & "W"
     Case 9:
        picShow.FontBold = True
        ShowText 100, 100, "Step 7:  Solve for Power"
        picShow.FontBold = False
        ShowText 100, 400, "a)  P = I x I x R"
        ShowText 100, 800, "b)  PL1 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(0).Resistance) & "Ohms"
        ShowText 100, 1100, "c)  PL1 = " & FormatNumber(g_uInductor(0).Power) & "W"
        ShowText 100, 1500, "d)  PL2 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(1).Resistance) & "Ohms"
        ShowText 100, 1800, "e)  PL2 = " & FormatNumber(g_uInductor(1).Power) & "W"
        ShowText 100, 2200, "f)  PL3 = " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uSource.Current) & "A x " & FormatNumber(g_uInductor(2).Resistance) & "Ohms"
        ShowText 100, 2500, "g)  PL3 = " & FormatNumber(g_uInductor(2).Power) & "W"

    End Select
    Exit Sub
  End If

End Sub

'Series inductive circuit with or without internal resistance
Private Sub Command2_Click()
  picShow.Cls
  Frame1.Width = 1200
  cmdShow.Enabled = True
  
  If chkQ.Value = vbChecked Then
    g_eMode = CM_SeriesQ
  Else
    g_eMode = CM_SeriesOnly
  End If
    
  SelectAndCalculateValues
  CreateCircuit
End Sub

'Initialization of program
Private Sub Form_Load()
  Randomize Timer
  LoadVariables  'located in global.bas
  Frame1.Width = 1200
End Sub

Private Sub picShow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'frmMain.Caption = x & " " & y
End Sub
