VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DC Circuit Troubleshooter v0.51 by C. Bolin, October 15, 2004"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   11895
   Begin VB.PictureBox picR 
      BackColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   4515
      Left            =   -30
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4455
      ScaleWidth      =   7665
      TabIndex        =   9
      Top             =   5490
      Width           =   7725
      Begin VB.Label lblMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Digital Display"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   360
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Shape shpMeter 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   1065
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   330
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.CheckBox chkFault 
      Caption         =   "Insert Fault(s)"
      Height          =   435
      Left            =   10590
      TabIndex        =   8
      Top             =   90
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   10620
      TabIndex        =   7
      Top             =   6630
      Width           =   1245
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Solution"
      Height          =   375
      Left            =   8220
      TabIndex        =   6
      Top             =   6660
      Width           =   1575
   End
   Begin VB.TextBox txtAns 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   8220
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdCombo2 
      Caption         =   "Create Combination Two"
      Height          =   375
      Left            =   8220
      TabIndex        =   4
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton cmdCombo1 
      Caption         =   "Create Combination One"
      Height          =   375
      Left            =   8220
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdPar 
      Caption         =   "Create Parallel"
      Height          =   375
      Left            =   8220
      TabIndex        =   2
      Top             =   420
      Width           =   2295
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Create Series"
      Height          =   375
      Left            =   8220
      TabIndex        =   1
      Top             =   30
      Width           =   2295
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5505
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   0
      Width           =   7725
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constants for components.  D = Device
Private Const D_R1 = 0
Private Const D_R2 = 1
Private Const D_R3 = 2
Private Const D_R4 = 3

'voltage nodes and voltage drops. V = Voltage
Private Const V_SOURCE = 0
Private Const V_R1 = 1
Private Const V_R2 = 2
Private Const V_R3 = 3
Private Const V_R4 = 4
Private Const V_A = 5
Private Const V_B = 6
Private Const V_C = 7
Private Const V_D = 8
Private Const V_E = 9
Private Const V_F = 10
Private Const V_G = 11
Private Const V_H = 12
Private Const V_I = 13
Private Const V_J = 14
Private Const V_K = 15
Private Const V_L = 16
Private Const V_M = 17
Private Const V_N = 18
Private Const V_O = 19
Private Const V_COMMON = 20

Private Enum RES_MODE
  NONE = 0
  SERIES
  PARALLEL
  COMBO1
  COMBO2
End Enum

Private Enum DEVICE_FAULT
  D_OK = 1
  D_OPEN = 10000000
  D_SHORT = 0
End Enum

Private Type Resistor
  Resistor As Single
  value As Single
  Fault As DEVICE_FAULT
  Voltage As Single
  Current As Single
End Type

Private Type COLOR_BAND
  First As Long
  Second As Long
  Third As Long
  Fourth As Long
End Type

'module variables
Private m_nMode As RES_MODE
Private m_nRes(12) As Single  'stores standard resistor values
Private m_nVolt(6) As Single  'stores standard voltage values
Private m_uR(3) As Resistor   '0-3...stores all specific info for resistors
Private m_nRT As Single 'total parameters
Private m_nVS As Single
Private m_nIT As Single
Private m_bFault As Boolean
Private m_uBand As COLOR_BAND 'store four long colors for color bands
Private m_nVoltage(20) As Single 'stores voltages
Private m_sFault As String 'fault description

'************************************************ chkFault
'select if program should add fault
Private Sub chkFault_Click()
  If chkFault.value = vbChecked Then
    m_bFault = True
  Else
    m_bFault = False
  End If
  
  'clears voltage values
  Dim i As Integer
  
  For i = 0 To UBound(m_nVoltage)
    m_nVoltage(i) = 0
  Next i
  pic.Cls
  picR.Cls
  shpMeter.Visible = True
  lblMeter.Visible = True
End Sub

'*********************************************** Form_Load
Private Sub Form_Load()
  
  'load standard resistor values
  m_nRes(0) = 100
  m_nRes(1) = 330
  m_nRes(2) = 470
  m_nRes(3) = 690
  m_nRes(4) = 1000
  m_nRes(5) = 2200
  m_nRes(6) = 4700
  m_nRes(7) = 10
  m_nRes(8) = 33
  m_nRes(9) = 47
  m_nRes(10) = 69
  m_nRes(11) = 160
  m_nRes(12) = 980
    
  'load standard voltage values
  m_nVolt(0) = 5
  m_nVolt(1) = 10
  m_nVolt(2) = 12
  m_nVolt(3) = 24
  m_nVolt(1) = 30
  m_nVolt(2) = 50
  m_nVolt(3) = 100
    
  m_nMode = NONE
  Randomize Timer
End Sub

'***************************************************** cmdExit
Private Sub cmdExit_Click()
  End
End Sub

'****************************************************** InsertFault
Private Sub InsertFault()
  Dim nRes As Integer
  Dim nType As Integer
  
  ClearFaults
  nRes = Rnd * 4 Mod 4 'creates 0, 1, 2 or 3...resistor to fault
  If nRes < 0 Then nRes = 0
  If nRes > 3 Then nRes = 3
  nType = Rnd * 2 Mod 2 'creates 0, 1
  If nType < 0 Then nType = 0
  If nType > 1 Then nType = 1
  m_sFault = "R" & CStr(nRes + 1) & " is " & IIf(nType = 0, "Opened", "Shorted")
  If nType = 0 Then
    m_uR(nRes).Fault = D_OPEN
  Else
    m_uR(nRes).Fault = D_SHORT
  End If
 'MsgBox m_sFault
 
End Sub

'****************************************************** ClearFaults
Private Sub ClearFaults()
  m_uR(0).Fault = D_OK
  m_uR(1).Fault = D_OK
  m_uR(2).Fault = D_OK
  m_uR(3).Fault = D_OK
  m_sFault = "No faults."
End Sub

'**************************************************** cmdSeries
Private Sub cmdSeries_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = SERIES
  
  shpMeter.Visible = True
  lblMeter.Visible = True
  
  'clears voltage values
  For i = 0 To UBound(m_nVoltage)
    m_nVoltage(i) = 0
  Next i
  
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
    
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  ClearFaults
  If m_bFault = True Then InsertFault
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value if RT
  m_nRT = 0
  For i = 0 To 3
    m_nRT = m_nRT + m_uR(i).value ' * m_uR(i).Fault)
  Next i
   
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  
  For i = 0 To 3
    m_uR(i).Voltage = m_uR(i).value * m_nIT
    m_uR(i).Current = m_nIT
  Next i
  
  'complete m_nVoltage table for meter reading
  m_nVoltage(V_SOURCE) = m_nVS
  m_nVoltage(V_A) = m_nVoltage(V_SOURCE)
  m_nVoltage(V_R1) = m_uR(0).Voltage
  m_nVoltage(V_B) = m_nVoltage(V_A) - m_nVoltage(V_R1)
  m_nVoltage(V_C) = m_nVoltage(V_B)
  m_nVoltage(V_R2) = m_uR(1).Voltage
  m_nVoltage(V_D) = m_nVoltage(V_C) - m_nVoltage(V_R2)
  m_nVoltage(V_R3) = m_uR(2).Voltage
  m_nVoltage(V_E) = m_nVoltage(V_D) - m_nVoltage(V_R3)
  m_nVoltage(V_R4) = m_uR(3).Voltage
  m_nVoltage(V_F) = m_nVoltage(V_E) - m_nVoltage(V_R4)
  m_nVoltage(V_G) = m_nVoltage(V_F)
  m_nVoltage(V_H) = m_nVoltage(V_G)
  m_nVoltage(V_COMMON) = 0
  
  'draw schematic circuit
  txtAns.Text = ""
  s = 1300: t = 1300
  pic.Cls
  DrawLineVert s, t
  DrawBattery s, t + 1300
  DrawLineVert s, t + 2600
  DrawLineHor s, t
  DrawResistorHor s + 1300, t
  DrawLineHor s + 2600, t
  DrawResistorVert s + 3900, t
  DrawResistorVert s + 3900, t + 1300
  DrawResistorVert s + 3900, t + 2600
  DrawLineHor s, t + 3900
  DrawLineHor s + 1300, t + 3900
  DrawLineHor s + 2600, t + 3900
  

  'annotate resistor values
  'pic.ForeColor = vbYellow
  pic.CurrentX = s + 1300
  pic.CurrentY = t - 500
  pic.Print "R1 = " & FormatResistance(m_uR(0).Resistor)
  pic.CurrentX = s + 4200
  pic.CurrentY = t + 500
  pic.Print "R2 = " & FormatResistance(m_uR(1).Resistor)
  pic.CurrentX = s + 4200
  pic.CurrentY = t + 1800
  pic.Print "R3 = " & FormatResistance(m_uR(2).Resistor)
  pic.CurrentX = s + 4200
  pic.CurrentY = t + 3100
  pic.Print "R4 = " & FormatResistance(m_uR(3).Resistor)
  pic.CurrentX = s + 500
  pic.CurrentY = t + 1800
  pic.Print "Vs = " & m_nVS & " V"

  'draw circuit board
  picR.Cls
  PCBDrawNode 1000, 200 'draw nodes first, then lines
  PCBDrawNode 2300, 200
  PCBDrawNode 3600, 200
  PCBDrawNode 4900, 200
  PCBDrawNode 4900, 1500
  PCBDrawNode 4900, 2800
  PCBDrawNode 4900, 4100
  PCBDrawNode 3600, 4100
  PCBDrawNode 2300, 4100
  PCBDrawNode 1000, 4100
  PCBDrawLineHor 1000, 200
  PCBDrawLineHor 3600, 200
  PCBDrawLineHor 1000, 4100
  PCBDrawLineHor 2300, 4100
  PCBDrawLineHor 3600, 4100
  PCBDrawResistorHor 2300, 200, m_uR(0).Resistor, 5
  PCBDrawResistorVert 4900, 200, m_uR(1).Resistor, 5
  PCBDrawResistorVert 4900, 1500, m_uR(2).Resistor, 5
  PCBDrawResistorVert 4900, 2800, m_uR(3).Resistor, 5
  picR.ForeColor = vbWhite
  picR.CurrentX = 750
  picR.CurrentY = 150
  picR.Print "+"
  picR.CurrentX = 750
  picR.CurrentY = 4000
  picR.Print "_"
  
End Sub

'**************************************************** cmdPar
Private Sub cmdPar_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = PARALLEL
  
  shpMeter.Visible = True
  lblMeter.Visible = True
   
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'clears voltage values
  For i = 0 To UBound(m_nVoltage)
    m_nVoltage(i) = 0
  Next i
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  ClearFaults
  If chkFault.value = vbChecked Then
    chkFault.value = vbUnchecked
    MsgBox "Faults disabled for this circuit.", vbOKOnly, "Comment"
    
  End If
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value if RT
  m_nRT = 0
  For i = 0 To 3
    m_nRT = m_nRT + 1 / m_uR(i).value '* m_uR(i).Fault)
  Next i
  If m_nRT = 0 Then
   MsgBox "Total Resistance = 0 Ohms"
   Exit Sub
  End If
  
  m_nRT = 1 / m_nRT
  
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  
  For i = 0 To 3
    m_uR(i).Voltage = m_nVS
    m_uR(i).Current = m_nVS / m_uR(i).value
  Next i
  
  'complete m_nVoltage table for meter reading
  m_nVoltage(V_SOURCE) = m_nVS
  m_nVoltage(V_A) = m_nVoltage(V_SOURCE)
  m_nVoltage(V_B) = m_nVoltage(V_A)
  m_nVoltage(V_C) = m_nVoltage(V_A)
  m_nVoltage(V_D) = m_nVoltage(V_A)
  m_nVoltage(V_K) = m_nVoltage(V_A)
  m_nVoltage(V_M) = m_nVoltage(V_A)
  m_nVoltage(V_I) = m_nVoltage(V_A)
  m_nVoltage(V_COMMON) = 0
  m_nVoltage(V_E) = m_nVoltage(V_COMMON)
  m_nVoltage(V_L) = m_nVoltage(V_E)
  m_nVoltage(V_N) = m_nVoltage(V_E)
  m_nVoltage(V_J) = m_nVoltage(V_E)
  m_nVoltage(V_F) = m_nVoltage(V_E)
  m_nVoltage(V_G) = m_nVoltage(V_E)
  m_nVoltage(V_H) = m_nVoltage(V_E)
  
  'draw circuit
  txtAns.Text = ""
  s = 1300: t = 1300
  pic.Cls
  'pic.ForeColor = vbGreen
  DrawLineVert s, t            'battery side
  DrawBattery s, t + 1300
  DrawLineVert s, t + 2600
  DrawLineHor s, t            'horizontal top
  DrawLineHor s + 1300, t
  DrawLineHor s + 2600, t
  DrawLineHor s + 3900, t
  DrawLineVert s + 1300, t
  DrawLineVert s + 2600, t
  DrawLineVert s + 3900, t
  DrawLineVert s + 5200, t
  DrawResistorVert s + 1300, t + 1300
  DrawResistorVert s + 2600, t + 1300
  DrawResistorVert s + 3900, t + 1300
  DrawResistorVert s + 5200, t + 1300
    
  DrawLineHor s, t + 3900       'horizontal bottom
  DrawLineHor s + 1300, t + 3900
  DrawLineHor s + 2600, t + 3900
  DrawLineHor s + 3900, t + 3900
  
  DrawLineVert s + 1300, t + 2600
  DrawLineVert s + 2600, t + 2600
  DrawLineVert s + 3900, t + 2600
  DrawLineVert s + 5200, t + 2600
  
  
  
  
  'annotate resistor values
 ' pic.ForeColor = vbYellow
  pic.CurrentX = s + 1500
  pic.CurrentY = t + 2300
  pic.Print "R1 = "
  pic.CurrentX = s + 1350
  pic.CurrentY = t + 2600
  pic.Print FormatResistance(m_uR(0).Resistor)
  
  pic.CurrentX = s + 2800
  pic.CurrentY = t + 2300
  pic.Print "R2 = "
  pic.CurrentX = s + 2650
  pic.CurrentY = t + 2600
  pic.Print FormatResistance(m_uR(1).Resistor)
  
  pic.CurrentX = s + 4100
  pic.CurrentY = t + 2300
  pic.Print "R3 = "
  pic.CurrentX = s + 3950
  pic.CurrentY = t + 2600
  pic.Print FormatResistance(m_uR(2).Resistor)
  
  pic.CurrentX = s + 5400
  pic.CurrentY = t + 2300
  pic.Print "R4 = "
  pic.CurrentX = s + 5250
  pic.CurrentY = t + 2600
  pic.Print FormatResistance(m_uR(3).Resistor)
  
  pic.CurrentX = s + 400
  pic.CurrentY = t + 1800
  pic.Print "Vs = "
  pic.CurrentX = s + 400
  pic.CurrentY = t + 2100
  pic.Print m_nVS & " V"

 'draw circuit board
  picR.Cls
  
  PCBDrawLineHor 1000, 200
  PCBDrawLineHor 2300, 200
  PCBDrawLineHor 3600, 200
  PCBDrawLineHor 1000, 4100
  PCBDrawLineHor 2300, 4100
  PCBDrawLineHor 3600, 4100
   PCBDrawLineHor 2300, 1500
  PCBDrawLineHor 3600, 1500
  PCBDrawLineHor 4900, 1500
  PCBDrawLineHor 2300, 2800
  PCBDrawLineHor 3600, 2800
  PCBDrawLineHor 4900, 2800
  PCBDrawNode 1000, 200 'draw nodes first, then lines
  PCBDrawNode 2300, 200
  PCBDrawNode 3600, 200
  PCBDrawNode 4900, 200
  PCBDrawNode 4900, 1500
  PCBDrawNode 4900, 2800
  PCBDrawNode 4900, 4100
  PCBDrawNode 3600, 4100
  PCBDrawNode 2300, 4100
  PCBDrawNode 1000, 4100
  PCBDrawNode 2300, 1500
  PCBDrawNode 3600, 1500
  PCBDrawNode 2300, 2800
  PCBDrawNode 3600, 2800
  PCBDrawNode 6200, 1500
  PCBDrawNode 6200, 2800
  
 
  
 
  
  
  
  PCBDrawLineVert 4900, 200
  PCBDrawLineVert 4900, 2800
  
  PCBDrawResistorVert 2300, 1500, m_uR(0).Resistor, 5
  PCBDrawResistorVert 3600, 1500, m_uR(1).Resistor, 5
  PCBDrawResistorVert 4900, 1500, m_uR(2).Resistor, 5
  PCBDrawResistorVert 6200, 1500, m_uR(3).Resistor, 5
  picR.ForeColor = vbWhite
  picR.CurrentX = 750
  picR.CurrentY = 150
  picR.Print "+"
  picR.CurrentX = 750
  picR.CurrentY = 4000
  picR.Print "_"
End Sub

'**************************************************** cmdCombo1
Private Sub cmdCombo1_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = COMBO1
  
  shpMeter.Visible = True
  lblMeter.Visible = True

  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'clears voltage values
  For i = 0 To UBound(m_nVoltage)
    m_nVoltage(i) = 0
  Next i
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  ClearFaults
  If m_bFault = True Then InsertFault
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
      
  'take into account of a shorted resistor in parallel to good resistor
  If m_uR(1).value = 0 Then
    m_uR(2).value = 0
  ElseIf m_uR(2).value = 0 Then
    m_uR(1).value = 0
  End If
       
  'calc value of RT
  m_nRT = 0
  m_nRT = m_uR(0).value + m_uR(3).value
  If m_uR(1).value > 0 Then
    m_nRT = m_nRT + (m_uR(1).value * m_uR(2).value) / (m_uR(1).value + m_uR(2).value)
  End If
  
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  m_uR(0).Current = m_nIT
  m_uR(0).Voltage = m_uR(0).Current * m_uR(0).value
  m_uR(3).Current = m_nIT
  m_uR(3).Voltage = m_uR(3).Current * m_uR(3).value
    
  m_uR(1).Voltage = m_nVS - (m_uR(0).Voltage + m_uR(3).Voltage)
  'MsgBox m_sFault
  
  '<<<<<<<<<<<<<<<<<<<<<<<<<<< fix this R2 and R3 are parallel
  If m_uR(1).value = 0 Then
    m_uR(1).Current = m_nIT
    m_uR(2).Current = 0
    m_uR(1).Voltage = 0
    m_uR(2).Voltage = 0
  ElseIf m_uR(2).value = 0 Then
    m_uR(2).Current = m_nIT
    m_uR(1).Current = 0
    m_uR(1).Voltage = 0
    m_uR(2).Voltage = 0
  Else
    m_uR(1).Current = m_uR(1).Voltage / m_uR(1).value
    m_uR(2).Current = m_uR(2).Voltage / m_uR(2).value
  End If
   
  
  'deals with opens and microvolt remnants
  For i = 0 To 3
    If m_uR(i).Voltage < 0.001 Then m_uR(i).Voltage = 0
  Next i
  
  'complete m_nVoltage table for meter reading
  m_nVoltage(V_SOURCE) = m_nVS
  m_nVoltage(V_A) = m_nVoltage(V_SOURCE)
  m_nVoltage(V_R1) = m_uR(0).Voltage
  m_nVoltage(V_B) = m_nVoltage(V_A) - m_nVoltage(V_R1)
  m_nVoltage(V_C) = m_nVoltage(V_B)
  m_nVoltage(V_D) = m_nVoltage(V_C)
  m_nVoltage(V_I) = m_nVoltage(V_D)
  m_nVoltage(V_R2) = m_uR(1).Voltage
  m_nVoltage(V_R3) = m_uR(2).Voltage
  m_nVoltage(V_E) = m_nVoltage(V_D) - m_nVoltage(V_R2)
  m_nVoltage(V_R4) = m_uR(3).Voltage
  m_nVoltage(V_F) = m_nVoltage(V_E)
  m_nVoltage(V_J) = m_nVoltage(V_E)
  m_nVoltage(V_G) = m_nVoltage(V_F)
  m_nVoltage(V_H) = m_nVoltage(V_G) - m_nVoltage(V_R4)
  m_nVoltage(V_COMMON) = 0
  
  'draw circuit
  txtAns.Text = ""
  s = 1300: t = 1300
  pic.Cls
  'pic.ForeColor = vbGreen
  DrawLineVert s, t
  DrawBattery s, t + 1300
  DrawLineVert s, t + 2600
  DrawLineHor s, t
  DrawResistorHor s + 1300, t
  DrawLineHor s + 2600, t
  DrawLineVert s + 3900, t
  DrawLineVert s + 3900, t + 2600
  
  DrawResistorVert s + 3250, t + 1300
  DrawResistorVert s + 4550, t + 1300
  DrawLineHor s + 3250, t + 1300
  DrawLineHor s + 3250, t + 2600
  
  DrawLineHor s, t + 3900
  DrawResistorHor s + 1300, t + 3900
  DrawLineHor s + 2600, t + 3900
  
  'annotate resistor values
  'pic.ForeColor = vbYellow
  pic.CurrentX = s + 1300
  pic.CurrentY = t - 500
  pic.Print "R1 = " & FormatResistance(m_uR(0).Resistor)
  
  pic.CurrentX = s + 3500
  pic.CurrentY = t + 1900
  pic.Print "R2 = "
  pic.CurrentX = s + 3350
  pic.CurrentY = t + 2200
  pic.Print FormatResistance(m_uR(1).Resistor)
  
  pic.CurrentX = s + 4800
  pic.CurrentY = t + 1900
  pic.Print "R3 = "
  pic.CurrentX = s + 4650
  pic.CurrentY = t + 2200
  pic.Print FormatResistance(m_uR(2).Resistor)
  
  pic.CurrentX = s + 1300
  pic.CurrentY = t + 3400
  pic.Print "R4 = " & FormatResistance(m_uR(3).Resistor)
  
  pic.CurrentX = s + 500
  pic.CurrentY = t + 1800
  pic.Print "Vs = " & m_nVS & " V"
    
  'draw circuit board
  picR.Cls
  PCBDrawNode 1000, 200 'draw nodes first, then lines
  PCBDrawNode 2300, 200
  PCBDrawNode 3600, 200
  PCBDrawNode 4900, 200
  PCBDrawNode 4900, 1500
  PCBDrawNode 6200, 1500
  PCBDrawNode 4900, 2800
  PCBDrawNode 6200, 2800
  PCBDrawNode 4900, 4100
  PCBDrawNode 3600, 4100
  PCBDrawNode 2300, 4100
  PCBDrawNode 1000, 4100
  PCBDrawLineHor 1000, 200
  PCBDrawLineHor 3600, 200
  PCBDrawLineHor 1000, 4100
  PCBDrawLineHor 4900, 1500
  PCBDrawLineHor 4900, 2800
  PCBDrawLineHor 3600, 4100
  PCBDrawLineVert 4900, 200
  PCBDrawLineVert 4900, 2800
  PCBDrawResistorHor 2300, 200, m_uR(0).Resistor, 5
  PCBDrawResistorVert 4900, 1500, m_uR(1).Resistor, 5
  PCBDrawResistorVert 6200, 1500, m_uR(2).Resistor, 5
  PCBDrawResistorHor 2300, 4100, m_uR(3).Resistor, 5
  picR.ForeColor = vbWhite
  picR.CurrentX = 750
  picR.CurrentY = 150
  picR.Print "+"
  picR.CurrentX = 750
  picR.CurrentY = 4000
  picR.Print "_"
  
End Sub

''**************************************************** cmdCombo2
Private Sub cmdCombo2_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = COMBO2
  
  shpMeter.Visible = True
  lblMeter.Visible = True

  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'clears voltage values
  For i = 0 To UBound(m_nVoltage)
    m_nVoltage(i) = 0
  Next i
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  ClearFaults
  If m_bFault = True Then InsertFault
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value of RT
  m_nRT = 0
  m_nRT = m_uR(0).value
  m_nRT = m_nRT + ((m_uR(1).value + m_uR(3).value) * m_uR(2).value) / (m_uR(1).value + m_uR(2).value + m_uR(3).value)
    
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  
  m_uR(0).Current = m_nIT
  m_uR(0).Voltage = m_uR(0).Current * m_uR(0).value
  
    '<<<<<<<<<<<<<<<<<<<<<<<<<<< fix this R2 and R3 are parallel
  If m_uR(2).value = 0 Then
    m_uR(2).Current = m_nIT
    m_uR(1).Current = 0
    m_uR(3).Current = 0
    m_uR(1).Voltage = 0
    m_uR(2).Voltage = 0
    m_uR(3).Voltage = 0
  Else
    m_uR(2).Voltage = m_nVS - m_uR(0).Voltage
    m_uR(2).Current = m_uR(2).Voltage / m_uR(2).value
    m_uR(1).Current = m_uR(0).Current - m_uR(2).Current
    m_uR(3).Current = m_uR(1).Current
    m_uR(1).Voltage = m_uR(1).Current * m_uR(1).value
    m_uR(3).Voltage = m_uR(3).Current * m_uR(3).value
  End If
    
  'complete m_nVoltage table for meter reading
  m_nVoltage(V_SOURCE) = m_nVS
  m_nVoltage(V_A) = m_nVoltage(V_SOURCE)
  m_nVoltage(V_R1) = m_uR(0).Voltage
  m_nVoltage(V_B) = m_nVoltage(V_A) - m_nVoltage(V_R1)
  m_nVoltage(V_C) = m_nVoltage(V_B)
  m_nVoltage(V_O) = m_nVoltage(V_C)
  m_nVoltage(V_I) = m_nVoltage(V_D)
  m_nVoltage(V_R2) = m_uR(1).Voltage
  m_nVoltage(V_R3) = m_uR(2).Voltage
  m_nVoltage(V_D) = m_nVoltage(V_C) - m_nVoltage(V_R2)
  m_nVoltage(V_R4) = m_uR(3).Voltage
  m_nVoltage(V_E) = m_nVoltage(V_D) - m_nVoltage(V_R4)
  m_nVoltage(V_COMMON) = 0
  m_nVoltage(V_I) = m_nVoltage(V_COMMON)
  m_nVoltage(V_J) = m_nVoltage(V_COMMON)
  m_nVoltage(V_E) = m_nVoltage(V_COMMON)
  m_nVoltage(V_F) = m_nVoltage(V_COMMON)
  m_nVoltage(V_G) = m_nVoltage(V_COMMON)
  m_nVoltage(V_H) = m_nVoltage(V_COMMON)
    
    
  'draw circuit
  txtAns.Text = ""
  s = 1300: t = 1300
  pic.Cls
  'pic.ForeColor = vbGreen
  DrawLineVert s, t
  DrawBattery s, t + 1300
  DrawLineVert s, t + 2600
  DrawLineHor s, t
  DrawResistorHor s + 1300, t
  DrawLineHor s + 2600, t
  DrawLineHor s + 3250, t
  DrawLineVert s + 3900, t + 2600
  DrawLineVert s + 4550, t + 1300
  DrawResistorVert s + 3250, t + 1300
  DrawLineHor s + 3250, t + 2600
  DrawResistorVert s + 3250, t
  DrawResistorVert s + 4550, t
  DrawLineHor s, t + 3900
  DrawLineHor s + 1300, t + 3900
  DrawLineHor s + 2600, t + 3900
  
  'annotate resistor values
  'pic.ForeColor = vbYellow
  pic.CurrentX = s + 1300
  pic.CurrentY = t - 500
  pic.Print "R1 = " & FormatResistance(m_uR(0).Resistor)
  
  pic.CurrentX = s + 3500
  pic.CurrentY = t + 600
  pic.Print "R2 = "
  pic.CurrentX = s + 3350
  pic.CurrentY = t + 900
  pic.Print FormatResistance(m_uR(1).Resistor)
  
  pic.CurrentX = s + 4800
  pic.CurrentY = t + 600
  pic.Print "R3 = "
  pic.CurrentX = s + 4650
  pic.CurrentY = t + 900
  pic.Print FormatResistance(m_uR(2).Resistor)
  
  pic.CurrentX = s + 3500
  pic.CurrentY = t + 1900
  pic.Print "R4 = "
  pic.CurrentX = s + 3350
  pic.CurrentY = t + 2200
  pic.Print FormatResistance(m_uR(3).Resistor)
 
  pic.CurrentX = s + 500
  pic.CurrentY = t + 1800
  pic.Print "Vs = " & m_nVS & " V"
  
   'draw circuit board
  picR.Cls
  PCBDrawNode 1000, 200 'draw nodes first, then lines
  PCBDrawNode 2300, 200
  PCBDrawNode 3600, 200
  PCBDrawNode 4900, 200
  PCBDrawNode 4900, 1500
  PCBDrawNode 6200, 1500
  PCBDrawNode 4900, 2800
  PCBDrawNode 6200, 200
  PCBDrawNode 4900, 4100
  PCBDrawNode 3600, 4100
  PCBDrawNode 2300, 4100
  PCBDrawNode 1000, 4100
  PCBDrawNode 6200, 2800
  PCBDrawLineHor 1000, 200
  PCBDrawLineHor 3600, 200
  PCBDrawLineHor 1000, 4100
  PCBDrawLineHor 4900, 200
  PCBDrawLineHor 4900, 2800
  PCBDrawLineHor 3600, 4100
  PCBDrawLineHor 2300, 4100
  PCBDrawLineVert 4900, 2800
  PCBDrawLineVert 6200, 1500
  PCBDrawResistorHor 2300, 200, m_uR(0).Resistor, 5
  PCBDrawResistorVert 4900, 200, m_uR(1).Resistor, 5
  PCBDrawResistorVert 6200, 200, m_uR(2).Resistor, 5
  PCBDrawResistorVert 4900, 1500, m_uR(3).Resistor, 5
  picR.ForeColor = vbWhite
  picR.CurrentX = 750
  picR.CurrentY = 150
  picR.Print "+"
  picR.CurrentX = 750
  picR.CurrentY = 4000
  picR.Print "_"
End Sub

'************************************************************* cmdShow
Private Sub cmdShow_Click()
  If m_nMode = NONE Then Exit Sub
  
  txtAns.Text = ""
  txtAns.Text = txtAns.Text & "Voltage Source: " & m_nVS & "V" & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "Total Resistance: " & IIf(m_nRT < 1000000, FormatResistance(m_nRT), "Infinite") & vbCrLf
  txtAns.Text = txtAns.Text & "Total Current: " & IIf(m_nRT < 1000000, FormatCurrent(m_nIT), "0 A") & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "VR1: " & IIf(m_uR(0).Voltage >= 0.001, FormatVoltage(m_uR(0).Voltage), "0 V") & vbCrLf
  txtAns.Text = txtAns.Text & "VR2: " & IIf(m_uR(1).Voltage >= 0.001, FormatVoltage(m_uR(1).Voltage), "0 V") & vbCrLf
  txtAns.Text = txtAns.Text & "VR3: " & IIf(m_uR(2).Voltage >= 0.001, FormatVoltage(m_uR(2).Voltage), "0 V") & vbCrLf
  txtAns.Text = txtAns.Text & "VR4: " & IIf(m_uR(3).Voltage >= 0.001, FormatVoltage(m_uR(3).Voltage), "0 V") & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "IR1: " & IIf(m_uR(0).Current >= 0.000001, FormatCurrent(m_uR(0).Current), "0 A") & vbCrLf
  txtAns.Text = txtAns.Text & "IR2: " & IIf(m_uR(1).Current >= 0.000001, FormatCurrent(m_uR(1).Current), "0 A") & vbCrLf
  txtAns.Text = txtAns.Text & "IR3: " & IIf(m_uR(2).Current >= 0.000001, FormatCurrent(m_uR(2).Current), "0 A") & vbCrLf
  txtAns.Text = txtAns.Text & "IR4: " & IIf(m_uR(3).Current >= 0.000001, FormatCurrent(m_uR(3).Current), "0 A") & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "Fault: " & m_sFault & vbCrLf
End Sub

'**************** P C B   D R A W I N G   P R O C E D U R E S ************

'**************************************************** PCBDrawNode
'draws a node on PCB (printed circuit board
Private Sub PCBDrawNode(X As Single, Y As Single)
  
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 105
  
  'soldered node
  picR.ForeColor = RGB(150, 150, 150)
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 65
  picR.ForeColor = vbWhite
  picR.Circle (X, Y), 55
  'picR.ForeColor = RGB(220, 220, 220)
  picR.ForeColor = vbWhite
  picR.FillColor = vbWhite
  picR.Circle (X + 10, Y + 10), 5
  
End Sub

'***************************************************** PCBDrawLineHor
Private Sub PCBDrawLineHor(ByVal X As Single, ByVal Y As Single)
   
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Line (X + 50, Y - 50)-(X + 1250, Y + 50), , BF
  
  'wiring
  'picR.ForeColor = RGB(150, 150, 150)
  picR.ForeColor = RGB(201, 196, 101)
  picR.FillColor = picR.ForeColor
  picR.Line (X + 50, Y - 25)-(X + 1250, Y + 25), , BF
  'picR.ForeColor = RGB(220, 220, 220)
  'picR.ForeColor = vbWhite
  'picR.FillColor = picR.ForeColor
  'picR.Line (x + 50, y - 15)-(x + 1250, y + 15), , BF
 
End Sub

'***************************************************** PCBDrawLineVert
Private Sub PCBDrawLineVert(ByVal X As Single, ByVal Y As Single)
  
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Line (X - 50, Y + 50)-(X + 50, Y + 1250), , BF
  
  'wiring
  'picR.ForeColor = RGB(150, 150, 150)
  picR.ForeColor = RGB(201, 196, 101)
  picR.FillColor = picR.ForeColor
  picR.Line (X - 25, Y + 50)-(X + 25, Y + 1250), , BF
  'picR.ForeColor = RGB(220, 220, 220)
  'picR.ForeColor = vbWhite
  'picR.FillColor = picR.ForeColor
  'picR.Line (x - 15, y + 50)-(x + 15, y + 1250), , BF
 
End Sub

'***************************************************** PCBDrawResistorHor
Private Sub PCBDrawResistorHor(ByVal X As Single, ByVal Y As Single, ByVal value As Single, tol As Single)
  Dim uBand As COLOR_BAND
 
  'wiring
  picR.ForeColor = vbBlack             'black layer
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 25               'black end points of wire
  picR.Circle (X + 1300, Y), 25
  picR.Line (X, Y - 35)-(X + 1300, Y + 35), , BF 'black line
  picR.ForeColor = RGB(208, 223, 221)
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 15                 'smaller silver points of wire
  picR.Circle (X + 1300, Y), 15
  picR.Line (X, Y - 5)-(X + 1300, Y + 5), , BF  'silver line
  
  'resistor body
  picR.ForeColor = vbBlack                 'black layer
  picR.Line (X + 350, Y - 80)-(X + 950, Y + 80), , BF
  
  picR.ForeColor = RGB(160, 91, 19)               '
  picR.FillColor = picR.ForeColor
  picR.Line (X + 370, Y - 70)-(X + 930, Y + 70), , BF '
  
  'add color bands
  uBand = GetColorBands(value, tol)
  picR.ForeColor = uBand.First
  picR.Line (X + 400, Y - 70)-(X + 450, Y + 70), , BF
  picR.ForeColor = uBand.Second
  picR.Line (X + 500, Y - 70)-(X + 550, Y + 70), , BF
  picR.ForeColor = uBand.Third
  picR.Line (X + 600, Y - 70)-(X + 650, Y + 70), , BF
  picR.ForeColor = uBand.Fourth
  picR.Line (X + 800, Y - 70)-(X + 850, Y + 70), , BF
  
  
End Sub

'***************************************************** GetColorBands
Private Function GetColorBands(nValue As Single, nTol As Single) As COLOR_BAND
  Dim uBand As COLOR_BAND
  Dim sVal As String
  
  sVal = CStr(nValue)
  
  'gets color for first band
  Select Case Mid(sVal, 1, 1)
    Case "0"
      uBand.First = vbBlack
    Case "1" 'brown
      uBand.First = RGB(194, 146, 92)
    Case "2"
      uBand.First = vbRed
    Case "3" 'orange
      uBand.First = RGB(255, 151, 63)
    Case "4"
      uBand.First = vbYellow
    Case "5"
      uBand.First = vbGreen
    Case "6"
      uBand.First = vbBlue
    Case "7"
      uBand.First = vbMagenta
    Case "8" 'gray
      uBand.First = RGB(139, 155, 160)
    Case "9"
      uBand.First = vbWhite
  End Select
  
  'gets color for second band
  Select Case Mid(sVal, 2, 1)
    Case "0"
      uBand.Second = vbBlack
    Case "1"
      uBand.Second = RGB(194, 146, 92)
    Case "2"
      uBand.Second = vbRed
    Case "3"
      uBand.Second = RGB(255, 151, 63)
    Case "4"
      uBand.Second = vbYellow
    Case "5"
      uBand.Second = vbGreen
    Case "6"
      uBand.Second = vbBlue
    Case "7"
      uBand.Second = vbMagenta
    Case "8"
      uBand.Second = RGB(139, 155, 160)
    Case "9"
      uBand.Second = vbWhite
  End Select

  'gets color for third band
  Select Case Len(sVal) - 2
    Case 0
      uBand.Third = vbBlack
    Case 1
      uBand.Third = RGB(194, 146, 92)
    Case 2
      uBand.Third = vbRed
    Case 3
      uBand.Third = RGB(255, 151, 63)
  End Select
  
  'tolerance color
  Select Case nTol
    Case 2 'red
      uBand.Fourth = vbRed
    Case 5 'gold
      uBand.Fourth = RGB(232, 207, 68)
    Case 10 'silver
      uBand.Fourth = RGB(208, 223, 221)
    Case 20 'no band
      uBand.Fourth = RGB(160, 91, 19)
  End Select

  GetColorBands = uBand

End Function


'***************************************************** PCBDrawResistorVert
Private Sub PCBDrawResistorVert(ByVal X As Single, ByVal Y As Single, ByVal value As Single, tol As Single)
  Dim uBand As COLOR_BAND

  'wiring
  picR.ForeColor = vbBlack             'black layer
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 25               'black end points of wire
  picR.Circle (X, Y + 1300), 25
  picR.Line (X - 35, Y)-(X + 35, Y + 1300), , BF 'black line
  picR.ForeColor = RGB(208, 223, 221)
  picR.FillColor = picR.ForeColor
  picR.Circle (X, Y), 15                 'smaller silver points of wire
  picR.Circle (X, Y + 1300), 15
  picR.Line (X - 5, Y)-(X + 5, Y + 1300), , BF 'silver line
 
  'resistor body
  picR.ForeColor = vbBlack                 'black layer
  picR.Line (X - 80, Y + 350)-(X + 80, Y + 950), , BF
  
  picR.ForeColor = RGB(160, 91, 19)               'top silver layer
  picR.FillColor = picR.ForeColor
  picR.Line (X - 70, Y + 370)-(X + 70, Y + 930), , BF 'silver line for wire
  
  'add color bands
  uBand = GetColorBands(value, tol)
  picR.ForeColor = uBand.First
  picR.Line (X - 70, Y + 400)-(X + 70, Y + 450), , BF
  picR.ForeColor = uBand.Second
  picR.Line (X - 70, Y + 500)-(X + 70, Y + 550), , BF
  picR.ForeColor = uBand.Third
  picR.Line (X - 70, Y + 600)-(X + 70, Y + 650), , BF
  picR.ForeColor = uBand.Fourth
  picR.Line (X - 70, Y + 800)-(X + 70, Y + 850), , BF
  
End Sub


'*********  S C H E M A T I C   D R A W I N G   P R O C E D U R E S ********
'**************************************************** DrawNode
'various drawing functions of components, lines and nodes
Private Sub DrawNode(X As Single, Y As Single)
  'pic.FillColor = vbBlack
  pic.Circle (X, Y), 50
End Sub

'**************************************************** DrawResistorHor
Private Sub DrawResistorHor(ByVal X As Single, ByVal Y As Single)
  'DrawNode x, y
  pic.Line (X, Y)-(X + 350, Y)
  pic.Line -(X + 400, Y - 200)
  pic.Line -(X + 500, Y + 200)
  pic.Line -(X + 600, Y - 200)
  pic.Line -(X + 700, Y + 200)
  pic.Line -(X + 800, Y - 200)
  pic.Line -(X + 900, Y + 200)
  pic.Line -(X + 950, Y)
  pic.Line -(X + 1300, Y)
  'DrawNode x + 1300, y
End Sub

'**************************************************** DrawResistorVert
Private Sub DrawResistorVert(ByVal X As Single, ByVal Y As Single)
  'DrawNode x, y
  pic.Line (X, Y)-(X, Y + 350)
  pic.Line -(X + 200, Y + 400)
  pic.Line -(X - 200, Y + 500)
  pic.Line -(X + 200, Y + 600)
  pic.Line -(X - 200, Y + 700)
  pic.Line -(X + 200, Y + 800)
  pic.Line -(X - 200, Y + 900)
  pic.Line -(X, Y + 950)
  pic.Line -(X, Y + 1300)
  'DrawNode x, y + 1300
End Sub

'**************************************************** DrawGrid
'draws grid
Private Sub DrawGrid()
  Dim i, j As Single
  pic.DrawWidth = 2
  For i = 0 To 10
    For j = 0 To 10
      pic.PSet (i * 1300, j * 1300)
    Next j
  Next i
  pic.DrawWidth = 1
End Sub

'**************************************************** DrawLineHor
Private Sub DrawLineHor(ByVal X As Single, ByVal Y As Single)
  pic.Line (X, Y)-(X + 1300, Y)
End Sub

'**************************************************** DrawLineVert
Private Sub DrawLineVert(ByVal X As Single, ByVal Y As Single)
  pic.Line (X, Y)-(X, Y + 1300)
End Sub

'**************************************************** DrawBattery
Private Sub DrawBattery(ByVal X As Single, ByVal Y As Single)
  'DrawNode x, y
  pic.Line (X, Y)-(X, Y + 350)
  pic.Line (X - 200, Y + 350)-(X + 200, Y + 350)
  pic.Line (X - 100, Y + 550)-(X + 100, Y + 550)
  pic.Line (X - 200, Y + 750)-(X + 200, Y + 750)
  pic.Line (X - 100, Y + 950)-(X + 100, Y + 950)
  pic.Line (X, Y + 950)-(X, Y + 1300)
  
  pic.Line (X + 200, Y + 200)-(X + 400, Y + 200)  'Plus (+) sign
  pic.Line (X + 300, Y + 100)-(X + 300, Y + 300)
  
  pic.Line (X + 200, Y + 1100)-(X + 400, Y + 1100)  'Negative (-) sign
  'DrawNode x, y + 1300
  
End Sub

'***************************************************** FormatResistance
Private Function FormatResistance(nVal As Single) As String
  Dim sOut As String
    
  If nVal < 1000 Then
    sOut = CStr(Format(nVal, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = Left(sOut, Len(sOut) - 1) & " Ohms"
    Else
      sOut = sOut & " Ohms"
    End If
  Else
    sOut = CStr(Format(nVal / 1000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = Left(sOut, Len(sOut) - 1) & "K Ohms"
    Else
      sOut = sOut & "K Ohms"
    End If
  End If
  FormatResistance = sOut
End Function

'***************************************************** FormatVoltage
Private Function FormatVoltage(nVal As Single) As String
  Dim sOut As String
  If nVal < 0.001 Then
    sOut = CStr(Format(nVal * 1000000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = sOut & "0 µV"
    Else
      sOut = sOut & " µV"
    End If
  ElseIf nVal < 1 Then
    sOut = CStr(Format(nVal * 1000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = sOut & "0 mV"
    Else
      sOut = sOut & " mV"
    End If
  Else
    sOut = CStr(Format(nVal, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = sOut & "0 V"
    Else
      sOut = sOut & " V"
    End If
  End If
  FormatVoltage = sOut
End Function

'*************************************************** FormatCurrent
Private Function FormatCurrent(nVal As Single) As String
  Dim sOut As String
    
  If nVal < 0.000001 Then
    sOut = CStr(Format(nVal * 1000000000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = Left(sOut, Len(sOut) - 1) & " nA"
    Else
      sOut = sOut & " nA"
    End If
  
  ElseIf nVal < 0.001 Then
    sOut = CStr(Format(nVal * 1000000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = Left(sOut, Len(sOut) - 1) & " µA"
    Else
      sOut = sOut & " µA"
    End If
      
  ElseIf nVal < 1 Then
    sOut = CStr(Format(nVal * 1000, "##.###"))
    If Right(sOut, 1) = "." Then
      sOut = Left(sOut, Len(sOut) - 1) & " mA"
    Else
      sOut = sOut & " mA"
    End If
  Else
    sOut = CStr(Format(nVal, "##.###")) & " A"
  End If
  FormatCurrent = sOut
End Function

Private Sub picR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim nVolt As Single
  
  nVolt = 0
  
  If Y > 125 And Y < 275 Then
    If X > 925 And X < 1075 Then nVolt = IIf(m_nVoltage(V_A) >= 0.001, m_nVoltage(V_A), 0)
    If X > 2225 And X < 2375 Then nVolt = IIf(m_nVoltage(V_A) >= 0.001, m_nVoltage(V_A), 0)
    If X > 3525 And X < 3675 Then nVolt = IIf(m_nVoltage(V_B) >= 0.001, m_nVoltage(V_B), 0)
    If X > 4825 And X < 4975 Then nVolt = IIf(m_nVoltage(V_C) >= 0.001, m_nVoltage(V_C), 0)
    If X > 6125 And X < 6275 Then nVolt = IIf(m_nVoltage(V_O) >= 0.001, m_nVoltage(V_O), 0)
  End If
  If Y > 1425 And Y < 1575 Then
    If X > 2225 And X < 2375 Then nVolt = IIf(m_nVoltage(V_K) >= 0.001, m_nVoltage(V_K), 0)
    If X > 3525 And X < 3675 Then nVolt = IIf(m_nVoltage(V_M) >= 0.001, m_nVoltage(V_M), 0)
    If X > 4825 And X < 4975 Then nVolt = IIf(m_nVoltage(V_D) >= 0.001, m_nVoltage(V_D), 0)
    If X > 6125 And X < 6275 Then nVolt = IIf(m_nVoltage(V_I) >= 0.001, m_nVoltage(V_I), 0)
  End If
  If Y > 2725 And Y < 2875 Then
    If X > 2225 And X < 2375 Then nVolt = IIf(m_nVoltage(V_L) >= 0.001, m_nVoltage(V_L), 0)
    If X > 3525 And X < 3675 Then nVolt = IIf(m_nVoltage(V_N) >= 0.001, m_nVoltage(V_N), 0)
    If X > 4825 And X < 4975 Then nVolt = IIf(m_nVoltage(V_E) >= 0.001, m_nVoltage(V_E), 0)
    If X > 6125 And X < 6275 Then nVolt = IIf(m_nVoltage(V_J) >= 0.001, m_nVoltage(V_J), 0)
  End If
  If Y > 4025 And Y < 4175 Then
    If X > 925 And X < 1075 Then nVolt = m_nVoltage(V_COMMON)
    If X > 2225 And X < 2375 Then nVolt = IIf(m_nVoltage(V_H) >= 0.001, m_nVoltage(V_H), 0)
    If X > 3525 And X < 3675 Then nVolt = IIf(m_nVoltage(V_G) >= 0.001, m_nVoltage(V_G), 0)
    If X > 4825 And X < 4975 Then nVolt = IIf(m_nVoltage(V_F) >= 0.001, m_nVoltage(V_F), 0)
  End If
  
  If nVolt >= 0.001 Then
    lblMeter.Caption = FormatVoltage(nVolt)
  Else
    lblMeter.Caption = "0 V"
  End If

End Sub
