VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resistor Problem Solver v0.4 by C. Bolin, October 6, 2004"
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
      ScaleHeight     =   4455
      ScaleWidth      =   7665
      TabIndex        =   9
      Top             =   5520
      Width           =   7725
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

'module variables
Private m_nMode As RES_MODE
Private m_nRes(12) As Single  'stores standard resistor values
Private m_nVolt(6) As Single  'stores standard voltage values
Private m_uR(3) As Resistor   '0-3...stores all specific info for resistors
Private m_nRT As Single 'total parameters
Private m_nVS As Single
Private m_nIT As Single
Private m_bFault As Boolean

'************************************************ chkFault
'select if program should add fault
Private Sub chkFault_Click()
  If chkFault.value = vbChecked Then
    m_bFault = True
  Else
    m_bFault = False
  End If
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

'**************************************************** cmdSeries
Private Sub cmdSeries_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = SERIES
  
  's = (x \ 1300) * 1300
  't = (y \ 1300) * 1300
  
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  m_uR(0).Fault = D_OK
  m_uR(1).Fault = D_OK
  m_uR(2).Fault = D_OK
  m_uR(3).Fault = D_OK
    
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
  PCBDrawLineHor 3600, 4100
  
  PCBDrawLineVert 4900, 1500
  PCBDrawResistorHor 2300, 200, m_uR(0).Resistor, 5
  PCBDrawResistorHor 2300, 4100, m_uR(3).Resistor, 5
  
  
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

End Sub

'**************************************************** cmdPar
Private Sub cmdPar_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = PARALLEL
  
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  m_uR(0).Fault = D_OK
  m_uR(1).Fault = D_OK
  m_uR(2).Fault = D_OK
  m_uR(3).Fault = D_OK
    
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

End Sub

'**************************************************** cmdCombo1
Private Sub cmdCombo1_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = COMBO1
  
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  m_uR(0).Fault = D_OK
  m_uR(1).Fault = D_OK
  m_uR(2).Fault = D_OK
  m_uR(3).Fault = D_OK
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value of RT
  m_nRT = 0
  m_nRT = m_uR(0).value + m_uR(3).value
  m_nRT = m_nRT + (m_uR(1).value * m_uR(2).value) / (m_uR(1).value + m_uR(2).value)
    
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  m_uR(0).Current = m_nIT
  m_uR(0).Voltage = m_uR(0).Current * m_uR(0).value
  m_uR(3).Current = m_nIT
  m_uR(3).Voltage = m_uR(3).Current * m_uR(3).value
    
  m_uR(1).Voltage = m_nVS - (m_uR(0).Voltage + m_uR(3).Voltage)
  m_uR(1).Current = m_uR(1).Voltage / m_uR(1).value
  m_uR(2).Voltage = m_uR(1).Voltage
  m_uR(2).Current = m_uR(2).Voltage / m_uR(2).value
  
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
End Sub

''**************************************************** cmdCombo2
Private Sub cmdCombo2_Click()
  Dim s, t As Single
  Dim i As Integer
  m_nMode = COMBO2
  
  'load resistor values for circuit
  m_uR(0).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(1).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(2).Resistor = m_nRes(Rnd * 6 Mod 6)
  m_uR(3).Resistor = m_nRes(Rnd * 6 Mod 6)
  
  'load source voltage
  m_nVS = m_nVolt(Rnd * 3 Mod 3)
  
  'consider faults
  m_uR(0).Fault = D_OK
  m_uR(1).Fault = D_OK
  m_uR(2).Fault = D_OK
  m_uR(3).Fault = D_OK
    
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
  
  m_uR(2).Voltage = m_nVS - m_uR(0).Voltage
  m_uR(2).Current = m_uR(2).Voltage / m_uR(2).value
  
  m_uR(1).Current = m_uR(0).Current - m_uR(2).Current
  m_uR(3).Current = m_uR(1).Current
  
  m_uR(1).Voltage = m_uR(1).Current * m_uR(1).value
  m_uR(3).Voltage = m_uR(3).Current * m_uR(3).value
    
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
End Sub

'************************************************************* cmdShow
Private Sub cmdShow_Click()
  If m_nMode = NONE Then Exit Sub
  
  txtAns.Text = ""
  txtAns.Text = txtAns.Text & "Voltage Source: " & m_nVS & "V" & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "Total Resistance: " & FormatResistance(m_nRT) & vbCrLf
  txtAns.Text = txtAns.Text & "Total Current: " & FormatCurrent(m_nIT) & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "VR1: " & FormatVoltage(m_uR(0).Voltage) & vbCrLf
  txtAns.Text = txtAns.Text & "VR2: " & FormatVoltage(m_uR(1).Voltage) & vbCrLf
  txtAns.Text = txtAns.Text & "VR3: " & FormatVoltage(m_uR(2).Voltage) & vbCrLf
  txtAns.Text = txtAns.Text & "VR4: " & FormatVoltage(m_uR(3).Voltage) & vbCrLf & vbCrLf
  txtAns.Text = txtAns.Text & "IR1: " & FormatCurrent(m_uR(0).Current) & vbCrLf
  txtAns.Text = txtAns.Text & "IR2: " & FormatCurrent(m_uR(1).Current) & vbCrLf
  txtAns.Text = txtAns.Text & "IR3: " & FormatCurrent(m_uR(2).Current) & vbCrLf
  txtAns.Text = txtAns.Text & "IR4: " & FormatCurrent(m_uR(3).Current) & vbCrLf & vbCrLf
End Sub

'**************** P C B   D R A W I N G   P R O C E D U R E S ************

'**************************************************** PCBDrawNode
'draws a node on PCB (printed circuit board
Private Sub PCBDrawNode(x As Single, y As Single)
  
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 105
  
  'soldered node
  picR.ForeColor = RGB(150, 150, 150)
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 65
  picR.ForeColor = vbWhite
  picR.Circle (x, y), 55
  'picR.ForeColor = RGB(220, 220, 220)
  picR.ForeColor = vbWhite
  picR.FillColor = vbWhite
  picR.Circle (x + 10, y + 10), 5
  
End Sub

'***************************************************** PCBDrawLineHor
Private Sub PCBDrawLineHor(ByVal x As Single, ByVal y As Single)
   
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Line (x + 50, y - 50)-(x + 1250, y + 50), , BF
  
  'wiring
  picR.ForeColor = RGB(150, 150, 150)
  picR.FillColor = picR.ForeColor
  picR.Line (x + 50, y - 25)-(x + 1250, y + 25), , BF
  picR.ForeColor = RGB(220, 220, 220)
  picR.ForeColor = vbWhite
  picR.FillColor = picR.ForeColor
  picR.Line (x + 50, y - 15)-(x + 1250, y + 15), , BF
 
End Sub

'***************************************************** PCBDrawLineVert
Private Sub PCBDrawLineVert(ByVal x As Single, ByVal y As Single)
  
  'shadow
  picR.ForeColor = RGB(0, 80, 0)
  picR.FillColor = picR.ForeColor
  picR.Line (x - 50, y + 50)-(x + 50, y + 1250), , BF
  
  'wiring
  picR.ForeColor = RGB(150, 150, 150)
  picR.FillColor = picR.ForeColor
  picR.Line (x - 25, y + 50)-(x + 25, y + 1250), , BF
  picR.ForeColor = RGB(220, 220, 220)
  picR.ForeColor = vbWhite
  picR.FillColor = picR.ForeColor
  picR.Line (x - 15, y + 50)-(x + 15, y + 1250), , BF
 
End Sub

'***************************************************** PCBDrawResistorHor
Private Sub PCBDrawResistorHor(ByVal x As Single, ByVal y As Single, ByVal value As Single, tol As Single)
   
 
  'wiring
  picR.ForeColor = RGB(150, 150, 150)
  picR.ForeColor = vbBlack
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 25
  picR.Circle (x + 1300, y), 25
  picR.Line (x + 50, y - 35)-(x + 1300, y + 35), , BF
  picR.ForeColor = RGB(240, 240, 240)
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 15
  picR.Circle (x + 1300, y), 15
  picR.Line (x, y - 5)-(x + 1300, y + 5), , BF
 
  'resistor body
  picR.ForeColor = RGB(160, 91, 19)
  picR.FillColor = picR.ForeColor
  picR.Line (x + 350, y - 80)-(x + 950, y + 80), , BF
  
End Sub

'***************************************************** PCBDrawResistorVert
Private Sub PCBDrawResistorVert(ByVal x As Single, ByVal y As Single, ByVal value As Single, tol As Single)
 
  'wiring
  picR.ForeColor = RGB(150, 150, 150)
  picR.ForeColor = vbBlack
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 25
  picR.Circle (x + 1300, y), 25
  picR.Line (x + 50, y - 35)-(x + 1300, y + 35), , BF
  picR.ForeColor = RGB(240, 240, 240)
  picR.FillColor = picR.ForeColor
  picR.Circle (x, y), 15
  picR.Circle (x + 1300, y), 15
  picR.Line (x, y - 5)-(x + 1300, y + 5), , BF
 
  'resistor body
  picR.ForeColor = RGB(160, 91, 19)
  picR.FillColor = picR.ForeColor
  picR.Line (x + 350, y - 80)-(x + 950, y + 80), , BF
  
End Sub


'*********  S C H E M A T I C   D R A W I N G   P R O C E D U R E S ********
'**************************************************** DrawNode
'various drawing functions of components, lines and nodes
Private Sub DrawNode(x As Single, y As Single)
  'pic.FillColor = vbBlack
  pic.Circle (x, y), 50
End Sub

'**************************************************** DrawResistorHor
Private Sub DrawResistorHor(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  pic.Line (x, y)-(x + 350, y)
  pic.Line -(x + 400, y - 200)
  pic.Line -(x + 500, y + 200)
  pic.Line -(x + 600, y - 200)
  pic.Line -(x + 700, y + 200)
  pic.Line -(x + 800, y - 200)
  pic.Line -(x + 900, y + 200)
  pic.Line -(x + 950, y)
  pic.Line -(x + 1300, y)
  'DrawNode x + 1300, y
End Sub

'**************************************************** DrawResistorVert
Private Sub DrawResistorVert(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  pic.Line (x, y)-(x, y + 350)
  pic.Line -(x + 200, y + 400)
  pic.Line -(x - 200, y + 500)
  pic.Line -(x + 200, y + 600)
  pic.Line -(x - 200, y + 700)
  pic.Line -(x + 200, y + 800)
  pic.Line -(x - 200, y + 900)
  pic.Line -(x, y + 950)
  pic.Line -(x, y + 1300)
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
Private Sub DrawLineHor(ByVal x As Single, ByVal y As Single)
  pic.Line (x, y)-(x + 1300, y)
End Sub

'**************************************************** DrawLineVert
Private Sub DrawLineVert(ByVal x As Single, ByVal y As Single)
  pic.Line (x, y)-(x, y + 1300)
End Sub

'**************************************************** DrawBattery
Private Sub DrawBattery(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  pic.Line (x, y)-(x, y + 350)
  pic.Line (x - 200, y + 350)-(x + 200, y + 350)
  pic.Line (x - 100, y + 550)-(x + 100, y + 550)
  pic.Line (x - 200, y + 750)-(x + 200, y + 750)
  pic.Line (x - 100, y + 950)-(x + 100, y + 950)
  pic.Line (x, y + 950)-(x, y + 1300)
  
  pic.Line (x + 200, y + 200)-(x + 400, y + 200)  'Plus (+) sign
  pic.Line (x + 300, y + 100)-(x + 300, y + 300)
  
  pic.Line (x + 200, y + 1100)-(x + 400, y + 1100)  'Negative (-) sign
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


Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim s, t As Single
  s = (x \ 1300) * 1300
  t = (y \ 1300) * 1300
  
  'DrawResistorVert s, t
  'DrawResistorHor s, t
  'DrawLineVert s, t - 1300
  'DrawLineHor s - 1300, t
 ' DrawBattery s, t
  
End Sub
