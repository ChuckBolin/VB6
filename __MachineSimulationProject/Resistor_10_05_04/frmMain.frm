VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resistor Problem Solver v0.2 by C. Bolin, October 2004"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11895
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
      Left            =   8910
      TabIndex        =   4
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton cmdCombo1 
      Caption         =   "Create Combination One"
      Height          =   375
      Left            =   8910
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdPar 
      Caption         =   "Create Parallel"
      Height          =   375
      Left            =   8910
      TabIndex        =   2
      Top             =   420
      Width           =   2295
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Create Series"
      Height          =   375
      Left            =   8910
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
      Height          =   7035
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   0
      Width           =   8115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum RES_MODE
  NONE = 0
  SERIES
  PARALLEL
  COMBO1
  COMBO2
End Enum

Private Enum RESISTOR_FAULT
  RESISTOR_OK = 1
  RESISTOR_OPEN = 1000000
  RESISTOR_SHORT = 0
End Enum

Private Type Resistor
  Resistor As Single
  Value As Single
  Fault As RESISTOR_FAULT
  Voltage As Single
  Current As Single
End Type

'module variables
Private m_nMode As RES_MODE
Private m_nRes(6) As Single
Private m_nVolt(3) As Single
Private m_uR(3) As Resistor
Private m_nRT As Single 'total parameters
Private m_nVS As Single
Private m_nIT As Single

'various drawing functions of components, lines and nodes
Private Sub DrawNode(x As Single, y As Single)
  'pic.FillColor = vbBlack
  pic.Circle (x, y), 50
End Sub

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

Private Sub DrawLineHor(ByVal x As Single, ByVal y As Single)
  pic.Line (x, y)-(x + 1300, y)
End Sub

Private Sub DrawLineVert(ByVal x As Single, ByVal y As Single)
  pic.Line (x, y)-(x, y + 1300)
End Sub

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
  m_uR(0).Fault = RESISTOR_OK
  m_uR(1).Fault = RESISTOR_OK
  m_uR(2).Fault = RESISTOR_OK
  m_uR(3).Fault = RESISTOR_OK
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).Value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value of RT
  m_nRT = 0
  m_nRT = m_uR(0).Value + m_uR(3).Value
  m_nRT = m_nRT + (m_uR(1).Value * m_uR(2).Value) / (m_uR(1).Value + m_uR(2).Value)
    
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  m_uR(0).Current = m_nIT
  m_uR(0).Voltage = m_uR(0).Current * m_uR(0).Value
  m_uR(3).Current = m_nIT
  m_uR(3).Voltage = m_uR(3).Current * m_uR(3).Value
    
  m_uR(1).Voltage = m_nVS - (m_uR(0).Voltage + m_uR(3).Voltage)
  m_uR(1).Current = m_uR(1).Voltage / m_uR(1).Value
  m_uR(2).Voltage = m_uR(1).Voltage
  m_uR(2).Current = m_uR(2).Voltage / m_uR(2).Value
  
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
  m_uR(0).Fault = RESISTOR_OK
  m_uR(1).Fault = RESISTOR_OK
  m_uR(2).Fault = RESISTOR_OK
  m_uR(3).Fault = RESISTOR_OK
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).Value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value of RT
  m_nRT = 0
  m_nRT = m_uR(0).Value
  m_nRT = m_nRT + ((m_uR(1).Value + m_uR(3).Value) * m_uR(2).Value) / (m_uR(1).Value + m_uR(2).Value + m_uR(3).Value)
    
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  
  m_uR(0).Current = m_nIT
  m_uR(0).Voltage = m_uR(0).Current * m_uR(0).Value
  
  m_uR(2).Voltage = m_nVS - m_uR(0).Voltage
  m_uR(2).Current = m_uR(3).Voltage / m_uR(3).Value
  
  m_uR(1).Current = m_uR(0).Current - m_uR(2).Current
  m_uR(3).Current = m_uR(1).Current
  
  m_uR(1).Voltage = m_uR(1).Current * m_uR(1).Value
  m_uR(3).Voltage = m_uR(3).Current * m_uR(3).Value
    
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

Private Sub cmdExit_Click()
  End
End Sub

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
  m_uR(0).Fault = RESISTOR_OK
  m_uR(1).Fault = RESISTOR_OK
  m_uR(2).Fault = RESISTOR_OK
  m_uR(3).Fault = RESISTOR_OK
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).Value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value if RT
  m_nRT = 0
  For i = 0 To 3
    m_nRT = m_nRT + 1 / m_uR(i).Value '* m_uR(i).Fault)
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
    m_uR(i).Current = m_nVS / m_uR(i).Value
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
  m_uR(0).Fault = RESISTOR_OK
  m_uR(1).Fault = RESISTOR_OK
  m_uR(2).Fault = RESISTOR_OK
  m_uR(3).Fault = RESISTOR_OK
    
  'calc actual resistor values
  For i = 0 To 3
    m_uR(i).Value = m_uR(i).Resistor * m_uR(i).Fault
  Next i
     
  'calc value if RT
  m_nRT = 0
  For i = 0 To 3
    m_nRT = m_nRT + m_uR(i).Value ' * m_uR(i).Fault)
  Next i
  
  'calc value of IT and voltage drops
  m_nIT = m_nVS / m_nRT
  For i = 0 To 3
    m_uR(i).Voltage = m_uR(i).Value * m_nIT
    m_uR(i).Current = m_nIT
  Next i
 
  
  
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

End Sub

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

Private Function FormatVoltage(nVal As Single) As String
  Dim sOut As String
    
  If nVal < 1 Then
    sOut = CStr(Format(nVal * 1000, "##.###")) & " mV"
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

Private Function FormatCurrent(nVal As Single) As String
  Dim sOut As String
    
  If nVal < 1 Then
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

Private Sub Form_Load()
  'DrawGrid
  'Dim bRet As Boolean
  
  'bRet = r.SetResistorNetwork("*&(*+*)")  '*&(*&*+*)&*&(*&*+*+*)&*
  'If bRet = False Then End
  'MsgBox r.GetTotalResistance
  
  'laod resistor values
  m_nRes(0) = 100
  m_nRes(1) = 330
  m_nRes(2) = 470
  m_nRes(3) = 690
  m_nRes(4) = 1000
  m_nRes(5) = 2200
  m_nRes(6) = 4700
  
  'load voltage values
  m_nVolt(0) = 5
  m_nVolt(1) = 10
  m_nVolt(2) = 12
  m_nVolt(3) = 24
  
  m_nMode = NONE
  
  Randomize Timer
  
  
  
End Sub

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
