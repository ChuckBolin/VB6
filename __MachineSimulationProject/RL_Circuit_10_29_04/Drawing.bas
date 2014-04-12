Attribute VB_Name = "Drawing"
Option Explicit

Public Const SPECIAL_OMEGA = 1    'resistance symbol
Public Const SPECIAL_RADICAL = 2  'short radical sign
Public Const SPECIAL_RADICAL2 = 3 'long radical sign
Public Const SPECIAL_SQUARE = 4   'squareroot
'***************************  D R A W I N G   P R O C E D U R E S ********
'**************************************************** DrawNode
'various drawing functions of components, lines and nodes
Public Sub DrawNode(x As Single, y As Single)
  frmMain.picDraw.FillColor = vbBlack
  frmMain.picDraw.Circle (x, y), 50
End Sub

'**************************************************** DrawGrid
'draws grid
Public Sub DrawGrid()
  Dim i, j As Single
  frmMain.picDraw.DrawWidth = 2
  For i = 0 To 10
    For j = 0 To 10
      frmMain.picDraw.PSet (i * 1300, j * 1300)
    Next j
  Next i
  frmMain.picDraw.DrawWidth = 1
End Sub

'**************************************************** DrawLineHor
Public Sub DrawLineHor(ByVal x As Single, ByVal y As Single)
  frmMain.picDraw.Line (x, y)-(x + 1300, y)
End Sub

'**************************************************** DrawLineVert
Public Sub DrawLineVert(ByVal x As Single, ByVal y As Single)
  frmMain.picDraw.Line (x, y)-(x, y + 1300)
End Sub

'**************************************************** DrawResistorHor
Public Sub DrawResistorHor(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  frmMain.picDraw.Line (x, y)-(x + 350, y)
  frmMain.picDraw.Line -(x + 400, y - 200)
  frmMain.picDraw.Line -(x + 500, y + 200)
  frmMain.picDraw.Line -(x + 600, y - 200)
  frmMain.picDraw.Line -(x + 700, y + 200)
  frmMain.picDraw.Line -(x + 800, y - 200)
  frmMain.picDraw.Line -(x + 900, y + 200)
  frmMain.picDraw.Line -(x + 950, y)
  frmMain.picDraw.Line -(x + 1300, y)
  'DrawNode x + 1300, y
End Sub

'**************************************************** DrawResistorVert
Public Sub DrawResistorVert(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  frmMain.picDraw.Line (x, y)-(x, y + 350)
  frmMain.picDraw.Line -(x + 200, y + 400)
  frmMain.picDraw.Line -(x - 200, y + 500)
  frmMain.picDraw.Line -(x + 200, y + 600)
  frmMain.picDraw.Line -(x - 200, y + 700)
  frmMain.picDraw.Line -(x + 200, y + 800)
  frmMain.picDraw.Line -(x - 200, y + 900)
  frmMain.picDraw.Line -(x, y + 950)
  frmMain.picDraw.Line -(x, y + 1300)
  'DrawNode x, y + 1300
End Sub


'**************************************************** DrawBattery
'Public Sub DrawBattery(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
'  picDraw.Line (x, y)-(x, y + 350)
'  picDraw.Line (x - 200, y + 350)-(x + 200, y + 350)
'  picDraw.Line (x - 100, y + 550)-(x + 100, y + 550)
'  picDraw.Line (x - 200, y + 750)-(x + 200, y + 750)
'  picDraw.Line (x - 100, y + 950)-(x + 100, y + 950)
'  picDraw.Line (x, y + 950)-(x, y + 1300)
'
'  picDraw.Line (x + 200, y + 200)-(x + 400, y + 200)  'Plus (+) sign
'  picDraw.Line (x + 300, y + 100)-(x + 300, y + 300)
'
'  picDraw.Line (x + 200, y + 1100)-(x + 400, y + 1100)  'Negative (-) sign
  'DrawNode x, y + 1300
  
'End Sub

'**************************************************** DrawInductorVert
Public Sub DrawInductorVert(ByVal x As Single, ByVal y As Single)
  frmMain.picDraw.Line (x, y)-(x, y + 350)
  frmMain.picDraw.Circle (x, y + 425), 75, , 4.57, 1.57
  frmMain.picDraw.Circle (x, y + 575), 75, , 4.57, 1.57
  frmMain.picDraw.Circle (x, y + 725), 75, , 4.57, 1.57
  frmMain.picDraw.Circle (x, y + 875), 75, , 4.57, 1.57
  frmMain.picDraw.Line (x, y + 950)-(x, y + 1300)
End Sub

'**************************************************** DrawInductorHor
Public Sub DrawInductorHor(ByVal x As Single, ByVal y As Single)
  frmMain.picDraw.Line (x, y)-(x + 350, y)
  frmMain.picDraw.Circle (x + 425, y), 75, , 0, 3.14
  frmMain.picDraw.Circle (x + 575, y), 75, , 0, 3.14
  frmMain.picDraw.Circle (x + 725, y), 75, , 0, 3.14
  frmMain.picDraw.Circle (x + 875, y), 75, , 0, 3.14
  frmMain.picDraw.Line (x + 950, y)-(x + 1300, y)
End Sub

'**************************************************** DrawACSource
Public Sub DrawACSource(ByVal x As Single, ByVal y As Single)
  'DrawNode x, y
  frmMain.picDraw.Line (x, y)-(x, y + 350)
  frmMain.picDraw.FillColor = vbWhite
  frmMain.picDraw.Circle (x, y + 650), 300
  frmMain.picDraw.Line (x, y + 950)-(x, y + 1300)
  frmMain.picDraw.CurrentX = x - 100
  frmMain.picDraw.CurrentY = y + 550
  frmMain.picDraw.Print "AC"
  'DrawNode x, y + 1300
  
End Sub

Public Sub DrawInductorCircuit()
  frmMain.picDraw.Cls
  frmMain.picDraw.DrawStyle = 0
  
  DrawLineVert 500, 200
  DrawACSource 500, 1500
  DrawLineVert 500, 2800
  DrawLineHor 500, 200
  DrawLineHor 1800, 200
  DrawLineHor 3100, 200
  DrawLineHor 500, 4100
  DrawLineHor 1800, 4100
  DrawLineHor 3100, 4100
  DrawLineVert 4400, 200
  DrawInductorVert 4400, 1500
  DrawText 4600, 2000, "L1 = " & FormatNumber(g_uInductor(0).Inductance) & "H"
  DrawText 1000, 2000, FormatNumber(g_uSource.Voltage) & "V"
  DrawText 1000, 2300, FormatNumber(g_uSource.Frequency) & "Hz"
  If g_eMode = CM_InductorQ Then
    DrawResistorVert 4400, 2800
    DrawText 4700, 3300, FormatNumber(g_uInductor(0).Resistance) & "Ohms"
    frmMain.picDraw.DrawStyle = 2
    frmMain.picDraw.FillStyle = 1
    frmMain.picDraw.Line (4000, 1200)-(4800, 3900), , B
  ElseIf g_eMode = CM_InductorOnly Then
    DrawLineVert 4400, 2800
  End If
End Sub

Public Sub DrawSeriesCircuit()
  frmMain.picDraw.Cls
  frmMain.picDraw.DrawStyle = 0
  
  DrawLineVert 500, 200
  DrawACSource 500, 1500
  DrawText 1000, 2000, FormatNumber(g_uSource.Voltage) & "V"
  DrawText 1000, 2300, FormatNumber(g_uSource.Frequency) & "Hz"
  DrawLineVert 500, 2800
  
  DrawLineHor 500, 200
  DrawInductorHor 1800, 200
  DrawText 2200, 500, "L1 = " & FormatNumber(g_uInductor(0).Inductance) & "H"
  
  DrawLineVert 4400, 200
  DrawInductorVert 4400, 1500
  DrawText 4600, 1700, "L2 = " & FormatNumber(g_uInductor(1).Inductance) & "H"
  
  DrawInductorHor 1800, 4100
  DrawText 2000, 3550, "L3 = " & FormatNumber(g_uInductor(2).Inductance) & "H"
  DrawLineHor 3100, 4100
  
  If g_eMode = CM_SeriesQ Then
    DrawResistorHor 3100, 200
    DrawText 3400, 500, FormatNumber(g_uInductor(0).Resistance) & "Ohms"
    DrawResistorVert 4400, 2800
    DrawText 4700, 3300, FormatNumber(g_uInductor(1).Resistance) & "Ohms"
    DrawResistorHor 500, 4100
    DrawText 950, 3550, FormatNumber(g_uInductor(2).Resistance) & "Ohms"
    
    frmMain.picDraw.DrawStyle = 2
    frmMain.picDraw.FillStyle = 1
    frmMain.picDraw.Line (1400, 0)-(4700, 600), , B
    frmMain.picDraw.Line (4000, 1200)-(4800, 3900), , B
    frmMain.picDraw.Line (200, 3800)-(3400, 4300), , B
  ElseIf g_eMode = CM_SeriesOnly Then
    DrawLineHor 3100, 200
    DrawLineVert 4400, 2800
    DrawLineHor 500, 4100
  End If
End Sub

'draws parallel inductor circuit with three inductors
Public Sub DrawParallelCircuit()
  frmMain.picDraw.Cls
  frmMain.picDraw.DrawStyle = 0
  
  DrawLineVert 500, 200
  DrawACSource 500, 1500
  DrawText 1000, 2000, FormatNumber(g_uSource.Voltage) & "V"
  DrawText 1000, 2300, FormatNumber(g_uSource.Frequency) & "Hz"
  DrawLineVert 500, 2800
  DrawLineHor 500, 200
  DrawLineHor 1800, 200
  DrawLineHor 3100, 200
  DrawLineHor 500, 4100
  DrawLineHor 1800, 4100
  DrawLineHor 3100, 4100
  DrawLineVert 1800, 200
  DrawLineVert 3100, 200
  DrawLineVert 4400, 200
  DrawLineVert 1800, 2800
  DrawLineVert 3100, 2800
  DrawLineVert 4400, 2800
  DrawInductorVert 1800, 1500
  DrawText 2000, 2000, "L1 = " & FormatNumber(g_uInductor(0).Inductance) & "H"
  DrawInductorVert 3100, 1500
  DrawText 3300, 2000, "L2 = " & FormatNumber(g_uInductor(1).Inductance) & "H"
  DrawInductorVert 4400, 1500
  DrawText 4600, 2000, "L3 = " & FormatNumber(g_uInductor(2).Inductance) & "H"
  
End Sub

'************************************************************** DrawSeriesRLCircuit
'draws series RL circuit
Public Sub DrawSeriesRLCircuit()

  frmMain.picDraw.Cls
  frmMain.picDraw.DrawStyle = 0
  
  DrawLineVert 500, 200
  DrawACSource 500, 1500
  DrawText 1000, 2000, FormatNumber(g_uSource.Voltage) & "V"
  DrawText 1000, 2300, FormatNumber(g_uSource.Frequency) & "Hz"
  DrawLineVert 500, 2800
  DrawLineHor 500, 200
  DrawLineHor 3100, 200
  DrawLineHor 500, 4100
  DrawLineHor 1800, 4100
  DrawLineHor 3100, 4100
  DrawLineVert 4400, 200
  DrawLineVert 4400, 2800
  DrawInductorVert 4400, 1500
  DrawText 4600, 2000, "L1 = " & FormatNumber(g_uInductor(0).Inductance) & "H"
  DrawResistorHor 1800, 200
  DrawText 2000, 500, FormatNumber(g_uResistor.Resistance) & "Ohms"
  
  
End Sub

'************************************************************** DrawParallelRLCircuit
'draws series RL circuit
Public Sub DrawParallelRLCircuit()

  frmMain.picDraw.Cls
  frmMain.picDraw.DrawStyle = 0
  
  DrawLineVert 500, 200
  DrawACSource 500, 1500
  DrawText 1000, 2000, FormatNumber(g_uSource.Voltage) & "V"
  DrawText 1000, 2300, FormatNumber(g_uSource.Frequency) & "Hz"
  DrawLineVert 500, 2800
  DrawLineHor 500, 200
  DrawLineHor 1800, 200
  DrawLineHor 3100, 200
  DrawLineHor 500, 4100
  DrawLineHor 1800, 4100
  DrawLineHor 3100, 4100
  DrawLineVert 4400, 200
  DrawLineVert 4400, 2800
  DrawLineVert 3100, 200
  DrawLineVert 3100, 2800
  DrawInductorVert 4400, 1500
  DrawText 4600, 2000, "L1 = " & FormatNumber(g_uInductor(0).Inductance) & "H"
  DrawResistorVert 3100, 1500
  DrawText 3400, 2000, FormatNumber(g_uResistor.Resistance) & "Ohms"
  
  
End Sub
'****************************************************** DrawText
Public Sub DrawText(ByVal x As Single, ByVal y As Single, s As String)
  frmMain.picDraw.CurrentX = x
  frmMain.picDraw.CurrentY = y
  frmMain.picDraw.Print s
End Sub

'****************************************************** ShowText
Public Sub ShowText(ByVal x As Single, ByVal y As Single, s As String)
  frmMain.picShow.CurrentX = x
  frmMain.picShow.CurrentY = y
  frmMain.picShow.Print s
End Sub

'****************************************************** ShowSpecial
'Prints radical sign, squared (2 superscript) and omega
Public Sub ShowSpecial(nChar As Integer, ByVal x As Single, ByVal y As Single)
  Dim s, t As Single 'for drawing references
  
  Select Case nChar
    Case SPECIAL_OMEGA
      s = x + 100: t = y + 100
      frmMain.picShow.Circle (s, t), 75, , 5.17, 4.07
      frmMain.picShow.Line (s - 95, t + 75)-(s - 35, t + 75)
      frmMain.picShow.Line (s + 35, t + 75)-(s + 95, t + 75)
    Case SPECIAL_RADICAL
      frmMain.picShow.Line (x, y)-(x + 40, y + 180) 'radical
      frmMain.picShow.Line -(x + 80, y)
      frmMain.picShow.Line -(x + 1600, y)
    Case SPECIAL_RADICAL2
      frmMain.picShow.Line (x, y)-(x + 40, y + 180) 'radical long
      frmMain.picShow.Line -(x + 80, y)
      frmMain.picShow.Line -(x + 2400, y)
    Case SPECIAL_SQUARE
      frmMain.picShow.CurrentX = x
      frmMain.picShow.CurrentY = y
      frmMain.picShow.FontSize = 8
      frmMain.picShow.Print "2"
      frmMain.picShow.FontSize = 12
  End Select
End Sub

'***************************************************** FormatNumber
'appends number with m, micro,n,K,M, etc
Public Function FormatNumber(nVal As Single)
  Dim sOut As String
  Dim sFactor As String
  
  If nVal < 0.000001 Then
    sOut = CStr(Format(nVal * 1000000000, "###.###"))
    sFactor = " n"
  ElseIf nVal < 0.001 Then
    sOut = CStr(Format(nVal * 1000000, "###.###"))
    sFactor = " µ"
  ElseIf nVal < 1 Then
    sOut = CStr(Format(nVal * 1000, "###.###"))
    sFactor = " m"
  ElseIf nVal < 1000 Then
    sOut = CStr(Format(nVal * 1, "###.###"))
    sFactor = " "
  ElseIf nVal < 1000000 Then
    sOut = CStr(Format(nVal * 0.001, "###.###"))
    sFactor = " K"
  ElseIf nVal < 1000000000 Then
    sOut = CStr(Format(nVal * 0.000001, "###.###"))
    sFactor = " M"
  End If
 
  If Right(sOut, 1) = "." Then
    sOut = Left(sOut, Len(sOut) - 1)
  Else
    sOut = sOut
  End If
  
  FormatNumber = sOut & sFactor
End Function


'***************************************************** FormatNumberNoUnits
'No appending.
Public Function FormatNumberNoUnits(nVal As Single)
  Dim sOut As String
  
  If nVal < 0.000001 Then
    sOut = CStr(Format(nVal * 1000000000, "###.###"))
  ElseIf nVal < 0.001 Then
    sOut = CStr(Format(nVal * 1000000, "###.###"))
  ElseIf nVal < 1 Then
    sOut = CStr(Format(nVal * 1000, "###.###"))
  ElseIf nVal < 1000 Then
    sOut = CStr(Format(nVal * 1, "###.###"))
  ElseIf nVal < 1000000 Then
    sOut = CStr(Format(nVal * 0.001, "###.###"))
  ElseIf nVal < 1000000000 Then
    sOut = CStr(Format(nVal * 0.000001, "###.###"))
  End If
 
  If Right(sOut, 1) = "." Then
    sOut = Left(sOut, Len(sOut) - 1)
  Else
    sOut = sOut
  End If
  
  FormatNumberNoUnits = sOut
End Function

