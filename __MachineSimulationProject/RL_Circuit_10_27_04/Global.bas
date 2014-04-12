Attribute VB_Name = "Global"
Option Explicit

Public Enum CIRCUIT_MODE  'Q means factor in internal resistance of inductor
  CM_InductorOnly = 1
  CM_InductorQ
  CM_SeriesOnly
  CM_SeriesQ
  CM_ParallelOnly
  CM_ParallelQ
  CM_LROnly
  CM_LRQ
End Enum

Public Type INDUCTOR
  Inductance As Single   'value of inductance in henries
  XL As Single           'inductive reactance
  Resistance As Single   'value of internal resistance in ohms
  Z As Single            'coil impedance
  Q As Single            'ratio of inductance/resistance
  Voltage As Single      'voltage drop
  Current As Single      'current flow through
  VAR As Single          'volt-amp reactance
  Power As Single        'power of any internal resistance
  VA As Single           'volt-amps of the coil
End Type

Public Type AC_SOURCE
  Voltage As Single      'voltage of source
  Frequency As Single    'frequency of source in Hz
  Current As Single      'total current
  Impedance As Single    'total impedance
End Type

Public Type RESISTOR
  Resistance As Single   'resistance of resistor
  Voltage As Single      'voltage drop across resistor
  Current As Single      'current through resistor
End Type

Public Const PI = 3.14159

Public g_eMode As CIRCUIT_MODE
Public g_uInductor(3) As INDUCTOR    'stores 3 inductors - randomly generated
Public g_uSource As AC_SOURCE        'stores data for AC power source
Public g_uResistor As RESISTOR       'stores data for resistor in LR circuit
Public g_nFreq(5) As Single          'stores various frequency values
Public g_nVolt(5) As Single          'stores various voltage values
Public g_nInductor(5) As Single      'stores various inductor values
Public g_nIntResistance(5) As Single 'stores various internal resistance values
Public g_nResistor(5) As Single      'stores various resistor values
Public g_nTotalSteps As Integer      'total steps in solution
Public g_nCurrentStep As Integer     'current solution step displayed


Public Sub LoadVariables()
  g_eMode = CM_InductorOnly  'default mode
  
  'loads various standard parameters
  g_nFreq(0) = 10
  g_nFreq(1) = 30
  g_nFreq(2) = 60
  g_nFreq(3) = 100
  g_nFreq(4) = 1000
  g_nFreq(5) = 5000
  g_nVolt(0) = 5
  g_nVolt(1) = 10
  g_nVolt(2) = 25
  g_nVolt(3) = 100
  g_nVolt(4) = 120
  g_nVolt(5) = 240
  g_nInductor(0) = 0.001
  g_nInductor(1) = 0.005
  g_nInductor(2) = 0.01
  g_nInductor(3) = 0.05
  g_nInductor(4) = 0.1
  g_nInductor(5) = 0.5
  g_nIntResistance(0) = 0.1
  g_nIntResistance(1) = 0.5
  g_nIntResistance(2) = 1
  g_nIntResistance(3) = 5
  g_nIntResistance(4) = 10
  g_nIntResistance(5) = 25
  g_nResistor(0) = 100
  g_nResistor(1) = 330
  g_nResistor(2) = 470
  g_nResistor(3) = 690
  g_nResistor(4) = 1000
  g_nResistor(5) = 2200
  
End Sub

