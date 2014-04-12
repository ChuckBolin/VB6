Attribute VB_Name = "Electrical"
'********************************************************************************
' MODULE Electrical.Bas - Written by Chuck Bolin, September 27, 2004
' Purpose:
' Indicates status of electrical fault possibilities
'********************************************************************************
Option Explicit
Option Base 1

'Public constants. Each electrical component and subcomponents have
'assigned constants.
'Placed into array e( ) and represents all components. True if component is
'connected correctly or operated correctly
'Placed into array f( ) and represents a faulty condition. True if faulty. False
'if no fault.
Public Const MAX_ELECTRICAL_COMPONENTS = 92
Public Const Q0 = 1       'Main Disconnect
Public Const Q0_12 = 2
Public Const Q0_34 = 3
Public Const Q0_56 = 4
Public Const F1 = 5       'Main fuses
Public Const F1_A = 6
Public Const F1_B = 7
Public Const F1_C = 8
Public Const F3 = 9       'Motor 1 fuses
Public Const F3_A = 10
Public Const F3_B = 11
Public Const F3_C = 12
Public Const F4 = 13      'Motor 2 fuses
Public Const F4_A = 14
Public Const F4_B = 15
Public Const F4_C = 16
Public Const F2_A = 17    'Transformer primary fuses
Public Const F2_B = 18
Public Const T1 = 19      'Transformer
Public Const T1_PRI = 20  'Transformer primary
Public Const T1_H1H2 = 21
Public Const T1_H2H3 = 22
Public Const T1_H3H4 = 23
Public Const T1_SEC = 24  'Transformer secondary
Public Const F5 = 25      'Fuses
Public Const F6 = 26
Public Const F7 = 27
Public Const F8 = 28
Public Const F9 = 29
Public Const F10 = 30
Public Const PLC_PS = 31  'PLC stuff
Public Const PLC_PS_FUSE = 32
Public Const PLC_PS_PWR_SW = 33
Public Const PLC_STOP_SW = 34
Public Const PLC_CLEAR_SW = 35
Public Const PS_24V = 36   '24V Power Supply
Public Const S0_AIR = 37   'Various switches in control circuit
Public Const S1_ESTOP = 38
Public Const S2_CONT_OFF = 39
Public Const S3_CONT_ON = 40
Public Const CR = 41      'Control Relay
Public Const CR_COIL = 42
Public Const CR_12 = 43
Public Const CR_34 = 44
Public Const RECEPT1 = 45  'Receptacles and lighting for machine
Public Const RECEPT2 = 46
Public Const LIGHTING = 47
Public Const S4_CONVEYOR_ON = 48  'Inputs
Public Const S5_CONVEYOR_OFF = 49
Public Const S6_AUTO = 50
Public Const S6_SEMI = 51
Public Const S7_AUTOSTART = 52
Public Const S8_AUTOSTOP = 53
Public Const S9_JOG = 54
Public Const PLC_INPUT_MOD1 = 55  'PLC Modules and voltage requirements
Public Const PLC_INPUT_MOD2 = 56
Public Const PLC_0V_MOD1 = 57
Public Const PLC_0V_MOD2 = 58
Public Const PLC_OUTPUT_MOD1 = 59
Public Const PLC_24V_MOD1 = 60
Public Const M1 = 61      'Motor relay 1
Public Const M1_COIL = 62
Public Const M1_12 = 63
Public Const M1_34 = 64
Public Const M2 = 65      'Motor relay 2
Public Const M2_COIL = 66
Public Const M2_12 = 67
Public Const M2_34 = 68
Public Const H1_AUTO = 69
Public Const H2_FAULT = 70
Public Const S10_PROX = 71
Public Const S11_Z1_RETRACT = 72
Public Const S12_Z1_EXTEND = 73
Public Const S13_PROX = 74
Public Const S14_Z3_RETRACT = 75
Public Const S15_Z3_EXTEND = 76
Public Const S16_PROX = 77
Public Const S17_PROX = 78
Public Const S18_Z5_CLAMP = 79
Public Const S19_PROX = 80
Public Const S20_Z7_RETRACT = 81
Public Const S21_Z7_EXTEND = 82
Public Const Y1_CYL = 83
Public Const Y2_SEP = 84
Public Const Y3_CYL = 85
Public Const Y4_SEP = 86
Public Const Y5_CLAMP = 87
Public Const Y6_SEP = 88
Public Const Y7_CYL = 89
Public Const H0_CONT_ON = 90
Public Const MOT1 = 91
Public Const MOT2 = 92

'Voltage nodes stored in array v( )
Public Const MAX_VOLTAGE_NODES = 80
Public Const V_L1 = 1
Public Const V_L2 = 2
Public Const V_L3 = 3
Public Const V_0L1 = 4
Public Const V_0L2 = 5
Public Const V_0L3 = 6
Public Const V_1L1 = 7
Public Const V_1L2 = 8
Public Const V_1L3 = 9
Public Const V_3L1 = 10
Public Const V_3L2 = 11
Public Const V_3L3 = 12
Public Const V_4L1 = 13
Public Const V_4L2 = 14
Public Const V_4L3 = 15
Public Const V_T1_H1 = 16
Public Const V_T1_H2 = 17
Public Const V_T1_H3 = 18
Public Const V_T1_H4 = 19
Public Const V_T1_X1 = 20
Public Const V_T1_X2 = 21
Public Const V_120V_L1 = 22
Public Const V_120V_N = 23
Public Const V_120V_L1_RECEPT = 24
Public Const V_120V_L1_PLC = 25
Public Const V_120V_L1_LIGHTING = 26
Public Const V_120V_L1_PS = 27
Public Const V_24V = 28
Public Const V_BE_24V = 29
Public Const V_AE_24V = 30
Public Const V_S0_AIR = 31
Public Const V_S1_ESTOP = 32
Public Const V_S2_CONT_OFF = 33
Public Const V_S3_CONT_ON = 34
Public Const V_S4_CONV_ON = 35
Public Const V_S5_CONV_OFF = 36
Public Const V_S6_AUTO = 37
Public Const V_S6_SEMI = 38
Public Const V_S7_AUTOSTART = 39
Public Const V_S8_AUTOSTOP = 40
Public Const V_S9_JOG = 41
Public Const V_AE_INPUT = 42
Public Const V_INPUT_MOD1_0V = 43
Public Const V_INPUT_MOD2_0V = 44
Public Const V_M1_CONV1 = 45
Public Const V_M2_CONV2 = 46
Public Const V_H1_AUTO = 47
Public Const V_H2_FAULT = 48
Public Const V_Y1_CYL = 49
Public Const V_Y2_SEP = 50
Public Const V_Y3_CYL = 51
Public Const V_Y4_SEP = 52
Public Const V_Y5_CLAMP = 53
Public Const V_Y6_SEP = 54
Public Const V_Y7_CYL = 55
Public Const V_S10_PROX = 56
Public Const V_S11_Z1_RETRACT = 57
Public Const V_S12_Z1_EXTEND = 58
Public Const V_S13_PROX = 59
Public Const V_S14_Z3_RETRACT = 60
Public Const V_S15_Z3_EXTEND = 61
Public Const V_S16_PROX = 62
Public Const V_S17_PROX = 63
Public Const V_S18_Z5_CLAMP = 64
Public Const V_S19_PROX = 65
Public Const V_S20_Z7_RETRACT = 66
Public Const V_S21_Z7_EXTEND = 67
Public Const V_M1_12 = 68
Public Const V_M1_34 = 69
Public Const V_M1_56 = 70
Public Const V_M2_12 = 71
Public Const V_M2_34 = 72
Public Const V_M2_56 = 73
Public Const V_Q1_12 = 74
Public Const V_Q1_34 = 75
Public Const V_Q1_56 = 76
Public Const V_Q2_12 = 77
Public Const V_Q2_34 = 78
Public Const V_Q2_56 = 79
Public Const V_OUTPUT_MOD1_24V = 80


'public declarations of arrays and variables
Public e(MAX_ELECTRICAL_COMPONENTS) As Boolean
Public f(MAX_ELECTRICAL_COMPONENTS) As Boolean
Public v(MAX_VOLTAGE_NODES) As Boolean

'******************************************************** InitializeElectrical()
'initialize electrical components
Public Sub InitializeElectrical()
  Dim i As Integer
  
  'all components are in correct position for circuit operation
  'and there are no faults inserted into the circuits
  For i = 1 To MAX_ELECTRICAL_COMPONENTS
    e(i) = True
    f(i) = False
  Next i
  
  'initialize all devices
  v(V_L1) = True  'main power coming into machine..all starts here
  v(V_L2) = True
  v(V_L3) = True
   
  'base position
  e(PLC_STOP_SW) = False
  e(PLC_CLEAR_SW) = False
  e(S4_CONVEYOR_ON) = False
  e(S5_CONVEYOR_OFF) = False
  e(S6_AUTO) = True
  e(S6_SEMI) = False
  e(S7_AUTOSTART) = False
  e(S8_AUTOSTOP) = False
  e(S9_JOG) = False
  e(S0_AIR) = True
  e(S1_ESTOP) = True
  e(S2_CONT_OFF) = True
  e(S3_CONT_ON) = True
End Sub

'******************************************************* RefreshElectrical
'refresh the status of all electrical variables...used to indicate faults
Public Sub RefreshElectrical()
  
  'Update voltage nodes
  'these are read as...voltage at 0L1 is present if voltage is present
  'at L1 and component Q0, contact 1 and 2 are closed and not fault
  'exists with component Q0, contacts 1 and 2
  v(V_0L1) = v(V_L1) And e(Q0_12) And Not f(Q0_12) 'Main disconnect
  v(V_0L2) = v(V_L2) And e(Q0_34) And Not f(Q0_34)
  v(V_0L3) = v(V_L3) And e(Q0_56) And Not f(Q0_56)
  v(V_1L1) = v(V_0L1) And e(F1_A) And Not f(F1_A)  '480V distribution
  v(V_1L2) = v(V_0L2) And e(F1_B) And Not f(F1_B)
  v(V_1L3) = v(V_0L3) And e(F1_C) And Not f(F1_C)
  v(V_T1_H1) = v(V_1L3) And e(F2_A) And Not f(F2_A) 'primary of transformer
  v(V_T1_H4) = v(V_1L2) And e(F2_B) And Not f(F2_B)
  v(V_T1_X1) = v(V_T1_H1) And v(V_T1_H4) And e(T1_PRI) And Not f(T1_PRI) 'trans sec voltage
  v(V_120V_L1) = v(V_T1_X1) And e(F3) And Not f(F3)
  v(V_120V_L1_PS) = v(V_120V_L1) And e(F9) And Not f(F9)
  v(V_120V_L1_PLC) = v(V_120V_L1) And e(F7) And Not f(F7)
  v(V_24V) = v(V_120V_L1_PS) And e(PS_24V) And Not f(PS_24V)
  v(V_BE_24V) = v(V_24V) And e(F10) And Not f(F10)
  v(V_AE_24V) = v(V_BE_24V) And e(CR) And Not f(CR)
  v(V_OUTPUT_MOD1_24V) = v(V_AE_24V)
  v(V_INPUT_MOD1_0V) = True
  v(V_INPUT_MOD2_0V) = True
  
  'enable loads
  e(MOT1) = v(V_Q1_12) And v(V_Q1_34) And v(V_Q1_56) And Not f(MOT1)
   
  
  'PLC inputs
  v(V_S4_CONV_ON) = v(V_BE_24V) And e(S4_CONVEYOR_ON) And Not f(S4_CONVEYOR_ON) And v(V_INPUT_MOD1_0V)
  v(V_S5_CONV_OFF) = v(V_BE_24V) And e(S5_CONVEYOR_OFF) And Not f(S5_CONVEYOR_OFF) And v(V_INPUT_MOD1_0V)
  v(V_S6_AUTO) = v(V_BE_24V) And e(S6_AUTO) And Not f(S6_AUTO) And v(V_INPUT_MOD1_0V)
  v(V_S6_SEMI) = v(V_BE_24V) And e(S6_SEMI) And Not f(S6_SEMI) And v(V_INPUT_MOD1_0V)
  v(V_S7_AUTOSTART) = v(V_BE_24V) And e(S7_AUTOSTART) And Not f(S7_AUTOSTART) And v(V_INPUT_MOD1_0V)
  v(V_S8_AUTOSTOP) = v(V_BE_24V) And e(S8_AUTOSTOP) And Not f(S8_AUTOSTOP) And v(V_INPUT_MOD1_0V)
  v(V_S9_JOG) = v(V_BE_24V) And e(S9_JOG) And Not f(S9_JOG) And v(V_INPUT_MOD1_0V)

  'set inputs in response to signals from sensors
  IO.SetInput 1, v(V_S4_CONV_ON)
  IO.SetInput 2, v(V_S5_CONV_OFF)
  IO.SetInput 3, v(V_S6_AUTO)
  IO.SetInput 4, v(V_S6_SEMI)
  IO.SetInput 5, v(V_S7_AUTOSTART)
  IO.SetInput 6, v(V_S8_AUTOSTOP)
  IO.SetInput 7, v(V_S9_JOG)
    
  
  'operate machine subsystems in response to required or missing voltages
  g_uPLC.PowerApplied = v(V_120V_L1_PLC) And e(PLC_PS) And Not f(PLC_PS) And e(PLC_PS_FUSE) And e(PLC_PS_PWR_SW)
    
  If e(PLC_INPUT_MOD1) And Not f(PLC_INPUT_MOD2) And v(V_INPUT_MOD1_0V) Then
    IO.EnableInputs
  Else
    IO.DisableInputs
  End If
  
  If e(PLC_OUTPUT_MOD1) And Not f(PLC_OUTPUT_MOD1) And e(V_OUTPUT_MOD1_24V) Then
    IO.EnableOutputs
  Else
    IO.DisableOutputs
  End If
  
  frmPLC.Caption = v(V_S6_AUTO)
  
  
End Sub
