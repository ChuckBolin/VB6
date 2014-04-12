Attribute VB_Name = "Electrical"
'********************************************************************************
' MODULE Electrical.Bas - Written by Chuck Bolin, September 27, 2004
' Purpose:
' Indicates status of electrical fault possibilities
'********************************************************************************
Option Explicit
Option Base 1

'public const
Public Const MAX_ELECTRICAL_COMPONENTS = 100
Public Const V_120_X1 = 1
Public Const V_120_X2 = 2
Public Const V_120_X2_GND = 3
Public Const D_120_TRANSFORMER_SECONDARY = 4
Public Const D_120_OCPD = 5
Public Const V_120_L1 = 6
Public Const V_120_N = 7
Public Const D_120_PLC_OCPD = 8
Public Const V_120_PLC_L1 = 9
Public Const V_120_PLC_N = 10
Public Const D_120_PLC_POWER_FUSE = 11
Public Const D_120_PLC_POWER_SWITCH = 12
Public Const D_120_PLC_POWER_SUPPLY = 13
Public Const V_24_PLC_POS = 14
Public Const V_24_PLC_NEG = 15
Public Const D_PLC_INPUT_MODULE = 16
Public Const D_PLC_OUTPUT_MODULE = 17


'public declarations of arrays and variables
Public e(MAX_ELECTRICAL_COMPONENTS) As Boolean

'initialize electrical components
Public Sub InitializeElectrical()

  'initialize all devices
  e(D_120_TRANSFORMER_SECONDARY) = True
  e(D_120_OCPD) = True
  e(D_120_PLC_OCPD) = True
  e(D_120_PLC_POWER_FUSE) = True
  e(D_120_PLC_POWER_SWITCH) = True
  e(D_120_PLC_POWER_SUPPLY) = True
  
  'initializes all supply voltages
  e(V_24_PLC_NEG) = True
  e(V_24_PLC_POS) = True
  e(D_PLC_INPUT_MODULE) = True
  e(D_PLC_OUTPUT_MODULE) = True
End Sub

'refresh the status of all electrical variables...used to indicate faults
Public Sub RefreshElectrical()
  
  'assume transformer secondary is okay with voltage
  'these electrical rules always apply.  Machine switches
  'and computer generated faults affect the state of these
  'arrays. TRUE means the device is good and not damaged.
  'FALSE indicates a problem with this device or terminal
  'D_ is a device.  V_ is a terminal where voltages may be read
  
  e(V_120_X1) = e(D_120_TRANSFORMER_SECONDARY)
  e(V_120_X2) = e(D_120_TRANSFORMER_SECONDARY)
  e(V_120_L1) = e(V_120_X1) And e(D_120_OCPD)
  e(V_120_N) = e(V_120_X2)
  e(V_120_PLC_L1) = e(V_120_L1) And e(D_120_PLC_OCPD)
  e(V_120_PLC_N) = e(V_120_N)
  
  g_uPLC.PowerApplied = e(V_120_PLC_L1) And e(D_120_PLC_POWER_FUSE) And e(D_120_PLC_POWER_SWITCH) And e(D_120_PLC_POWER_SUPPLY)
  
  If e(D_PLC_INPUT_MODULE) And e(V_24_PLC_NEG) Then
    IO.EnableInputs
  Else
    IO.DisableInputs
  End If
  
  If e(V_24_PLC_POS) And e(D_PLC_OUTPUT_MODULE) Then
    IO.EnableOutputs
  Else
    IO.DisableOutputs
  End If
  
  'frmPLC.Caption = e(V_24_PLC_POS)
  
  
End Sub
