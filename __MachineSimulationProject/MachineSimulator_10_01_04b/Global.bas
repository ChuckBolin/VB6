Attribute VB_Name = "Global"
Option Explicit

'public types
'PLC type
Public Type PLC_INFO
  PowerApplied As Boolean 'plc is electrical normal..doesn't means switches in correct position
                          'used for fault insertion
  StopIndicator As Boolean
  Power12VIndicator As Boolean
  Power5VIndicator As Boolean
End Type

'public variables
Public g_uPLC As PLC_INFO
Public IO As New CBooleanEvaluator
Public g_nPhaseV As Single

'*********************************************************** Main()
'program starts here
Public Sub Main()
  LoadVariables
  InitializeElectrical
  frmMain.Show
  frmPLC.Show
  frmCP.Show
  frmMachine.Show
  frmFault.Show
  frmCab.Show
End Sub

'*********************************************************** LoadVariables()
'loads all variables
Public Sub LoadVariables()
  g_nPhaseV = 277  'phase voltage to ground
  
  
End Sub

'************************************************************ ProcessSystem ()
'manages and processes all troubleshooting and system characteristics
Public Sub ProcessSystem()

  'process PLC hardware
  If e(PLC_PS_PWR_SW) And Not e(PLC_STOP_SW) And g_uPLC.PowerApplied Then
    IO.StartProcess
  Else
    IO.StopProcess
  End If
  
  'plc indicators
  g_uPLC.Power12VIndicator = e(PLC_PS_PWR_SW) And g_uPLC.PowerApplied
  g_uPLC.Power5VIndicator = e(PLC_PS_PWR_SW) And g_uPLC.PowerApplied
  g_uPLC.StopIndicator = g_uPLC.PowerApplied And e(PLC_STOP_SW)
  
  'control outputs
  If Not e(PLC_STOP_SW) And Not e(PLC_CLEAR_SW) And g_uPLC.PowerApplied Then
    IO.EnableOutputs
  Else
    IO.DisableOutputs
  End If
  
  If v(V_INPUT_MOD1_0V) And v(V_INPUT_MOD2_0V) Then
    IO.EnableInputs
  Else
    IO.DisableInputs
  End If
  
  'frmPLC.Caption = Not e(PLC_CLEAR_SW) ' And v(V_INPUT_MOD2_0V) 'IO.GetInput(2)
End Sub
