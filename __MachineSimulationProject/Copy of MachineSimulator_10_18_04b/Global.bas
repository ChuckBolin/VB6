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
Public g_nPhaseV As Single '277V
Public g_nHotV As Single '120V
Public g_sVersionNumber As String
Public g_sVersionDate As String
Public g_sVersionAuthor As String

'*********************************************************** Main()
'program starts here
Public Sub Main()
  Randomize Timer
  LoadVariables
  InitializeElectrical
  frmMain.Show
  
  frmCP.Show
  frmMachine.Show
  frmDraw.Show
End Sub

'*********************************************************** LoadVariables()
'loads all variables
Public Sub LoadVariables()
  g_nPhaseV = 277  'phase voltage to ground
  g_nHotV = 120 'hot (120V) voltage to ground
  
  g_sVersionNumber = "v .01"
  g_sVersionDate = "Date: 10/18/2004"
  g_sVersionAuthor = "Chuck Bolin"
  
  
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
