Attribute VB_Name = "Global"
Option Explicit

'public types
'PLC type
Public Type PLC_INFO
  PowerApplied As Boolean 'plc is electrical normal..doesn't means switches in correct position
                          'used for fault insertion
  PowerSwitch As Boolean  'plc power on/off switch is on if true
  RunSwitch As Boolean    'run switch is in run mode if in up position
  ClearOutputSwitch As Boolean 'outputs are cleared if in up position
  RunIndicator As Boolean
  Power12VIndicator As Boolean
  Power5VIndicator As Boolean

End Type


'public variables
Public g_uPLC As PLC_INFO
Public IO As New CBooleanEvaluator

'program starts here
Public Sub Main()
  LoadVariables
  InitializeElectrical
  frmMain.Show
  frmPLC.Show
End Sub

'loads all variables
Public Sub LoadVariables()
  
  'initialize PLC in frmPLC
  g_uPLC.PowerSwitch = True
  g_uPLC.RunSwitch = True
  g_uPLC.ClearOutputSwitch = False
  'g_uPLC.Power12VIndicator = True
  'g_uPLC.Power5VIndicator = True
  'g_uPLC.RunIndicator = False
  
End Sub

'manages and processes all troubleshooting and system characteristics
Public Sub ProcessSystem()

  'PLC power switch
  If g_uPLC.PowerSwitch And g_uPLC.RunSwitch And g_uPLC.PowerApplied Then
    IO.StartProcess
    frmPLC.Caption = "Running"
  Else
    IO.StopProcess
    frmPLC.Caption = "Stopped"
  End If
  
  If g_uPLC.ClearOutputSwitch And g_uPLC.PowerApplied Then
    IO.DisableOutputs
  Else
    IO.EnableOutputs
  End If
    
  g_uPLC.Power12VIndicator = g_uPLC.PowerSwitch And g_uPLC.PowerApplied
  g_uPLC.Power5VIndicator = g_uPLC.PowerSwitch And g_uPLC.PowerApplied
  g_uPLC.RunIndicator = g_uPLC.PowerApplied And g_uPLC.RunSwitch
  


End Sub
