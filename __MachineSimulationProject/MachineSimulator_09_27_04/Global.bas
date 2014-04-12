Attribute VB_Name = "Global"
Option Explicit

'public types
'PLC type
Public Type PLC_INFO
  Power120V As Boolean
  Power12VIndicator As Boolean
  Power5VIndicator As Boolean
  Input24V As Boolean
  Output0V As Boolean
  PowerSwitch As Boolean
  PowerFuse As Boolean
  RunSwitch As Boolean
  RunIndicator As Boolean
  ClearOutputSwitch As Boolean
End Type


'public variables
Public g_uPLC As PLC_INFO
Public IO As New CBooleanEvaluator

'program starts here
Public Sub Main()
  LoadVariables
  frmMain.Show
  frmPLC.Show
End Sub

'loads all variables
Public Sub LoadVariables()
  
  'initialize PLC in frmPLC
  g_uPLC.Power120V = True
  g_uPLC.Power12VIndicator = True
  g_uPLC.Power5VIndicator = True
  g_uPLC.PowerSwitch = True
  g_uPLC.PowerFuse = True
  g_uPLC.Input24V = True
  g_uPLC.Output0V = True
  g_uPLC.RunIndicator = False
  g_uPLC.RunSwitch = True
  g_uPLC.ClearOutputSwitch = False
  
End Sub

'manages and processes all troubleshooting and system characteristics
Public Sub ProcessSystem()

  'PLC power switch
  g_uPLC.Power12VIndicator = g_uPLC.PowerSwitch And g_uPLC.Power120V And g_uPLC.PowerFuse
  g_uPLC.Power5VIndicator = g_uPLC.PowerSwitch And g_uPLC.Power120V And g_uPLC.PowerFuse
  g_uPLC.RunIndicator = g_uPLC.RunSwitch
End Sub
