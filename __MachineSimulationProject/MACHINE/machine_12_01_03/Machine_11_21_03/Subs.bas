Attribute VB_Name = "Subs"
'********************************************************
' Global Sub Procedures
' Written by Chuck Bolin, November 2003
'********************************************************
Option Explicit

'starting point of program is here
Public Sub Main()
  InitializeVariables
  frmMain.Show
End Sub



'initializes all global variables
Public Sub InitializeVariables()
  
  'cosmetic variables
  gstrProgramName = "Machine Simulation"
  gstrProgramDate = "November 21, 2003"
  gstrProgramVersion = "v0.1"
  
  'array variables
  gintTotalObjects = 0
  gintTotalCylinders = 0
  gintTotalTrays = 0
End Sub
