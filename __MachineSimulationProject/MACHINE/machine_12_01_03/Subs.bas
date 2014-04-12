Attribute VB_Name = "Subs"
'************************************************************
' Global Sub Procedures
' Written by Chuck Bolin, November 2003
'************************************************************
Option Explicit

'************************************************************
' M A I N     M A I N     M A I N     M A I N
'starting point of program is here
'************************************************************
Public Sub Main()
  frmMain.e.FileName = App.Path & "\logfile.txt"
  frmMain.e.StartProgram
  InitializeVariables
  frmMain.Show
End Sub

'************************************************************
' I N I T I A L I Z E   V A R I A B L E S
'initializes all global variables
'************************************************************
Public Sub InitializeVariables()
  
  'cosmetic variables
  gstrProgramName = "Machine Simulation"
  gstrProgramDate = "November 21, 2003"
  gstrProgramVersion = "v0.1b"
  
  frmMain.e.LogData gstrProgramName & " " & gstrProgramVersion & " " & gstrProgramDate
  
    
  'array variables
  gintTotalObjects = 0
  gintTotalCylinders = 0
  gintTotalTrays = 0
End Sub

'*************************************************************
' A D D  O B J E C T
'adds an object to the machine
'*************************************************************
Public Sub AddObject(obj As Integer)
  Dim coord As COORDINATE_PAIR
  Dim intNextIndex As Integer
    
  Select Case obj
    Case gCYLINDER:
      'loads another control onto form
      intNextIndex = frmMach.cyl.UBound + 1
      Load frmMach.cyl(intNextIndex)
      frmMach.cyl(intNextIndex).Orientation = 0
      frmMach.cyl(intNextIndex).CylinderLength = 800 '1500
      frmMach.cyl(intNextIndex).CylinderWidth = 200 '400
      frmMach.cyl(intNextIndex).SetSize
      gintTotalCylinders = gintTotalCylinders + 1
      frmMach.cyl(intNextIndex).designation = "Y" & CStr(gintTotalCylinders)
      frmMach.cyl(intNextIndex).Visible = True
      frmMach.cyl(intNextIndex).SetFocus
      
      'increase size of arrays
      gintTotalObjects = gintTotalObjects + 1
      ReDim Preserve gObj(gintTotalObjects)
      ReDim Preserve gCyl(gintTotalCylinders)
      
    Case gPARTTRAY:
      intNextIndex = frmMach.tray.UBound + 1
      Load frmMach.tray(intNextIndex)
      frmMach.tray(intNextIndex).Visible = True
      
    Case gSHAPE:
      intNextIndex = frmMach.Shape.UBound + 1
      Load frmMach.Shape(intNextIndex)
      frmMach.Shape(intNextIndex).Visible = True
      frmMach.Shape(intNextIndex).RunTime = True
      frmMach.Shape(intNextIndex).ZOrder 1
  End Select
End Sub
