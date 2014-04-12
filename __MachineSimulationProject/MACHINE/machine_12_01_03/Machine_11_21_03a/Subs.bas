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
  gstrProgramVersion = "v0.1a"
  
  'array variables
  gintTotalObjects = 0
  gintTotalCylinders = 0
  gintTotalTrays = 0
End Sub

'adds an object to the machine
Public Sub AddObject(obj As Integer)
  Dim coord As COORDINATE_PAIR
  
  Select Case obj
  
    Case gCYLINDER:
    
      'increase size of array
      gintTotalCylinders = gintTotalCylinders + 1
      gintTotalObjects = gintTotalObjects + 1
      ReDim Preserve gObj(gintTotalObjects)
      ReDim Preserve gCyl(gintTotalCylinders)
      Set gCyl(gintTotalCylinders).obj = frmMach.Cylinder1
      
      'calculate corners of control
      gObj(gintTotalObjects).type = gCYLINDER
      gCyl(gintTotalCylinders).obj.Left = 3000
      gCyl(gintTotalCylinders).obj.Top = 3000
      coord.X = gCyl(gintTotalCylinders).obj.Left
      coord.Y = gCyl(gintTotalCylinders).obj.Top
      gObj(gintTotalObjects).quad.NW = coord
      coord.X = gCyl(gintTotalCylinders).obj.Left + gCyl(gintTotalCylinders).obj.Width
      coord.Y = gCyl(gintTotalCylinders).obj.Top
      gObj(gintTotalObjects).quad.NE = coord
      coord.X = gCyl(gintTotalCylinders).obj.Left + gCyl(gintTotalCylinders).obj.Width
      coord.Y = gCyl(gintTotalCylinders).obj.Top - gCyl(gintTotalCylinders).obj.Height
      gObj(gintTotalObjects).quad.SE = coord
      coord.X = gCyl(gintTotalCylinders).obj.Left
      coord.Y = gCyl(gintTotalCylinders).obj.Top - gCyl(gintTotalCylinders).obj.Height
      gObj(gintTotalObjects).quad.SW = coord
      gCyl(gintTotalCylinders).obj.Visible = True
      
      
      
      
      
    Case gPARTTRAY:
    
  End Select
  
End Sub
