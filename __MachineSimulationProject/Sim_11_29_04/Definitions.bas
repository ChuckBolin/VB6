Attribute VB_Name = "Definitions"
Option Explicit

Public Enum OBJECT_TYPE
  Tray = 1
  Conveyor
  ProximitySwitch
End Enum

Public Type RECTANGLE
  X1 As Single
  Y1 As Single
  X2 As Single
  Y2 As Single
  Width As Single
  Height As Single
End Type

Public Type RECT_OBJECT
  R As RECTANGLE
  Visible As Boolean
  BackColor As Long
  Speed As Single 'speed of conveyor
  Type As OBJECT_TYPE
End Type

Public s(10) As RECT_OBJECT
Public g_nMax As Integer 'maximum number of rectangular objects
Public g_nFocus As Integer 'item with focus of mouse click  '-1 equal no focus

Public Sub LoadObjects()
  g_nMax = 3
  g_nFocus = -1
  
  'conveyor
  s(0).BackColor = vbWhite
  s(0).Visible = True
  s(0).R.X1 = 1000
  s(0).R.Y1 = 1000
  s(0).R.X2 = 5000
  s(0).R.Y2 = 1200
  s(0).R.Height = 200
  s(0).R.Width = 4000
  s(0).Speed = 100
  s(0).Type = Conveyor
 
  'tray 1
  s(1).BackColor = vbBlack
  s(1).Visible = True
  s(1).R.X1 = 2000
  s(1).R.Y1 = 2000
  s(1).R.X2 = 2100
  s(1).R.Y2 = 2100
  s(1).R.Height = 100
  s(1).R.Width = 100
  s(1).Speed = 0
  s(1).Type = Tray
  
  'proximity switch
  s(2).BackColor = vbBlack
  s(2).Visible = True
  s(2).R.X1 = 4000
  s(2).R.Y1 = 700
  s(2).R.X2 = 4100
  s(2).R.Y2 = 900
  s(2).R.Height = 200
  s(2).R.Width = 100
  s(2).Speed = 0
  s(3).Type = ProximitySwitch

End Sub
