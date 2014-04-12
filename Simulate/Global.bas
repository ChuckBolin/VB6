Attribute VB_Name = "Global"
Option Explicit

Public Enum SIMULATE_OBJECTS
  soActive = 1
  soPassive = 2
  soRect = 3
  soCircle = 4
  soDone = 4
  soInMotion = 5
  soExtend = 6
  soRetract = 7
  soWait = 8
  soActuator = 9
  soTray = 10
End Enum

Public Type OBJECTS
  Shape As SIMULATE_OBJECTS
  Type As SIMULATE_OBJECTS
  CenterX As Single 'center of object
  CenterY As Single
  X1 As Single  'relative values to centerx,centery
  X2 As Single
  Y1 As Single
  Y2 As Single
  Change As String
  DValue As Single
  Value As Single
  Min As Single
  Max As Single
  State As SIMULATE_OBJECTS
  Command As SIMULATE_OBJECTS
End Type

Public Const g_nMaxObjects = 1

Public o(g_nMaxObjects) As OBJECTS

'loads all object data
Public Sub LoadObjects()
  o(0).Shape = soRect
  o(0).Type = soActuator
  o(0).CenterX = 10
  o(0).CenterY = 30
  o(0).X1 = -2
  o(0).X2 = 2
  o(0).Y1 = 2
  o(0).Y2 = -2
  o(0).Change = "X2"
  o(0).DValue = 0.5
  o(0).Min = 0
  o(0).Max = 30
  o(0).Value = 0
  o(0).State = soDone
  
  o(1).Shape = soRect
  o(1).Type = soTray
  o(1).CenterX = 16
  o(1).CenterY = 30
  o(1).X1 = -4
  o(1).X2 = 4
  o(1).Y1 = 4
  o(1).Y2 = -4
  o(1).Change = ""
  o(1).DValue = 0
  o(1).Min = 0
  o(1).Max = 0
  o(1).Value = 0
  o(1).State = soWait
End Sub

