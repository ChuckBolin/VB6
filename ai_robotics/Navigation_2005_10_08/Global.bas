Attribute VB_Name = "Global"
Option Explicit

'moving items
Public Type MOBILE_OBJECT
  X As Single
  Y As Single
  VX As Single
  VY As Single
  Velocity As Single
  Direction As Single
  Turn As Single 'amount of turning to affect direction
  MaxVel As Single
  MinVel As Single
  Energy As Single
End Type

'beacons - fixed navigation resources
Public Type NAV_BEACON
  X As Single
  Y As Single
  ID As Integer
  Offset As Single 'error source
End Type

'stores leg information
Public Type AUTO_LEG
  X1 As Single
  Y1 As Single
  X2 As Single
  Y2 As Single
  Width As Single '1/2 of lane width...mpy by 2
  Orientation As Integer '1=N,2=E,3=S,4=W
End Type

Public Type RECT_COORD
  X As Single
  Y As Single
End Type

Public bot As MOBILE_OBJECT
Public nav(3) As NAV_BEACON
Public dr As MOBILE_OBJECT 'this is dead reckoning info..not real..best guess
Public leg(5) As AUTO_LEG
Public g_nLegNum As Integer 'number of leg
Public g_nLastLegNum As Integer 'last leg number..1 is always first leg

'constants used to make data realistic...error prone
Public Const NAV_DR_VEL_FACTOR = 0.9
Public Const NAV_DR_DIR_FACTOR = 1
Public Const NAV_TRIANGULATION_FACTOR = 1

Public Const LANE_POS_LEFT_FAR = 1  'fuzzy logic positions within lane
Public Const LANE_POS_LEFT = 2
Public Const LANE_POS_CENTER = 3
Public Const LANE_POS_RIGHT = 4
Public Const LANE_POS_RIGHT_FAR = 5

Public Const LANE_DIR_LEFT_FAR = 1  'fuzzy logic directions within lane
Public Const LANE_DIR_LEFT = 2
Public Const LANE_DIR_CENTER = 3
Public Const LANE_DIR_RIGHT = 4
Public Const LANE_DIR_RIGHT_FAR = 5


Public Sub LoadVariables()
 
 'this is the bot
 bot.X = 10000
 bot.Y = 10000
 bot.Direction = 1.57
 bot.Velocity = 0
 bot.Turn = 0
 bot.MaxVel = 30
 bot.MinVel = -15
 bot.Energy = 100000
 dr.X = bot.X
 dr.Y = bot.Y
 g_nLegNum = 1
 g_nLastLegNum = 5
  
 'beacon data
 nav(1).ID = 1: nav(1).X = 14000: nav(1).Y = 16000
 nav(2).ID = 2: nav(2).X = 21000: nav(2).Y = 10000
 nav(3).ID = 3: nav(3).X = 7000: nav(3).Y = 5000
 
 'route data
 leg(1).X1 = 10000: leg(1).Y1 = 10000: leg(1).X2 = 10000: leg(1).Y2 = 13000: leg(1).Width = 300: leg(1).Orientation = 1
 leg(2).X1 = 10000: leg(2).Y1 = 13000: leg(2).X2 = 17000: leg(2).Y2 = 13000: leg(2).Width = 300: leg(2).Orientation = 2
 leg(3).X1 = 17000: leg(3).Y1 = 13000: leg(3).X2 = 17000: leg(3).Y2 = 7000: leg(3).Width = 300: leg(3).Orientation = 3
 leg(4).X1 = 17000: leg(4).Y1 = 7000: leg(4).X2 = 10000: leg(4).Y2 = 7000: leg(4).Width = 300: leg(4).Orientation = 4
 leg(5).X1 = 10000: leg(5).Y1 = 7000: leg(5).X2 = 10000: leg(5).Y2 = 10000: leg(5).Width = 300: leg(5).Orientation = 1

End Sub

'this returns the best estimation of robot position based upon
'available triangulation data
Public Function GetTriangulationPosition() As RECT_COORD
  Dim i As Integer
  
End Function

'*****************************************************************
' This is the Autonomous routine. Its purpose is to modify two
' variables:  bot.Turn and bot.velocity
'*****************************************************************
Public Sub Autonomous()
  Dim nLaneDir As Single 'depends upon orientation..angle in radians
  Dim nDirDiff As Single 'difference between dr.direction and nLaneDir
  Dim nDist As Single 'distance to next waypoint position
  Dim nFuzzyPos As Integer 'stores fuzzy position within lane
  Dim nFuzzyDir As Integer 'stores fuzzy direction within lane
  
  'determine lane direction
  If leg(g_nLegNum).Orientation = 1 Then
    nLaneDir = PI / 2
  ElseIf leg(g_nLegNum).Orientation = 2 Then
    nLaneDir = 0
  ElseIf leg(g_nLegNum).Orientation = 3 Then
    nLaneDir = 3 * PI / 2
  ElseIf leg(g_nLegNum).Orientation = 4 Then
    nLaneDir = PI
  Else
  End If

  'calc direction difference (angular in radians)
  nDirDiff = dr.Direction - nLaneDir
  If nDirDiff > 0.75 Then
    nFuzzyDir = LANE_DIR_LEFT_FAR
  ElseIf nDirDiff > 0.3 Then
    nFuzzyDir = LANE_DIR_LEFT
  ElseIf nDirDiff < -0.75 Then
    nFuzzyDir = LANE_DIR_RIGHT_FAR
  ElseIf nDirDiff > -0.3 Then
    nFuzzyDir = LANE_DIR_RIGHT
  Else
    nFuzzyDir = LANE_DIR_CENTER
  End If
    
  nDist = GetTargetDistance2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  If nDist < 1000 Then
    g_nLegNum = g_nLegNum + 1
    If g_nLegNum > g_nLastLegNum Then g_nLegNum = 1
  End If
  bot.Velocity = 20
  
  'determine direction
  If nFuzzyDir = LANE_DIR_LEFT_FAR Then
    bot.Turn = bot.Turn - 0.05
  ElseIf nFuzzyDir = LANE_DIR_LEFT Then
    bot.Turn = bot.Turn - 0.001
  ElseIf nFuzzyDir = LANE_DIR_CENTER Then
    'bot.Turn = bot.Turn
  ElseIf nFuzzyDir = LANE_DIR_RIGHT Then
    bot.Turn = bot.Turn + 0.001
  ElseIf nFuzzyDir = LANE_DIR_RIGHT_FAR Then
    bot.Turn = bot.Turn + 0.05
  End If
  
  frmMain.Caption = nFuzzyDir
  
  
  
End Sub

