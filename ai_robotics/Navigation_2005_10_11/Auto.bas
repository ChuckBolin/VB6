Attribute VB_Name = "Auto"
Option Explicit

'*****************************************************************
' This is the Autonomous routine. Its purpose is to modify two
' variables:
' Inputs:   bot.X
'           bot.Y
'           bot.Direction
'           dr.X
'           dr.Y
'           dr.Direction
'
'           leg(1).X1
'           leg(1).Y1
'           leg(1).X2
'           leg(1).Y2
'           leg(1).Orientation
'           leg(1).Width
'
' Outputs: bot.Turn
'          bot.velocity
'
' Try using these functions with bot position and waypoint info (leg())
'  GetTargetDistance2D ()
'  GetTargetDirection2D()
'*****************************************************************

'October 11, 2005
Public Sub Autonomous()
  Dim nWPDir As Single
  Dim nWPDist As Single
  Dim nDirDiff As Single 'difference between nWPDir and bot.dir
  
  'get to direction and distance to next waypoint
  nWPDir = GetTargetDirection2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  nWPDist = GetTargetDistance2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  
  'calc angular difference
  nDirDiff = nWPDir - bot.Direction
  
  If Abs(nDirDiff) >= PI Then
    If nDirDiff > 0 Then
      nWPDir = nWPDir - PI
    Else
      nWPDir = nWPDir + PI
    End If
    nDirDiff = nWPDir - (bot.Direction - PI)
    'MsgBox "Here"
  End If
  'frmMain.Caption = nDirDiff
  'determine turning amount
  If nDirDiff > 0 Then
    bot.Turn = bot.Turn + 0.002
  ElseIf nDirDiff < 0 Then
    bot.Turn = bot.Turn - 0.002
  Else
    bot.Turn = 0
  End If
  
  
  'get distance to next waypoint
  If nWPDist > 1000 Then
    bot.Velocity = bot.MaxVel / 3
  ElseIf nWPDist > 500 Then
    bot.Velocity = bot.MaxVel / 3
  ElseIf nWPDist > 200 Then
    bot.Velocity = bot.MaxVel / 3
  Else  'arrived at waypoint
    g_nLegNum = g_nLegNum + 1
    If g_nLegNum > g_nLastLegNum Then g_nLegNum = 1
  End If
End Sub

'October 10, 2005
Public Sub Autonomous2()
  'Dim nLaneDir As Single 'depends upon orientation..angle in radians
  'Dim nDirDiff As Single 'difference between dr.direction and nLaneDir
  'Dim nDist As Single 'distance to next waypoint position
  'Dim nFuzzyPos As Integer 'stores fuzzy position within lane
  'Dim nFuzzyDir As Integer 'stores fuzzy direction within lane
  
  bot.Turn = bot.Turn + 0.001
  bot.Velocity = 10
  
  
  'determine lane direction
  'If leg(g_nLegNum).Orientation = 1 Then 'north
  '  nLaneDir = PI / 2
  'ElseIf leg(g_nLegNum).Orientation = 2 Then 'east
  '  nLaneDir = 0
  'ElseIf leg(g_nLegNum).Orientation = 3 Then 'south
  '  nLaneDir = 3 * PI / 2
  'ElseIf leg(g_nLegNum).Orientation = 4 Then 'west
  '  nLaneDir = PI
  'Else
  'End If

  'calc direction difference (angular in radians)
  'nDirDiff = dr.Direction - nLaneDir
  'nDirDiff = bot.Direction - nLaneDir
  'If nDirDiff > 0.75 Then
  '  nFuzzyDir = LANE_DIR_LEFT_FAR
  'ElseIf nDirDiff > 0.3 Then
  '  nFuzzyDir = LANE_DIR_LEFT
  'ElseIf nDirDiff < -0.75 Then
  '  nFuzzyDir = LANE_DIR_RIGHT_FAR
  'ElseIf nDirDiff > -0.3 Then
  '  nFuzzyDir = LANE_DIR_RIGHT
  'Else
  '  nFuzzyDir = LANE_DIR_CENTER
  'End If
    
  'nDist = GetTargetDistance2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  'nDist = GetTargetDistance2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  'frmMain.Caption = leg(g_nLegNum).X2 & " " & leg(g_nLegNum).Y2 & " " & leg(g_nLegNum).Orientation
  'If nDist < 1000 Then
  '  g_nLegNum = g_nLegNum + 1
  '  If g_nLegNum > g_nLastLegNum Then g_nLegNum = 1
  'End If
  'bot.Velocity = 20
 '
  'determine direction
  'If nFuzzyDir = LANE_DIR_LEFT_FAR Then
  '  bot.Turn = bot.Turn - 0.05
  'ElseIf nFuzzyDir = LANE_DIR_LEFT Then
  '  bot.Turn = bot.Turn - 0.001
  'ElseIf nFuzzyDir = LANE_DIR_CENTER Then
    'bot.Turn = bot.Turn
  'ElseIf nFuzzyDir = LANE_DIR_RIGHT Then
  '  bot.Turn = bot.Turn + 0.001
  'ElseIf nFuzzyDir = LANE_DIR_RIGHT_FAR Then
  '  bot.Turn = bot.Turn + 0.05
  'End If
  
  'frmMain.Caption = nFuzzyDir
  
  
  
End Sub


