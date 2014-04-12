VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' CBOTS.CLS - Written by Chuck Bolin, August 28, 2004
' Updated: 12.08.04 CB
' Create a new instance like...
' Public P As New CBots
'**********************************************************
Option Explicit

'constants
Private Const MAX_BOTS = 6
Private Const PI = 3.14159
Private Const HOME = 1
Private Const VISITOR = 2

'type of bot info
Private Type BOT_INFO
  x As Single
  y As Single
  TargetX As Single
  TargetY As Single
  TargetBearing As Single 'direction to TargetX,Y
  Velocity As Single
  MaxVelocity As Single
  Direction As Single     'direction bot is heading
  Diameter As Single
  CX As Single            'collision avoidance coordinate
  CY As Single
  CAvoid As Boolean       'true if following CX,CY
  Team As Integer         'home or visitor
  Color As Long           'color of object when visible
  IntermediateX() As Single 'for complex movements
  IntermediateY() As Single
  IntermediateIndex As Integer 'current intermediate point
End Type

'create 100 bots
Private B(MAX_BOTS) As BOT_INFO
Private nWithinRangeOfTarget As Single
Private nMax_X As Single
Private nMin_X As Single
Private nMax_Y As Single
Private nMin_Y As Single

'*************************************** GetIntermediateBearing
Private Function GetIntermediateBearing(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
    
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  Dim ix As Single
  Dim iy As Single
  Dim index As Integer
  Dim inum As Integer
  
  inum = GetNumIntermediatePos(id)  'total number of inter points
  If inum < 1 Then Exit Sub         'none exists...exit
  index = B(i).IntermediateIndex    'current position in array of inter pts
  If index >= inum Then Exit Sub    'already gone through all points
  
  ix = B(i).IntermediateX           'get current ix,iy
  iy = B(i).IntermediateY
  
  dy = iy - B(id).y
  dx = ix - B(id).x
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetIntermediateBearing = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetIntermediateBearing = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetIntermediateBearing = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetIntermediateBearing = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetIntermediateBearing = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetIntermediateBearing = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetIntermediateBearing = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetIntermediateBearing = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetIntermediateBearing = PI + PI / 2
  Else
    '?
  End If
  
  'keep values between 0 and 2*PI
  If GetIntermediateBearing > 2 * PI Then GetIntermediateBearing = GetIntermediateBearing - 2 * PI
  If GetIntermediateBearing < 0 Then GetIntermediateBearing = GetIntermediateBearing + 2 * PI
End Function

'************************************** GetMaxBots
'returns the max bots in class
Public Function GetMaxBots() As Integer
  GetMaxBots = MAX_BOTS
End Function

'************************************** X, Y coordinates
'Sets X position
Public Function SetX(id As Integer, x As Single) As Boolean
  SetX = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).x = x
  SetX = True
End Function

'Returns X position
Public Function GetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetX = B(id).x
End Function

'Sets Y position
Public Function SetY(id As Integer, y As Single) As Boolean
  SetY = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).y = y
  SetY = True
End Function

'Returns y position
Public Function GetY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetY = B(id).y
End Function

'************************************** Velocity
Public Function GetVelocity(id As Integer)
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetVelocity = B(id).Velocity
End Function

Public Function SetVelocity(id As Integer, Velocity As Single) As Boolean
  SetVelocity = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).Velocity = Velocity
  SetVelocity = True
End Function

'************************************** MaxVelocity
Public Function SetMaxVelocity(id As Integer, vel As Single) As Boolean
  SetMaxVelocity = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).MaxVelocity = vel
  SetMaxVelocity = True
End Function

'************************************** GetMaxVelocity
Public Function GetMaxVelocity(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetMaxVelocity = B(id).MaxVelocity
End Function

'************************************** Direction
Public Function GetDirection(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetDirection = B(id).Direction
End Function

'************************************** SetDiameter
Public Function SetDiameter(id As Integer, Diameter As Single) As Boolean
  SetDiameter = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).Diameter = Diameter
  SetDiameter = True
End Function

'************************************** GetDiameter
Public Function GetDiameter(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetDiameter = B(id).Diameter
End Function

'************************************** TargetX, TargetY coordinates
'Sets Target X position
Public Function SetTargetX(id As Integer, x As Single) As Boolean
  SetTargetX = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).TargetX = x
  SetTargetX = True
End Function

'Returns Target X position
Public Function GetTargetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetX = B(id).TargetX
End Function

'Sets Target Y position
Public Function SetTargetY(id As Integer, y As Single) As Boolean
  SetTargetY = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  B(id).TargetY = y
  SetTargetY = True
End Function

'Returns Target y position
Public Function GetTargetY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetY = B(id).TargetY
End Function

'*************************************** GetNumIntermediatePos
Public Function GetNumIntermediatePos(id As Integer) As Integer
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetNumIntermediatePos = UBound(B(id).IntermediateX)
End Function

'*************************************** AddIntermediatePos
Public Function AddIntermediatePos(id As Integer, x As Single, y As Single) As Boolean
  AddIntermediatePos = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  ReDim Preserve B(id).IntermediateX(UBound(B(id).IntermediateX) + 1)
  ReDim Preserve B(id).IntermediateY(UBound(B(id).IntermediateY) + 1)
  B(id).IntermediateX(UBound(IntermediateX)) = x
  B(id).IntermediateY(UBound(IntermediateY)) = y
  AddIntermediatePos = True
End Function

'*************************************** ClearIntermediatePos
Public Function ClearIntermediatePos(id As Integer) As Boolean
  ClearIntermediatePos = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  ReDim B(id).IntermediateX(0)
  ReDim B(id).IntermediateY(0)
  ClearIntermediatePos = True
End Function

'****************************************** SetTeam
Public Function SetTeam(id As Integer, Team As Integer) As Boolean
  SetTeam = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  If Team < 1 Or Team > 2 Then Exit Function
  B(id).Team = Team
  SetTeam = True
End Function

'****************************************** GetTeam
Public Function GetTeam(id As Integer) As Integer
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTeam = B(id).Team
End Function

'****************************************** SetColor
Public Function SetColor(id As Integer, Color As Long) As Boolean
  SetColor = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  If Color < 0 Or Color > 2 ^ 24 Then Exit Function
  B(id).Color = Color
  SetColor = True
End Function

'****************************************** GetColor
Public Function GetColor(id As Integer) As Long
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetColor = B(id).Color
End Function

'****************************************** AtTarget
Public Function AtTarget(id As Integer) As Boolean
  AtTarget = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  If Sqr((B(id).TargetX - B(id).x) ^ 2 + (B(id).TargetY - B(id).y) ^ 2) < B(id).Diameter Then
    AtTarget = True
  End If
End Function

'****************************************** GetAvoidStatus
'returns true if in avoidance mode
Public Function GetAvoidStatus(id As Integer) As Boolean
  GetAvoidStatus = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetAvoidStatus = B(id).CAvoid
End Function

'****************************************** GetCX
Public Function GetCX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetCX = B(id).CX
End Function

'****************************************** GetCY
Public Function GetCY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetCY = B(id).CY
End Function

'**************************************** GetBotsInRegion - Private
'returns bot ID number within a rectangular region
'(x1,y1) are top-left of region
'(x2,y2) are bottom-right of region
Private Function GetBotsInRegion(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Variant()
  Dim vBots() As Variant
  Dim nTemp As Single
  Dim i As Integer
  
  'check parameters, make sure x2 is more positive than x1
  If x1 > x2 Then
    nTemp = x2
    x2 = x1
    x1 = nTemp
  End If
  If y2 > y1 Then
    nTemp = y1
    y1 = y2
    y2 = nTemp
  End If
  
  'evaluate all bots to see if their (x,y) position is within
  'region. Add to variant array if so
  ReDim vBots(0)
  For i = 1 To MAX_BOTS
    DoEvents
    If B(i).x > x1 And B(i).x < x2 And B(i).y < y1 And B(i).y > y2 Then
      ReDim Preserve vBots(UBound(vBots) + 1)
      vBots(UBound(vBots)) = i
    End If
  Next i
  
  'returns variant array of bot IDs in the region
  GetBotsInRegion = vBots
End Function

'************************************** GetTargetBearing
Private Function GetTargetBearing(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
    
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  
  dy = B(id).TargetY - B(id).y
  dx = B(id).TargetX - B(id).x
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetTargetBearing = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetTargetBearing = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetTargetBearing = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetTargetBearing = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetTargetBearing = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetTargetBearing = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetTargetBearing = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetTargetBearing = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetTargetBearing = PI + PI / 2
  Else
    '?
  End If
  
  'keep values between 0 and 2*PI
  If GetTargetBearing > 2 * PI Then GetTargetBearing = GetTargetBearing - 2 * PI
  If GetTargetBearing < 0 Then GetTargetBearing = GetTargetBearing + 2 * PI
End Function

'****************************************** GetTargetDistance
Public Function GetTargetDistance(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetDistance = Sqr((B(id).TargetX - B(id).x) ^ 2 + (B(id).TargetY - B(id).y) ^ 2)
End Function

'************************************* ObstacleID
'returns ID number of closest bot between id.x,y and one of two choices
'1) id.targetx,id.targety or 2) id.ix,id.iy
Private Function ObstacleID(id As Integer, what As Integer) As Integer
  If id < 1 Or id > MAX_BOTS Then Exit Function
  If what < 1 Or what > 2 Then Exit Function

  Dim nDir As Single 'direction of bot with id
  Dim nTargetDistance As Single 'distance from x,y to targetx,targety
  Dim nBrg As Single 'direction/bearing to other bots
  Dim nDist As Single 'distance to other bots
  Dim nBeta As Single 'angular difference between nBrg and nDir
  Dim nCPA As Single  'bots closes point of approach to bot(id) line
                      'of sight (LOS)
  Dim i As Integer    'counting
  Dim vBots() As Integer  'stores ids of all bots in the way
  Dim nShortDist As Single 'stores value of shortest distance
  Dim nShortID As Single   'stores ID of bot with shortest distance
  Dim nTemp As Single
  
  'initialize variables
  ObstacleID = 0
  nShortID = 0
  ReDim vBots(0)
  nDir = B(id).Direction
  
  If what = 1 Then 'get distance to target
    nTargetDistance = Sqr((B(id).TargetX - B(id).x) ^ 2 + (B(id).TargetY - B(id).y) ^ 2)
  Else             'get distance to next intermediate point
    nTargetDistance = Sqr((B(id).IntermediateX(B(id).IntermediateIndex) - B(id).x) ^ 2 + (B(id).IntermediateY(B(id).IntermediateIndex) - B(id).y) ^ 2)
  End If
  
  'consider all bots...add to list 'vbots()' if they are in the way
  For i = 1 To MAX_BOTS
    DoEvents
    If i <> id Then 'don't evaluate self
      nDist = GetBotDistance(id, i)
      If nTargetDistance >= nDist Then  'this bot could be in the way
        nBrg = GetBotDirection(id, i)
        nBeta = nBrg - nDir
        nCPA = nDist * Sin(nBeta)
        If Abs(nCPA) < B(id).Diameter * 2 Then 'collision possible
          ReDim Preserve vBots(UBound(vBots) + 1) 'increase array size
          vBots(UBound(vBots)) = i  'add this bot ID to list
        End If
      End If
    End If
  Next i
  
  'sort through all bots that are in the way and pick the bot
  'that is closest to the bot being evaluated
  If UBound(vBots) > 0 Then
    nShortID = vBots(1)
    nShortDist = GetBotDistance(id, vBots(1))
    If UBound(vBots) > 1 Then
      For i = 2 To UBound(vBots)
        DoEvents
        nTemp = GetBotDistance(id, vBots(i))
        If nTemp < nShortDist Then
          nShortDist = nTemp
          nShortID = vBots(i)
        End If
      Next i
    End If
  End If
  
  ObstacleID = nShortID
  
End Function

'***************************************** GetAvoidanceDirection
Private Function GetAvoidanceDirection(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  Dim bRet As Boolean
  Dim nBrg As Single 'bearing to tx,ty from x,y
  Dim nCW As Single 'direction to CW
  Dim nCCW As Single 'direction to CCW
  Dim nDiffCW As Single
  Dim nDiffCCW As Single
    
  'bearing to target
  nBrg = GetTargetBearing(id1)
  
  'calc direction to CW position
  nCW = GetCW(id1, id2)
  
  'calc direction to CCW position
  nCCW = GetCCW(id1, id2)
  
  'find smallest angle, going CW or going CCW
  nDiffCW = Abs(nBrg - nCW)
  nDiffCCW = Abs(nBrg - nCCW)
  
  'choose closest avoidance direction
  If nDiffCW <= nDiffCCW Then
    GetAvoidanceDirection = nCW
  Else
    GetAvoidanceDirection = nCCW
  End If
End Function

'************************************** GetATAN
'returns angle based upon opp and adj sides using ATN()
Private Function GetATan(dy As Single, dx As Single) As Single
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetATan = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetATan = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetATan = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetATan = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetATan = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetATan = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetATan = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetATan = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetATan = PI + PI / 2
  Else
    '?
  End If
  
  'keep values between 0 and 2*PI
  If GetATan > 2 * PI Then GetATan = GetATan - 2 * PI
  If GetATan < 0 Then GetATan = GetATan + 2 * PI
End Function

'************************************** GetBotDistance
Private Function GetBotDistance(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  GetBotDistance = Sqr((B(id1).x - B(id2).x) ^ 2 + (B(id1).y - B(id2).y) ^ 2)
End Function

'************************************** GetBotDirection
Private Function GetBotDirection(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  
  dy = B(id2).y - B(id1).y
  dx = B(id2).x - B(id1).x
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetBotDirection = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetBotDirection = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetBotDirection = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetBotDirection = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetBotDirection = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetBotDirection = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetBotDirection = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetBotDirection = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetBotDirection = PI + PI / 2
  Else
    '?
  End If
  
  'keep values between 0 and 2*PI
  If GetBotDirection > 2 * PI Then GetBotDirection = GetBotDirection - 2 * PI
  If GetBotDirection < 0 Then GetBotDirection = GetBotDirection + 2 * PI
End Function

'****************************************** GetCCW
'calculates point perpendicular to line between bot and bot in the way
Private Function GetCCW(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  B(id1).CX = B(id2).x - (B(id1).Diameter * 2) * Sin(nBeta)
  B(id1).CY = B(id2).y + (B(id1).Diameter * 2) * Cos(nBeta)
  nBeta = GetATan(B(id1).CY - B(id1).y, B(id1).CX - B(id1).x)
  GetCCW = nBeta
End Function

'****************************************** GetCW
Private Function GetCW(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  B(id1).CX = B(id2).x + (B(id1).Diameter * 2) * Sin(nBeta)
  B(id1).CY = B(id2).y - (B(id1).Diameter * 2) * Cos(nBeta)
  
  nBeta = GetATan(B(id1).CY - B(id1).y, B(id1).CX - B(id1).x)
  GetCW = nBeta
End Function

'****************************************** UpdateBots
'*****************************************************
'*****************************************************
'Performs all required updates for bot movements
'based upon current position, required behavior, direction
'and velocity.
Public Sub UpdateBots()
  Dim i As Integer
  Dim nObID As Integer
  Dim nDist As Single
  Dim nObDist As Single
  Dim nObDir As Single
  Dim nDir As Single
  Dim ix As Single  'processing of intermediate coordinates
  Dim iy As Single
  Dim index As Integer
  Dim inum As Integer
  Dim ip As Boolean 'true if intermediate point must be processed next
  
  For i = 1 To MAX_BOTS
    DoEvents
    
    'could be intermediate points instead of just tx,ty
    'inum = GetNumIntermediatePos(i)
    'If inum > 0 Then
    '  index = b(i).IntermediateIndex
    '  If index < inum Then
    '    ip = True
    '    ix = b(i).IntermediateX(index)           'get current ix,iy
    '    iy = b(i).IntermediateY(index)
    '  Else
    '    ip = False
    '  End If
    'End If
    B(i).CAvoid = False
    
    If B(i).Velocity <> 0 Then
      nObID = ObstacleID(i, 1)  'find closest obstacle between bot i and his targetx,targety
      'frmMain.Caption = nObID
      
      If nObID < 1 Then 'if no one is in the way, go to target
        B(i).Direction = GetTargetBearing(i)
        'frmMain.Caption = "Direct"
      Else  'someone is in the way
'>>>>>>>> must add code to get obstacles between x,y and ix,iy
        
        'distance between bot in the way and this bot's target
        nObDist = Sqr((B(i).TargetX - B(nObID).x) ^ 2 + (B(i).TargetY - B(nObID).y) ^ 2)
        If GetTargetDistance(i) <= nObDist Then  'obstacle is not between bot and tx,ty
          B(i).Direction = GetTargetBearing(i)  'go straight
          'frmMain.Caption = "Avoid (Dir): " & nObID
        Else
          'GetTargetDistance(i) > nObDist Then
          'nDir = GetTargetBearing(i)
          'nObDir = GetBotDirection(i, nObID)
          'If nObDir > nDir - 1.3 And nObDir < nDir + 1.3 Then
          '  nDist = GetBotDistance(i, nObID)
          '  If nDist < b(i).Diameter Then  'too close'<<<
          '    b(i).CX = b(i).TargetX
          '    b(i).CY = b(i).TargetY
          '    b(i).Direction = GetTargetBearing(i)
              'b(i).Direction = GetAvoidanceDirection(i, nObID)
              'b(i).CAvoid = False
          '  Else
              B(i).Direction = GetAvoidanceDirection(i, nObID)
              B(i).CAvoid = True
             ' frmMain.Caption = "Avoid (CX): " & nObID
          '  End If
          'Else
          '  b(i).Direction = nDir
          'End If
        End If
        
      End If
      B(i).x = B(i).x + B(i).Velocity * Cos(B(i).Direction)
      B(i).y = B(i).y + B(i).Velocity * Sin(B(i).Direction)
      'If b(i).X < nMin_X Then b(i).X = nMin_X + 5
      'If b(i).X > nMax_X Then b(i).X = nMax_X - 5
      'If b(i).Y > nMax_Y Then b(i).Y = nMax_Y - 5
      'If b(i).Y < nMin_Y Then b(i).Y = nMin_Y + 5
    End If
  Next i
End Sub

'****************************************** Initialize Class
Private Sub Class_Initialize()
  Dim i As Integer
  
  For i = 1 To MAX_BOTS
    DoEvents
    B(i).MaxVelocity = 3
    B(i).Velocity = 0.25
    B(i).Diameter = 1
    ReDim B(i).IntermediateX(0)
    ReDim B(i).IntermediateY(0)
    B(i).IntermediateIndex = 0
  Next i
  
  'field restraints
  nMin_X = 5
  nMax_X = 95
  nMin_Y = 5
  nMax_Y = 95
End Sub
