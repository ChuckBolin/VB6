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
'**********************************************************
Option Explicit

'constants
Private Const MAX_BOTS = 22
Private Const PI = 3.14159

'type of bot info
Private Type BOT_INFO
  X As Single
  Y As Single
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
End Type

'create 100 bots
Private b(MAX_BOTS) As BOT_INFO
Private nWithinRangeOfTarget As Single
Private nMax_X As Single
Private nMin_X As Single
Private nMax_Y As Single
Private nMin_Y As Single

'************************************** GetMaxBots
'returns the max bots in class
Public Function GetMaxBots() As Integer
  GetMaxBots = MAX_BOTS
End Function

'************************************** X, Y coordinates
'Sets X position
Public Function SetX(id As Integer, X As Single) As Boolean
  SetX = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).X = X
  SetX = True
End Function

'Returns X position
Public Function GetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetX = b(id).X
End Function

'Sets Y position
Public Function SetY(id As Integer, Y As Single) As Boolean
  SetY = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).Y = Y
  SetY = True
End Function

'Returns y position
Public Function GetY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetY = b(id).Y
End Function

'************************************** Velocity
Public Function GetVelocity(id As Integer)
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetVelocity = b(id).Velocity
End Function

Public Function SetVelocity(id As Integer, Velocity As Single) As Boolean
  SetVelocity = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).Velocity = Velocity
  SetVelocity = True
End Function

'************************************** MaxVelocity
Public Function SetMaxVelocity(id As Integer, vel As Single) As Boolean
  SetMaxVelocity = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).MaxVelocity = vel
  SetMaxVelocity = True
End Function

'************************************** GetMaxVelocity
Public Function GetMaxVelocity(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetMaxVelocity = b(id).MaxVelocity
End Function

'************************************** Direction
Public Function GetDirection(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetDirection = b(id).Direction
End Function

'************************************** SetDiameter
Public Function SetDiameter(id As Integer, Diameter As Single) As Boolean
  SetDiameter = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).Diameter = Diameter
  SetDiameter = True
End Function

'************************************** GetDiameter
Public Function GetDiameter(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetDiameter = b(id).Diameter
End Function

'************************************** TargetX, TargetY coordinates
'Sets Target X position
Public Function SetTargetX(id As Integer, X As Single) As Boolean
  SetTargetX = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).TargetX = X
  SetTargetX = True
End Function

'Returns Target X position
Public Function GetTargetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetX = b(id).TargetX
End Function

'Sets Target Y position
Public Function SetTargetY(id As Integer, Y As Single) As Boolean
  SetTargetY = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).TargetY = Y
  SetTargetY = True
End Function

'Returns Target y position
Public Function GetTargetY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetY = b(id).TargetY
End Function

'****************************************** AtTarget
Public Function AtTarget(id As Integer) As Boolean
  AtTarget = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  If Sqr((b(id).TargetX - b(id).X) ^ 2 + (b(id).TargetY - b(id).Y) ^ 2) < b(id).Velocity Then
    AtTarget = True
  End If
End Function

'****************************************** GetAvoidStatus
'returns true if in avoidance mode
Public Function GetAvoidStatus(id As Integer) As Boolean
  GetAvoidStatus = False
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetAvoidStatus = b(id).CAvoid
End Function

'****************************************** GetCX
Public Function GetCX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetCX = b(id).CX
End Function

'****************************************** GetCY
Public Function GetCY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetCY = b(id).CY
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
    If b(i).X > x1 And b(i).X < x2 And b(i).Y < y1 And b(i).Y > y2 Then
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
  
  dy = b(id).TargetY - b(id).Y
  dx = b(id).TargetX - b(id).X
  
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
  GetTargetDistance = Sqr((b(id).TargetX - b(id).X) ^ 2 + (b(id).TargetY - b(id).Y) ^ 2)
End Function

'************************************* ObstacleID
'returns ID number of closest bot between id.x,y and id.targetx,id.targety
Private Function ObstacleID(id As Integer) As Integer
  If id < 1 Or id > MAX_BOTS Then Exit Function

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
  nDir = b(id).Direction
  nTargetDistance = Sqr((b(id).TargetX - b(id).X) ^ 2 + (b(id).TargetY - b(id).Y) ^ 2)
    
  'consider all bots
  For i = 1 To MAX_BOTS
    DoEvents
    If i <> id Then 'don't evaluate self
      nDist = GetBotDistance(id, i)
      If nTargetDistance >= nDist Then  'this bot could be in the way
        nBrg = GetBotDirection(id, i)
        nBeta = nBrg - nDir
        nCPA = nDist * Sin(nBeta)
        If Abs(nCPA) < b(id).Diameter Then  'collision possible
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
  
  If id1 Mod 2 = 0 Then       'even number ID numbers
    bRet = GetCCW(id1, id2)
  Else                        'odd number ID numbers
    bRet = GetCW(id1, id2)
  End If
  
  dy = b(id1).CY - b(id1).Y
  dx = b(id1).CX - b(id1).X
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetAvoidanceDirection = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetAvoidanceDirection = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetAvoidanceDirection = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetAvoidanceDirection = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetAvoidanceDirection = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetAvoidanceDirection = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetAvoidanceDirection = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetAvoidanceDirection = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetAvoidanceDirection = PI + PI / 2
  Else
    '?
  End If
  
  'keep values between 0 and 2*PI
  If GetAvoidanceDirection > 2 * PI Then GetAvoidanceDirection = GetAvoidanceDirection - 2 * PI
  If GetAvoidanceDirection < 0 Then GetAvoidanceDirection = GetAvoidanceDirection + 2 * PI
End Function

'************************************** GetBotDistance
Private Function GetBotDistance(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  GetBotDistance = Sqr((b(id1).X - b(id2).X) ^ 2 + (b(id1).Y - b(id2).Y) ^ 2)
End Function

'************************************** GetBotDirection
Private Function GetBotDirection(id1 As Integer, id2 As Integer) As Single
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  
  dy = b(id2).Y - b(id1).Y
  dx = b(id2).X - b(id1).X
  
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
Private Function GetCCW(id1 As Integer, id2 As Integer) As Boolean
  GetCCW = False
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  b(id1).CX = b(id2).X - (b(id1).Diameter) * Sin(nBeta)
  b(id1).CY = b(id2).Y + (b(id1).Diameter) * Cos(nBeta)
  GetCCW = True
End Function

'****************************************** GetCW
Private Function GetCW(id1 As Integer, id2 As Integer) As Boolean
  GetCW = False
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  b(id1).CX = b(id2).X + (b(id1).Diameter) * Sin(nBeta)
  b(id1).CY = b(id2).Y - (b(id1).Diameter) * Cos(nBeta)
  GetCW = True
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
    
  For i = 1 To MAX_BOTS
    DoEvents
    b(i).CAvoid = False
    If b(i).Velocity <> 0 Then
      nObID = ObstacleID(i)
      If nObID < 1 Then
        b(i).Direction = GetTargetBearing(i)
      Else
        'a behavior exists in which two bots actually revolve about
        'each other forever. This code tries to correct that
        '
        
        'distance between bot in the way and this bot's target
        nObDist = Sqr((b(i).TargetX - b(nObID).X) ^ 2 + (b(i).TargetY - b(nObID).Y) ^ 2)
        If GetTargetDistance(i) <= nObDist Then
          b(i).Direction = GetTargetBearing(i)
        ElseIf GetTargetDistance(i) > nObDist Then
          nDir = GetTargetBearing(i)
          nObDir = GetBotDirection(i, nObID)
          If nObDir > nDir - 0.3 And nObDir < nDir + 0.3 Then
            nDist = GetBotDistance(i, nObID)
            If nDist < b(i).Diameter * 2 Then 'too close
              b(i).CX = b(i).TargetX
              b(i).CY = b(i).TargetY
              b(i).Direction = GetTargetBearing(i)
            Else
              b(i).Direction = GetAvoidanceDirection(i, nObID)
              b(i).CAvoid = True
            End If
          Else
            b(i).Direction = nDir
          End If
        End If
      End If
      b(i).X = b(i).X + b(i).Velocity * Cos(b(i).Direction)
      b(i).Y = b(i).Y + b(i).Velocity * Sin(b(i).Direction)
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
    b(i).MaxVelocity = 3
    b(i).Velocity = 0.25
    b(i).Diameter = 1
  Next i
  
  'field restraints
  nMin_X = 5
  nMax_X = 95
  nMin_Y = 5
  nMax_Y = 95
End Sub
