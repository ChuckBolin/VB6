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
Private Const MAX_BOTS = 100
Private Const PI = 3.14159

'type of bot info
Private Type BOT_INFO
  X As Single
  Y As Single
  TargetX As Single
  TargetY As Single
  Velocity As Single
  MaxVelocity As Single
  Direction As Single
  Diameter As Single
  CX As Single            'collision avoidance coordinate
  CY As Single
End Type

'create 100 bots
Private b(MAX_BOTS) As BOT_INFO

'************************************** GetMaxBots
'returns the max bots in class
Public Function GetMaxBots() As Integer
  GetMaxBots = MAX_BOTS
End Function

'************************************** X, Y coordinates
'Sets X position
Public Sub SetX(id As Integer, X As Single)
  If id < 1 Or id > MAX_BOTS Then Exit Sub
  b(id).X = X
End Sub

'Returns X position
Public Function GetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetX = b(id).X
End Function

'Sets Y position
Public Sub SetY(id As Integer, Y As Single)
  If id < 1 Or id > MAX_BOTS Then Exit Sub
  b(id).Y = Y
End Sub

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

'************************************** MaxVelocity
Public Function SetMaxVelocity(id As Integer, vel As Single)
  If id < 1 Or id > MAX_BOTS Then Exit Function
  b(id).MaxVelocity = vel
End Function

'************************************** GetMaxVelocity
Public Function GetMaxVelocity(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetMaxVelocity = b(id).MaxVelocity
End Function

'************************************** Direction
Public Function GetDirection(id As Integer)
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetDirection = b(id).Direction
End Function

'************************************** TargetX, TargetY coordinates
'Sets Target X position
Public Sub SetTargetX(id As Integer, X As Single)
  If id < 1 Or id > MAX_BOTS Then Exit Sub
  b(id).TargetX = X
End Sub

'Returns Target X position
Public Function GetTargetX(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetX = b(id).TargetX
End Function

'Sets Target Y position
Public Sub SetTargetY(id As Integer, Y As Single)
  If id < 1 Or id > MAX_BOTS Then Exit Sub
  b(id).TargetY = Y
End Sub

'Returns Target y position
Public Function GetTargetY(id As Integer) As Single
  If id < 1 Or id > MAX_BOTS Then Exit Function
  GetTargetY = b(id).TargetY
End Function

'****************************************** UpdateBots
'*****************************************************
'Performs all required updates for bot movements
'based upon current position, required behavior, direction
'and velocity.
Public Sub UpdateBots()

End Sub

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
    If b(i).X > x1 And b(i).X < x2 And b(i).Y < y1 And b(i).Y > y2 Then
      ReDim Preserve vBots(UBound(vBots) + 1)
      vBots(UBound(vBots)) = i
    End If
  Next i
  
  'returns variant array of bot IDs in the region
  GetBotsInRegion = vBots
End Function

'************************************* ObstacleExists
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
          nTemp = GetBotDistance(id, vBots(i))
          If nTemp < nShortDist Then
            nShortDist = nTemp
            nShortID = vBots(id)
          End If
      Next i
    End If
  End If
  
  ObstacleID = nShortID
  
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
Private Sub GetCCW(id1 As Integer, id2 As Integer)
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Function
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Function
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  b(id1).CX = b(id2).X - (b(id1).Diameter) * Sin(nBeta)
  b(id1).CY = b(id2).Y + (b(id1).Diameter) * Cos(nBeta)
End Sub

'****************************************** GetCW
Private Sub GetCW(id1 As Integer, id2 As Integer)
  If id1 < 1 Or id1 > MAX_BOTS Then Exit Sub
  If id2 < 1 Or id2 > MAX_BOTS Then Exit Sub
  
  Dim nBeta As Single
  
  nBeta = GetBotDirection(id1, id2)
  b(id1).CX = b(id2).X + (b(id1).Diameter) * Sin(nBeta)
  b(id1).CY = b(id2).Y - (b(id1).Diameter) * Cos(nBeta)
End Sub

'****************************************** Initialize Class
Private Sub Class_Initialize()
  Dim i As Integer
  
  For i = 1 To MAX_BOTS
    b(i).MaxVelocity = 3
    b(i).Velocity = 1
    b(i).Diameter = 2
  Next i
End Sub
