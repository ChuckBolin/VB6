Attribute VB_Name = "BotGlobal"
'********************************************************************************************************************
' BotGlobal.Bas
' Requires MathFunctions.Bas
'********************************************************************************************************************
Option Explicit

'constants
Public Const MAX_BOTS = 50
Public Const MAX_BOT_TYPES = 2  'used to distinguish them into categories
Public Const LIMIT_NORTH = 100
Public Const LIMIT_SOUTH = 0
Public Const LIMIT_WEST = 0
Public Const LIMIT_EAST = 100
Public Const RANGE_NEAR = 5

'objects
Public bot(MAX_BOTS) As New CBot

'*****************************
' INITIALIZE_BOTS
'*****************************
'sets initial parameters
Public Sub InitializeBots()
  Dim i As Integer
  Randomize Timer
  For i = 0 To MAX_BOTS - 1
    bot(i).X = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
    bot(i).Y = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
    bot(i).TX = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
    bot(i).TY = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
    bot(i).BotType = 2 'GetRandomSingle(1, MAX_BOT_TYPES)
    bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
    bot(i).Diameter = 4
    bot(i).StuckX = bot(i).X
    bot(i).StuckY = bot(i).Y
    If bot(i).BotType = 1 Then
      bot(i).Speed = 0.7
    Else
      bot(i).Speed = 0
    End If
  Next i
  
  'only one bot moves
  bot(MAX_BOTS).X = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
  bot(MAX_BOTS).Y = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
  bot(MAX_BOTS).TX = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
  bot(MAX_BOTS).TY = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
  bot(MAX_BOTS).BotType = 1 'GetRandomSingle(1, MAX_BOT_TYPES)
  bot(MAX_BOTS).Direction = GetTargetDirection2D(bot(MAX_BOTS).X, bot(MAX_BOTS).Y, bot(MAX_BOTS).TX, bot(MAX_BOTS).TY)
  bot(MAX_BOTS).Diameter = 4
  bot(MAX_BOTS).StuckX = bot(MAX_BOTS).X
  bot(MAX_BOTS).StuckY = bot(MAX_BOTS).Y
  If bot(MAX_BOTS).BotType = 1 Then
    bot(MAX_BOTS).Speed = 0.7
  Else
    bot(MAX_BOTS).Speed = 0
  End If
  
  
End Sub

'*****************************
' UPDATE_BOTS
'*****************************
'updates X and Y position of
'bots based upon speed, dir
Public Sub UpdateBots()
  Dim i As Integer
  Dim j As Integer
  Dim nDir As Single 'stores direction
  Dim nRange As Single 'range to other bots
  Dim nTheta As Single
  Dim nBrg As Single 'bearing to other bot
  Dim nObstacleCt As Integer 'counts obstacles
  
  For i = 0 To MAX_BOTS
    'If bot(i).Found = False Then
      DoEvents
      
      'determine if line of sight has an obstacle
      nObstacleCt = 0
      For j = 0 To MAX_BOTS
        
        If i <> j Then  'evaluate other bots
          nRange = GetTargetDistance2D(bot(i).X, bot(i).Y, bot(j).X, bot(j).Y)
          If nRange < RANGE_NEAR Then
            nBrg = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(j).X, bot(j).Y)
            nTheta = GetAngleRadiansArctan(bot(i).Diameter, nRange)
            If bot(i).Direction < nBrg + nTheta And bot(i).Direction > nBrg - nTheta Then
              bot(i).Obstacle = True
              nObstacleCt = nObstacleCt + 1
              bot(i).CX = bot(i).X + (nRange * Cos(nBrg + nTheta))
              bot(i).CY = bot(i).Y + (nRange * Sin(nBrg + nTheta))
            End If
        End If
      End If
    Next j
    If nObstacleCt = 0 Then bot(i).Obstacle = False
      
      'calc direction differently...collision or not
      If bot(i).Obstacle = False Then
        bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
      ElseIf bot(i).Obstacle = True And bot(i).Stuck = False Then
        bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).CX, bot(i).CY)
      Else
        '?
      End If
      
      nDir = bot(i).Direction
      bot(i).DX = bot(i).Speed * Cos(nDir)
      bot(i).DY = bot(i).Speed * Sin(nDir)
      bot(i).X = bot(i).X + bot(i).DX
      bot(i).Y = bot(i).Y + bot(i).DY
      
      'stop movement if at target X,Y
      If GetTargetDistance2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY) < 2 Then bot(i).Found = True
      
      'stop movement if at boundary of picture box
      If bot(i).X > LIMIT_EAST Then bot(i).Found = True
      If bot(i).X < LIMIT_WEST Then bot(i).Found = True
      If bot(i).Y < LIMIT_SOUTH Then bot(i).Found = True: bot(i).Y = LIMIT_SOUTH
      If bot(i).Y > LIMIT_NORTH Then bot(i).Found = True: bot(i).Y = LIMIT_NORTH
      
      'create new target x,y values
      If bot(i).Found = True Then
          bot(i).TX = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
          bot(i).TY = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
          bot(i).Found = False
      End If
    'End If
     
     'keeps bot from getting stuck
     If i = MAX_BOTS Then
     bot(i).StuckCount = bot(i).StuckCount + 1
     If bot(i).StuckCount > 5 Then
       bot(i).StuckCount = 0
       bot(i).Stuck = False
       If GetTargetDistance2D(bot(i).X, bot(i).Y, bot(i).StuckX, bot(i).StuckY) < 2 * bot(i).Diameter Then
         bot(i).TX = GetRandomSingle(LIMIT_WEST, LIMIT_EAST)
         bot(i).TY = GetRandomSingle(LIMIT_SOUTH, LIMIT_NORTH)
         bot(i).StuckX = bot(i).X
         bot(i).StuckY = bot(i).Y
         bot(i).Stuck = True
       End If
     End If
     End If
     
    
  Next i
End Sub
