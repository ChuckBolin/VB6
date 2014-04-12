Attribute VB_Name = "BotGlobal"
'********************************************************************************************************************
' BotGlobal.Bas
' Requires MathFunctions.Bas
'********************************************************************************************************************
Option Explicit

'constants
Public Const MAX_BOTS = 10
Public Const MAX_BOT_TYPES = 5  'used to distinguish them into categories
Public Const LIMIT_NORTH = 100
Public Const LIMIT_SOUTH = 0
Public Const LIMIT_WEST = 0
Public Const LIMIT_EAST = 100

'objects
Public bot(MAX_BOTS) As New CBot

'*****************************
' INITIALIZE_BOTS
'*****************************
'sets initial parameters
Public Sub InitializeBots()
  Dim i As Integer
  Randomize Timer
  For i = 0 To MAX_BOTS
    bot(i).X = GetRandomSingle(5, 95)
    bot(i).Y = GetRandomSingle(5, 95)
    bot(i).TX = GetRandomSingle(5, 95)
    bot(i).TY = GetRandomSingle(5, 95)
    bot(i).BotType = GetRandomSingle(0, MAX_BOT_TYPES)
    bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
    'MsgBox bot(i).X & vbTab & bot(i).Y & vbTab & bot(i).TX & vbTab & bot(i).TY & vbTab & bot(i).Direction & vbCrLf
    bot(i).Speed = 1
  Next i
End Sub

'*****************************
' UPDATE_BOTS
'*****************************
'updates X and Y position of
'bots based upon speed, dir
Public Sub UpdateBots()
  Dim i As Integer
  Dim nDir As Single 'stores direction
    
  For i = 0 To MAX_BOTS
    If bot(i).Found = False Then
      bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
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
      If bot(i).Y < LIMIT_SOUTH Then bot(i).Found = True
      If bot(i).Y > LIMIT_NORTH Then bot(i).Found = True
      
      'create new target x,y values
      If bot(i).Found = True Then
          bot(i).TX = GetRandomSingle(5, 95)
          bot(i).TY = GetRandomSingle(5, 95)
          bot(i).Found = False
      End If
    End If
  Next i
End Sub
