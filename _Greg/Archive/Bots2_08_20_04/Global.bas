Attribute VB_Name = "Global"
Option Explicit

Public Type BOTS
  ID As Integer
  x As Single
  y As Single
  dir As Single
  vel As Single
  TargetX As Single
  TargetY As Single
  Diameter As Single
  CloseCount As Integer 'number of close bots
  InRange() As Integer 'ID numbers of bots within search quad
End Type

Public Type RECT
  X_Min As Single
  X_Max As Single
  Y_Min As Single
  Y_Max As Single
End Type
 
Public Type PAIR
  x As Single
  y As Single
End Type

Public Const BOT_MAIN = 0
Public Const MAX_BOTS = 50

Public bot(MAX_BOTS) As BOTS
Public nRange As RECT 'stores values of search quad

'loads all necessary bot data
Public Sub LoadBotData()
  Dim i As Integer
  
  'loads bot info
  bot(0).x = 50:  bot(0).y = 50: bot(0).TargetX = 75: bot(0).TargetY = 90: bot(0).vel = 0
  For i = 0 To MAX_BOTS
    bot(i).Diameter = 3
    bot(i).ID = i
  Next i
  For i = 1 To MAX_BOTS
    bot(i).x = GetRandomInteger(5, 95)
    bot(i).y = GetRandomInteger(5, 95)
  Next i
  
End Sub

'updates bot information such as close bots
Public Sub UpdateBots()
  Dim i As Integer
  Dim j As Integer
  Dim nIndex As Integer
    
  'find all bots within search quad
  For i = BOT_MAIN To BOT_MAIN       'bot of interest
    If i = BOT_MAIN Then
      ReDim bot(BOT_MAIN).InRange(0)
      bot(BOT_MAIN).CloseCount = 0
    End If
    For j = 0 To MAX_BOTS     'all other bots
      
      'find ID numbers of all bots within box and stores in dynamic array InRange
      If i = BOT_MAIN And i <> j Then
        If bot(j).x > bot(BOT_MAIN).x + nRange.X_Min And bot(j).x < bot(BOT_MAIN).x + nRange.X_Max And bot(j).y > bot(BOT_MAIN).y + nRange.Y_Min And bot(j).y < bot(BOT_MAIN).y + nRange.Y_Max Then
          If ObstacleExists(bot(0).x, bot(0).y, bot(0).TargetX, bot(0).TargetY, bot(0).dir, bot(j).x, bot(j).y) = True Then
            bot(BOT_MAIN).InRange(nIndex) = bot(j).ID
            bot(BOT_MAIN).CloseCount = bot(BOT_MAIN).CloseCount + 1
            nIndex = nIndex + 1
            ReDim Preserve bot(BOT_MAIN).InRange(nIndex)
          End If
        End If
      End If
    Next j
  Next i

  bot(BOT_MAIN).dir = GetTargetDirection2D(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)

End Sub

Public Function ObstacleExists(x As Single, y As Single, tx As Single, ty As Single, dir As Single, cx As Single, cy As Single) As Boolean
  Dim brg As Single
  Dim beta As Single
  Dim nDist As Single
  Dim nCPA As Single
  
  ObstacleExists = False
  nCPA = 0
  brg = GetTargetDirection2D(x, y, cx, cy)
  nDist = GetTargetDistance2D(x, y, cx, cy)
  If GetTargetDistance2D(x, y, tx, ty) < nDist Then Exit Function 'obstacle beyond target point
  beta = brg - dir
  nCPA = nDist * Sin(beta)
  If Abs(nCPA) < bot(0).Diameter Then ObstacleExists = True
  frmMain.txtData.Text = nCPA & ", " & bot(0).Diameter
  
End Function

'returns a quad that represents search pattern for contacts
Public Function GetSearchQuad(dir As Single, vel As Single, range As Single) As RECT
  Dim r As RECT
  Dim dx As Single
  Dim dy As Single
  Dim nOffset As Single 'minimum offset
  
  If range < 5 Then range = 5
  If range > 100 Then range = 100
  
  dx = Cos(dir) * (vel + 1) * range
  dy = Sin(dir) * (vel + 1) * range
  nOffset = 10
  
  If dx >= 0 Then
    r.X_Max = dx
    r.X_Min = -(vel + 1) * 5
  Else
    r.X_Max = (vel + 1) * 5
    r.X_Min = dx
  End If
  
  If dy >= 0 Then
    r.Y_Max = dy
    r.Y_Min = -(vel + 1) * 5
  Else
    r.Y_Max = (vel + 1) * 5
    r.Y_Min = dy
  End If
  
  If r.X_Max >= 0 And r.X_Max < nOffset Then r.X_Max = nOffset
  If r.X_Min < 0 And r.X_Min > -nOffset Then r.X_Min = -nOffset
  If r.Y_Max >= 0 And r.Y_Max < nOffset Then r.Y_Max = nOffset
  If r.Y_Min < 0 And r.Y_Min > -nOffset Then r.Y_Min = -nOffset
  
  
  GetSearchQuad = r
End Function

'calculates point perpendicular to line between bot and bot in the way
Public Function GetCCW(x As Single, y As Single, cx As Single, cy As Single) As PAIR
  Dim beta As Single
  beta = GetTargetDirection2D(x, y, cx, cy)
  GetCCW.x = cx - (bot(BOT_MAIN).Diameter) * Sin(beta)
  GetCCW.y = cy + (bot(BOT_MAIN).Diameter) * Cos(beta)
End Function

Public Function GetCW(x As Single, y As Single, cx As Single, cy As Single) As PAIR
  Dim beta As Single
  beta = GetTargetDirection2D(x, y, cx, cy)
  GetCW.x = cx + (bot(BOT_MAIN).Diameter) * Sin(beta)
  GetCW.y = cy - (bot(BOT_MAIN).Diameter) * Cos(beta)
End Function
