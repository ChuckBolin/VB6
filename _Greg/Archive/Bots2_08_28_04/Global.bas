Attribute VB_Name = "Global"
Option Explicit

Public Type BOTS
  id As Integer
  X As Single
  Y As Single
  dir As Single
  vel As Single
  TargetX As Single
  TargetY As Single
  TargetFound As Boolean 'true if arrived at target
  Obstacle As Boolean 'true if something in the way
  cx As Single 'coordinate to pursue until no more obstacle
  cy As Single
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
  X As Single
  Y As Single
End Type

Public Const BOT_MAIN = 0
Public Const MAX_BOTS = 20

Public bot(MAX_BOTS) As BOTS
Public nRange As RECT 'stores values of search quad

Public b As New CBots  'new class

'loads all necessary bot data
Public Sub LoadBotData()
  Dim i As Integer
  
  'loads bot info
  bot(0).X = 50:  bot(0).Y = 50: bot(0).TargetX = 75: bot(0).TargetY = 90: bot(0).vel = 0.5: bot(0).TargetFound = False: bot(0).Obstacle = False
  For i = 0 To MAX_BOTS
    bot(i).Diameter = 3
    bot(i).id = i
  Next i
  For i = 1 To MAX_BOTS
    bot(i).X = GetRandomInteger(5, 95)
    bot(i).Y = GetRandomInteger(5, 95)
  Next i
  
  'loads new class CBots
  Dim sB As String
  
  For i = 1 To b.GetMaxBots
    b.SetX i, GetRandomSingle(5, 95)
    b.SetY i, GetRandomSingle(5, 95)
    'sB = sB & b.GetX(i) & ", " & b.GetY(i) & vbCrLf
  Next i
  'MsgBox sB
  
  'Dim vData() As Variant
  'vData = b.GetBotsInRegion(0, 100, 100, 0)
  'MsgBox UBound(vData)
  
  
  
  
  
End Sub

'updates bot information such as close bots
Public Sub UpdateBots()
  Dim i As Integer
  Dim j As Integer
  Dim nIndex As Integer
  Dim uPair As PAIR
  
  
  
  'find all bots within search quad
  For i = BOT_MAIN To BOT_MAIN       'bot of interest
    If i = BOT_MAIN Then
      ReDim bot(BOT_MAIN).InRange(0)
      bot(BOT_MAIN).CloseCount = 0
      'bot(BOT_MAIN).Obstacle = True
      'bot(BOT_MAIN).TargetFound = True
    End If
    
    For j = 0 To MAX_BOTS     'all other bots
      
      'find ID numbers of all bots within box and stores in dynamic array InRange
      If i = BOT_MAIN And i <> j Then
        'if j bot is inside search quad
        If bot(j).X > bot(BOT_MAIN).X + nRange.X_Min And bot(j).X < bot(BOT_MAIN).X + nRange.X_Max And bot(j).Y > bot(BOT_MAIN).Y + nRange.Y_Min And bot(j).Y < bot(BOT_MAIN).Y + nRange.Y_Max Then
          
          'if j bot is between main bot and target
          If ObstacleExists(bot(0).X, bot(0).Y, bot(0).TargetX, bot(0).TargetY, bot(0).dir, bot(j).X, bot(j).Y) = True Then
            
            bot(BOT_MAIN).CloseCount = bot(BOT_MAIN).CloseCount + 1
             ReDim Preserve bot(BOT_MAIN).InRange(bot(BOT_MAIN).CloseCount)
             bot(BOT_MAIN).InRange(bot(BOT_MAIN).CloseCount) = bot(j).id
             bot(BOT_MAIN).Obstacle = True
             uPair = GetCCW(bot(BOT_MAIN).X, bot(BOT_MAIN).Y, bot(j).X, bot(j).Y)
             bot(BOT_MAIN).cx = uPair.X 'target position
             bot(BOT_MAIN).cy = uPair.Y
             Exit For
          Else
            bot(BOT_MAIN).Obstacle = False
            bot(BOT_MAIN).cx = bot(BOT_MAIN).TargetX
            bot(BOT_MAIN).cy = bot(BOT_MAIN).TargetY
          End If
        End If
      End If
    Next j
  If i = BOT_MAIN Then
    If bot(BOT_MAIN).Obstacle = False Then 'no obstacle
      bot(BOT_MAIN).dir = GetTargetDirection2D(bot(BOT_MAIN).X, bot(BOT_MAIN).Y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)
    Else
      bot(BOT_MAIN).dir = GetTargetDirection2D(bot(BOT_MAIN).X, bot(BOT_MAIN).Y, bot(BOT_MAIN).cx, bot(BOT_MAIN).cy)
    End If
    bot(BOT_MAIN).X = bot(BOT_MAIN).X + bot(BOT_MAIN).vel * Cos(bot(BOT_MAIN).dir)
    bot(BOT_MAIN).Y = bot(BOT_MAIN).Y + bot(BOT_MAIN).vel * Sin(bot(BOT_MAIN).dir)
  
  '  If bot(BOT_MAIN).Obstacle = True Then
     ' If GetTargetDistance2D(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY) < GetTargetDistance2D(bot(BOT_MAIN).cx, bot(BOT_MAIN).cy, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY) Then
     '   bot(BOT_MAIN).Obstacle = True
     ' End If
   ' End If
    
    'If GetTargetDistance2D(bot(BOT_MAIN).x, bot(BOT_MAIN).y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY) < bot(BOT_MAIN).Diameter / 2 Then
    '  bot(BOT_MAIN).TargetFound = True
      'bot(BOT_MAIN).Obstacle = False
    '  bot(BOT_MAIN).vel = 0
    'End If
    'is bot at target?
    'If bot(BOT_MAIN).x > bot(BOT_MAIN).TargetX - bot(BOT_MAIN).Diameter And bot(BOT_MAIN).x < bot(BOT_MAIN).TargetX + bot(BOT_MAIN).Diameter Then
    '  If bot(BOT_MAIN).y > bot(BOT_MAIN).TargetY - bot(BOT_MAIN).Diameter And bot(BOT_MAIN).y < bot(BOT_MAIN).TargetY + bot(BOT_MAIN).Diameter Then
    '    bot(BOT_MAIN).TargetFound = True
    '    bot(BOT_MAIN).Obstacle = False
    '  End If
    'End If
    
    'If bot(BOT_MAIN).Obstacle = True And bot(BOT_MAIN).x > bot(BOT_MAIN).cx - bot(BOT_MAIN).Diameter * 2 And bot(BOT_MAIN).x < bot(BOT_MAIN).cx + bot(BOT_MAIN).Diameter * 2 Then
    '  If bot(BOT_MAIN).y > bot(BOT_MAIN).cy - bot(BOT_MAIN).Diameter * 2 And bot(BOT_MAIN).y < bot(BOT_MAIN).cy + bot(BOT_MAIN).Diameter * 2 Then
    '    'bot(BOT_MAIN).TargetFound = True
    '    bot(BOT_MAIN).Obstacle = False
    '    bot(BOT_MAIN).cx = bot(BOT_MAIN).TargetX
    '    bot(BOT_MAIN).cy = bot(BOT_MAIN).TargetY
    '  End If
    'End If
  
  End If
    

  Next i


End Sub

Public Function ObstacleExists(X As Single, Y As Single, TX As Single, TY As Single, dir As Single, cx As Single, cy As Single) As Boolean
  Dim brg As Single
  Dim beta As Single
  Dim nDist As Single
  Dim nCPA As Single
  
  ObstacleExists = False
  nCPA = 0
  brg = GetTargetDirection2D(X, Y, cx, cy)
  nDist = GetTargetDistance2D(X, Y, cx, cy)
  If GetTargetDistance2D(X, Y, TX, TY) < nDist Then Exit Function 'obstacle beyond target point
  beta = brg - dir
  nCPA = nDist * Sin(beta)
  If Abs(nCPA) < bot(0).Diameter Then ObstacleExists = True
  'frmMain.txtData.Text = nCPA & ", " & bot(0).Diameter
  
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
Public Function GetCCW(X As Single, Y As Single, cx As Single, cy As Single) As PAIR
  Dim beta As Single
  beta = GetTargetDirection2D(X, Y, cx, cy)
  GetCCW.X = cx - (bot(BOT_MAIN).Diameter) * Sin(beta)
  GetCCW.Y = cy + (bot(BOT_MAIN).Diameter) * Cos(beta)
End Function

Public Function GetCW(X As Single, Y As Single, cx As Single, cy As Single) As PAIR
  Dim beta As Single
  beta = GetTargetDirection2D(X, Y, cx, cy)
  GetCW.X = cx + (bot(BOT_MAIN).Diameter) * Sin(beta)
  GetCW.Y = cy - (bot(BOT_MAIN).Diameter) * Cos(beta)
End Function
