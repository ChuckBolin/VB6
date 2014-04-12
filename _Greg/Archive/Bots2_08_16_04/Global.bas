Attribute VB_Name = "Global"
Option Explicit

Public Type BOTS
  ID As Integer
  X As Single
  Y As Single
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
 

Public Const BOT_MAIN = 0
Public Const MAX_BOTS = 5

Public bot(MAX_BOTS) As BOTS
Public nRange As RECT 'stores values of search quad

'loads all necessary bot data
Public Sub LoadBotData()
  Dim i As Integer
  
  'loads bot info
  bot(0).X = 50:  bot(0).Y = 50: bot(0).TargetX = 75: bot(0).TargetY = 90: bot(0).vel = 0
  For i = 0 To MAX_BOTS
    bot(i).Diameter = 3
    bot(i).ID = i
  Next i
  For i = 1 To MAX_BOTS
    bot(i).X = GetRandomInteger(5, 95)
    bot(i).Y = GetRandomInteger(5, 95)
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
        If bot(j).X > bot(BOT_MAIN).X + nRange.X_Min And bot(j).X < bot(BOT_MAIN).X + nRange.X_Max And bot(j).Y > bot(BOT_MAIN).Y + nRange.Y_Min And bot(j).Y < bot(BOT_MAIN).Y + nRange.Y_Max Then
          bot(BOT_MAIN).InRange(nIndex) = bot(j).ID
          bot(BOT_MAIN).CloseCount = bot(BOT_MAIN).CloseCount + 1
          nIndex = nIndex + 1
          ReDim Preserve bot(BOT_MAIN).InRange(nIndex)
        End If
      End If
    Next j
  Next i

  bot(BOT_MAIN).dir = GetTargetDirection2D(bot(BOT_MAIN).X, bot(BOT_MAIN).Y, bot(BOT_MAIN).TargetX, bot(BOT_MAIN).TargetY)

End Sub

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

