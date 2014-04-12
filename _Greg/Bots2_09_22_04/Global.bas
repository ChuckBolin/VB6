Attribute VB_Name = "Global"
Option Explicit

Public Type BOTS
  id As Integer
  x As Single
  y As Single
  dir As Single
  vel As Single
  TargetX As Single
  TargetY As Single
  TargetFound As Boolean 'true if arrived at target
  Obstacle As Boolean 'true if something in the way
  CX As Single 'coordinate to pursue until no more obstacle
  CY As Single
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

Public Const BOT_MAIN = 3
'Public Const MAX_BOTS = 20

'Public bot(MAX_BOTS) As BOTS
Public g_nRange As RECT 'stores values of search quad

Public b As New CBots  'new class

'loads all necessary bot data
Public Sub LoadBotData()
  Dim i As Integer
  Dim bRet As Boolean 'grabs function returns for 'setting...'
  
  'loads bot info
  For i = 1 To b.GetMaxBots
    bRet = b.SetDiameter(i, 3)
    bRet = b.SetX(i, GetRandomSingle(10, 90))
    bRet = b.SetY(i, GetRandomSingle(10, 90))
    bRet = b.SetTargetX(i, GetRandomSingle(10, 90))
    bRet = b.SetTargetY(i, GetRandomSingle(10, 90))
    bRet = b.SetVelocity(i, 0.5)
  Next i
 ' bRet = b.SetVelocity(BOT_MAIN, 0.5)

  
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

