Attribute VB_Name = "Variables"
Public Const BOT_FOUND = 1
Public Const BOT_FINDING = 2
Public Type BOT
  x As Single  'current position
  y As Single
  tx As Single 'target destination
  ty As Single
  rx As Single 'random destination if stuck
  ry As Single
  Stuck As Boolean 'true if stuck
  Count As Integer 'counts steps
  Vel As Single
  Dir As Single 'direction in radians
  Range As Single 'search range
  Mode As Integer '
End Type
