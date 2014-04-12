Attribute VB_Name = "Module1"
Option Explicit

Public Const PI = 3.14159
Public Const MAX_VELOCITY = 20
Public Const MIN_VELOCITY = -20
Public Const STEP_VELOCITY = 2
Public Const SHIP_LENGTH = 300

Public Type MOVING_OBJECT
  X As Single
  Y As Single
  Velocity As Single  'speed of X,Y point of object
  Heading As Single   'actual direction of moving object
  Thrust As Single    '<> 0 if thrusting
  Angle As Single     'direction ship is physically pointed
End Type

Public S As MOVING_OBJECT


'*****************************
' GET_TARGET_DIRECTION_2D
'*****************************
'Calcs direction in computer radians
' from X,Y to a target X,Y
Public Function GetAngle(ByVal dx As Single, ByVal dy As Single) As Single
  'Dim dx As Single
  'Dim dy As Single
  Dim dir As Single
  
  'dy = TY - Y   'deltas...target my x,y position
  'dx = TX - X
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetAngle = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetAngle = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetAngle = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetAngle = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetAngle = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetAngle = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetAngle = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetAngle = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetAngle = PI + PI / 2
  Else
    '?
  End If
  'keep values between 0 and 2*PI
  If GetAngle > 2 * PI Then GetAngle = GetAngle - 2 * PI
  If GetAngle < 0 Then GetAngle = GetAngle + 2 * PI
  
End Function
