Attribute VB_Name = "ArmCode"
Option Explicit

Public Type POINT_2D
  X As Single
  Y As Single
End Type

Public Type ROBOT_BUG
  X As Single
  Y As Single
  turn As Single
  direction As Single 'radians
  size As Single
  speed As Single
  a As POINT_2D
  b As POINT_2D
  c As POINT_2D
  elbow1 As POINT_2D 'left elbow
  wrist1 As POINT_2D 'left wrist
  elbow2 As POINT_2D 'right elbow
  wrist2 As POINT_2D 'right wrist
  leftShoulder As Single 'angle
  leftElbow As Single 'angle
  rightShoulder As Single
  rightElbow As Single
  arm As Single
  forearm As Single
  leftWristDir As Single
  leftWristSpeed As Single
  rightWristDir As Single
  rightWristSpeed As Single
  
End Type

Public bug As ROBOT_BUG   'bug information
Public target As POINT_2D 'target..destination

Public Sub LoadBug()
   bug.X = 50
   bug.Y = 50
   bug.direction = 3.9
   bug.size = 5
   bug.speed = 1
   bug.arm = 8
   bug.forearm = 5
   bug.leftShoulder = 0.78
   bug.leftElbow = -0.5
   bug.rightShoulder = -0.78
   bug.rightElbow = 0.5
   target.X = 50
   target.Y = 50
   
End Sub

Public Sub UpdateBug()
  
  Dim distance As Single
  Dim direction As Single
  
  distance = getDistance(bug.X, bug.Y, target.X, target.Y)
  direction = getDirection(bug.X, bug.Y, target.X, target.Y)
  
  If direction < 1.57 And bug.direction > 4.71 Then
    direction = direction + 6.28
  ElseIf bug.direction < 1.57 And direction > 4.71 Then
    bug.direction = bug.direction + 6.28
  End If
  
  If direction > bug.direction Then
    bug.turn = 0.1
  ElseIf direction < bug.direction Then
    bug.turn = -0.1
  Else
    bug.turn = 0
  End If
  bug.direction = bug.direction + bug.turn

  'bug.direction = direction
  
  If distance < 9 Then
    bug.speed = 0
  Else
    bug.speed = 1
  End If

  bug.X = bug.X + getVectorX(bug.speed, bug.direction)
  bug.Y = bug.Y + getVectorY(bug.speed, bug.direction)
  bug.a.X = bug.X + getVectorX(bug.size, bug.direction + 0.78)
  bug.a.Y = bug.Y + getVectorY(bug.size, bug.direction + 0.78)
  bug.b.X = bug.X + getVectorX(bug.size, bug.direction - 0.78)
  bug.b.Y = bug.Y + getVectorY(bug.size, bug.direction - 0.78)
  bug.c.X = bug.X + getVectorX(bug.size, bug.direction + 3.14)
  bug.c.Y = bug.Y + getVectorY(bug.size, bug.direction + 3.14)

  bug.elbow1.X = bug.a.X + getVectorX(bug.arm, bug.leftShoulder + bug.direction)
  bug.elbow1.Y = bug.a.Y + getVectorY(bug.arm, bug.leftShoulder + bug.direction)
  bug.elbow2.X = bug.b.X + getVectorX(bug.arm, bug.rightShoulder + bug.direction)
  bug.elbow2.Y = bug.b.Y + getVectorY(bug.arm, bug.rightShoulder + bug.direction)

  bug.wrist1.X = bug.elbow1.X + getVectorX(bug.forearm, bug.leftShoulder + bug.leftElbow + bug.direction)
  bug.wrist1.Y = bug.elbow1.Y + getVectorY(bug.forearm, bug.leftShoulder + bug.leftElbow + bug.direction)
  bug.wrist2.X = bug.elbow2.X + getVectorX(bug.forearm, bug.rightShoulder + bug.rightElbow + bug.direction)
  bug.wrist2.Y = bug.elbow2.Y + getVectorY(bug.forearm, bug.rightShoulder + bug.rightElbow + bug.direction)


End Sub

Public Function getVectorX(vector As Single, angle As Single)
  getVectorX = Cos(angle) * vector
End Function

Public Function getVectorY(vector As Single, angle As Single)
  getVectorY = Sin(angle) * vector
End Function

'returns direction from s to t
Public Function getDirection(sx As Single, sy As Single, tx As Single, ty As Single) As Single
  Dim hyp As Single
  Dim opp As Single
  Dim adj As Single
  'Dim angle As Single
  Dim direction As Single
  
  hyp = getDistance(sx, sy, tx, ty)
  adj = tx - sx
  opp = ty - sy
  If adj = 0 Then adj = 0.0001
  direction = Atn(opp / adj)
  
  If adj < 0 Then direction = direction + 3.14
  
  If direction < 0 Then
    direction = direction + 6.28
  ElseIf direction > 6.28 Then
    direction = direction - 6.28
  Else
  End If

  getDirection = direction
End Function

'returns distance from s to t
Public Function getDistance(sx As Single, sy As Single, tx As Single, ty As Single) As Single
  getDistance = Sqr((tx - sx) ^ 2 + (ty - sy) ^ 2)
End Function
