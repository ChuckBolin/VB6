Attribute VB_Name = "MathFunctions"
'***************************************************************
' MathFunctions.bas - Written by Chuck Bolin, July 16, 2004
' DtoR() - Converts degrees to radians
' RtoD() - Converts radians to degrees
' DtoC() - Converts degrees to compass degrees
' RtoC() - Converts radians to compass degrees
' CtoR() - Converts compass degrees to radians
' GetTargetDirection() - Calculates direction from x,y to tx,ty
'****************************************************************

'Converts Degrees to Radians
Public Function DtoR(deg As Single) As Single
  DtoR = deg * 3.14159 / 180
End Function

'Converts Radians to Degrees
Public Function RtoD(rad As Single) As Single
  RtoD = rad * 180 / 3.14159
  If RtoD > 360 Then RtoD = RtoD - 360
  If RtoD < 0 Then RtoD = RtoD + 360
End Function

'Converts Degrees to Compass Degrees
Public Function DtoC(deg As Single) As Single
  DtoC = 450 - deg
  
  'ensure deg is between 0 and 360
  If DtoC > 360 Then DtoC = DtoC - 360
  If DtoC < 0 Then DtoC = DtoC + 360
End Function

'Converts Radians to Compass Degrees
Public Function RtoC(rad As Single) As Single
  RtoC = rad * 180 / 3.14159 - 450
  
  If RtoC > 360 Then RtoC = RtoC - 360
  If RtoC < 0 Then RtoC = RtoC + 360
End Function

'converts Compass degrees to radians
Public Function CtoR(compass As Single) As Single
  CtoR = (450 - compass) * 3.14159 / 180
End Function

'Calcs direction (in radians) from X,Y to a target X,Y
Public Function GetTargetDirection(x As Single, y As Single, tx As Single, ty As Single)
  Dim dx As Single
  Dim dy As Single
  
  dy = ty - y   'deltas...target my x,y position
  dx = tx - x
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetTargetDirection = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetTargetDirection = 3.14159 + Atn(dy / dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetTargetDirection = 3.14159 + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetTargetDirection = Atn(dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetTargetDirection = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetTargetDirection = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetTargetDirection = 3.14159
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetTargetDirection = 1.5708
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetTargetDirection = 4.7124
  Else
    '?
  End If
  
  If GetTargetDirection > 6.2832 Then GetTargetDirection = GetTargetDirection - 6.2832
  If GetTargetDirection < 0 Then GetTargetDirection = GetTargetDirection + 6.2832
  GetTargetDirection = 1.57 - GetTargetDirection
  'MsgBox Format(dy, "##.#") & "      " & Format(dx, "##.#") & "     " & Format(GetTargetDirection, "#.#")
  
End Function

