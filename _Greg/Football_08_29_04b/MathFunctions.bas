Attribute VB_Name = "MathFunctions"
'***************************************************************
' MathFunctions.bas - Written by Chuck Bolin, July 16, 2004
' DtoR() - Converts computer degrees to computer radians radians
' RtoD() - Converts computer radians to computer degrees
' DtoC() - Converts computer degrees to compass degrees
' RtoC() - Converts computer radians to compass degrees
' CtoR() - Converts compass degrees to computer radians
' CRtoR() - Converts compass radians to computer radians
' GetTargetDirection2D() - Calculates direction from x,y to tx,ty
' GetTargetDistance2D() - Returns distance between to 2D points
' Updated: August 4, 2004, CB
' GetRandomInteger(min, max) - Returns random integer within range
' GetRandomSingle (min, max) - Returns random single within range
' Updated: August 5, 2004, CB
' GetAngleRadiansArctan(opposite, adjacent) - Returns angle
'
'   NOTE:
'   Perform all trig calculations in computer radians.
'   Use other rad/deg/compass conversion for display
'   purposes. Saves you a lot of heartache. CB - 8/5/2004
'****************************************************************
Option Explicit

'public constants
Public Const PI = 3.14159

'public types
Public Type COORDINATE_PAIR
  X As Single
  Y As Single
End Type

'Converts computer Degrees to computer Radians
Public Function DtoR(deg As Single) As Single
  DtoR = deg * PI / 180
End Function

'Converts computer Radians to computer Degrees
Public Function RtoD(rad As Single) As Single
  RtoD = rad * 180 / PI
  If RtoD > 360 Then RtoD = RtoD - 360
  If RtoD < 0 Then RtoD = RtoD + 360
End Function

'Converts computer Degrees to Compass Degrees
Public Function DtoC(deg As Single) As Single
  DtoC = 450 - deg
  
  'ensure deg is between 0 and 360
  If DtoC > 360 Then DtoC = DtoC - 360
  If DtoC < 0 Then DtoC = DtoC + 360
End Function

'Converts computer Radians to Compass Degrees
Public Function RtoC(rad As Single) As Single
  'RtoC = rad * 180 / PI - 450
   RtoC = 450 - (rad * 180) / PI
   
  If RtoC > 360 Then RtoC = RtoC - 360
  If RtoC < 0 Then RtoC = RtoC + 360
End Function

'converts Compass degrees to computer radians
Public Function CtoR(compass As Single) As Single
  CtoR = (450 - compass) * PI / 180
End Function

'converts Compass radians to computer radians
Public Function CRtoR(compass As Single) As Single
  CRtoR = (7.85 - compass) '*  3.14159 / 180
End Function

'*****************************
' GET_TARGET_DISTANCE_2D
'*****************************
'calculates distance from X,Y
'to target X,Y
Public Function GetTargetDistance2D(X As Single, Y As Single, TX As Single, TY As Single) As Single
  GetTargetDistance2D = Sqr((TX - X) * (TX - X) + (TY - Y) * (TY - Y))
End Function

'*****************************
' GET_TARGET_DIRECTION_2D
'*****************************
'Calcs direction in computer radians
' from X,Y to a target X,Y
Public Function GetTargetDirection2D(X As Single, Y As Single, TX As Single, TY As Single) As Single
  Dim dx As Single
  Dim dy As Single
  Dim dir As Single
  
  dy = TY - Y   'deltas...target my x,y position
  dx = TX - X
  
  If dy > 0 And dx > 0 Then 'both positive...quadrant I
    GetTargetDirection2D = Atn(dy / dx)
  ElseIf dy > 0 And dx < 0 Then 'quadrant II
    GetTargetDirection2D = PI - Atn(dy / -dx)
  ElseIf dy < 0 And dx < 0 Then 'quadrant III
    GetTargetDirection2D = PI + Atn(dy / dx)
  ElseIf dy < 0 And dx > 0 Then 'quadrant IV
    GetTargetDirection2D = 2 * PI - Atn(-dy / dx)
  ElseIf dy = 0 And dx = 0 Then 'on top of each other
    GetTargetDirection2D = 0
  ElseIf dy = 0 And dx > 0 Then 'at 0 radians
    GetTargetDirection2D = 0
  ElseIf dy = 0 And dx < 0 Then 'at 3.14159 radians
    GetTargetDirection2D = PI
  ElseIf dy > 0 And dx = 0 Then 'at 1.5708 radians
    GetTargetDirection2D = PI / 2
  ElseIf dy < 0 And dx = 0 Then 'at 4.7124 radians
    GetTargetDirection2D = PI + PI / 2
  Else
    '?
  End If
  'keep values between 0 and 2*PI
  If GetTargetDirection2D > 2 * PI Then GetTargetDirection2D = GetTargetDirection2D - 2 * PI
  If GetTargetDirection2D < 0 Then GetTargetDirection2D = GetTargetDirection2D + 2 * PI
  
End Function

'*****************************
' GET_RANDOM_INTEGER
'*****************************
'returns random whole integer
'number within prescribed range
Public Function GetRandomInteger(nMin As Integer, nMax As Integer) As Integer
  Dim nTemp As Integer 'temp variable
  
  If VarType(nMin) <> vbInteger Then Exit Function 'must be an integer
  If VarType(nMax) <> vbInteger Then Exit Function
  
  'swap if nMin is greater than nMax
  If nMin > nMax Then
    nTemp = nMax
    nMax = nMin
    nMin = nTemp
  End If
  
  'return same value if min and max are the same
  If nMin = nMax Then
    GetRandomInteger = nMin
    Exit Function
  End If
  
  'produce randomized integer
  'Randomize Timer
  GetRandomInteger = nMin + CInt((Rnd * (nMax - 1)))
  
  'just in case value is over/under max/min
  If GetRandomInteger > nMax Then GetRandomInteger = nMax
  If GetRandomInteger < nMin Then GetRandomInteger = nMin
  
End Function

'*****************************
' GET_RANDOM_SINGLE
'*****************************
'returns random single
'number within prescribed range
Public Function GetRandomSingle(nMin As Single, nMax As Single) As Single
  Dim nTemp As Integer 'temp variable
  
  If VarType(nMin) <> vbSingle Then Exit Function 'must be an integer
  If VarType(nMax) <> vbSingle Then Exit Function
  
  'swap if nMin is greater than nMax
  If nMin > nMax Then
    nTemp = nMax
    nMax = nMin
    nMin = nTemp
  End If
  
  'return same value if min and max are the same
  If nMin = nMax Then
    GetRandomSingle = nMin
    Exit Function
  End If
  
  'produce randomized integer
  'Randomize Timer
  GetRandomSingle = nMin + (Rnd * (nMax - 1))
  
  'just in case value is over/under max/min
  If GetRandomSingle > nMax Then GetRandomSingle = nMax
  If GetRandomSingle < nMin Then GetRandomSingle = nMin
  
End Function

'******************************************
'  GET_ANGLE_RADIANS_ARCTAN
'******************************************
'returns angle in radians given
'opp and adj sides of triangle
Public Function GetAngleRadiansArctan(opp As Single, adj As Single) As Single
  If adj = 0 Then adj = 0.001
  GetAngleRadiansArctan = Atn(opp / adj)
End Function






