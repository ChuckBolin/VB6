Attribute VB_Name = "VRobot"
'****************************************************************************
' VRobot.bas - Written by Chuck Bolin, April 2005, Team 342
' This file contains code to create and update the virtual robot. In
' addition, it contains code to process Virtual Code for autonomous
' operations.
' Subs/Functions:
'
' Public Sub LoadRobotVariables()
' Public Function Limit_Mix(intermediate_value As Integer) As Integer
' Public Function GetWheelSpeed(ByVal nPWM As Byte) As Single
' Public Sub UpdateRobot()
' Public Sub EnforceFieldBoundaries()
' Public Function LoadSampleVirtualCode() As String
' Public Function LoadVirtualCodeIntoArray(sInput As String) As Boolean
' Public Sub ProcessVirtualCode()
' Private Sub InitializeVCVariables()
' Private Sub InitializeVCVariables()
' Private Sub ResetVCVariables()
' Private Sub CreateVCVariable(sInput As String)
' Private Sub SetVCVariableValue(sInput As String)
' Private Function GetVCVariableValue(sVarName As String) As String
' Private Sub IncrementVCVariableValue()
' Private Sub DecrementVCVariableValue()
'****************************************************************************
Option Explicit

Public Const RPS_MAX = 10  'inches per second at max PWM value
Public Const RPS_UPPER_DEAD = 147 'wheel doesn't turn inside these dead band values
Public Const RPS_LOWER_DEAD = 107
Public Const FIELD_TOP = 27
Public Const FIELD_BOTTOM = -27
Public Const FIELD_LEFT = -13.5
Public Const FIELD_RIGHT = 13.5

Public Type ROBOT
  Center As COORDINATE_PAIR     'center of robot
  Length As Single
  Width As Single
  LeftFront As COORDINATE_PAIR  'left front corner
  RightFront As COORDINATE_PAIR 'right front corner
  LeftBack As COORDINATE_PAIR   'left back corner
  RightBack As COORDINATE_PAIR  'right back corner
  Reference As COORDINATE_PAIR  'this is point in front center
  Offset As Single 'angle offset (rads) for calculating corners
  Hypotenuse As Single 'distance from center to corners
  Direction As Single
  Velocity As Single
  Joy_X As Byte
  Joy_Y As Byte
  LeftMotor As Single
  RightMotor As Single
  AxleDistance As Single 'distance between left and right side
End Type

Public Type VIRTUAL_CODE_VARIABLE
  Symbol As String
  Type As String
  Scope As String
  Value As String
End Type

Public VR As ROBOT 'virtual robot
Public g_nCounter As Integer 'stores 15 seconds..counts down in auto
Public g_sProg() As String 'array to store virtual code program
Public g_nMaxLines As Integer 'max lines in code..1 greater than 0
Public Const ROBOT_MAX_VARIABLES = 200
Public g_uVCVar() As VIRTUAL_CODE_VARIABLE 'stores up to 200 variables

'*********************************************
' This is called when program starts to
' initialize the robot
'*********************************************
Public Sub LoadRobotVariables()
  VR.Center.X = 0 'initial position on field
  VR.Center.Y = 0
  VR.Length = 3  'feet
  VR.Width = 2 'feet
  VR.Offset = Atn((VR.Width / 2) / (VR.Length / 2))
  VR.Hypotenuse = Sqr((VR.Width / 2) ^ 2 + (VR.Length / 2) ^ 2)
  VR.Joy_X = 127
  VR.Joy_Y = 127
  VR.AxleDistance = 1.8 'feet
  VR.Direction = 0
  VR.Velocity = 0
  VR.LeftMotor = 127
  VR.RightMotor = 127
End Sub

'***********************************************************************
'This function duplicates the code in the default FRC code for mixing
'joystick x and y values to create two drives signals to the left and
'right motors
'***********************************************************************
Public Function Limit_Mix(intermediate_value As Integer) As Integer
  Static limited_value As Integer
  
  If intermediate_value < 2000 Then
    limited_value = 2000
  ElseIf intermediate_value > 2254 Then
    limited_value = 2254
  Else
    limited_value = intermediate_value
  End If
  
  Limit_Mix = limited_value - 2000
End Function

'***********************************************************************
' Returns distance traveled per second based upon PWM value
'***********************************************************************
Public Function GetWheelSpeed(ByVal nPWM As Byte) As Single
  Dim nSlope As Single
  'frmMain.Caption = nPWM
  If nPWM < RPS_UPPER_DEAD And nPWM > RPS_LOWER_DEAD Then
    GetWheelSpeed = 0
  ElseIf nPWM >= RPS_UPPER_DEAD Then
    nSlope = (nPWM - RPS_UPPER_DEAD) / (255 - RPS_UPPER_DEAD)
    GetWheelSpeed = nSlope * RPS_MAX
  ElseIf nPWM <= RPS_LOWER_DEAD Then
    nSlope = -((RPS_LOWER_DEAD - nPWM) / (RPS_LOWER_DEAD))
    GetWheelSpeed = nSlope * RPS_MAX
  End If
End Function

'**********************************************************************
' Updates position and orientation of virtual robot based wheel speeds
' and previous position. Called every 26 mSec (or 25 mSec)
'**********************************************************************
Public Sub UpdateRobot()
  Dim nDeltaAngle As Single 'change in angle
  Dim nLeft As Single
  Dim nRight As Single
  Dim nDeltaXY As COORDINATE_PAIR
  
  'left and right speeds
  nLeft = GetWheelSpeed(VR.LeftMotor) / 40
  nRight = GetWheelSpeed(VR.RightMotor) / 40
  
  'update direction
  nDeltaAngle = GetAngleRadiansArctan((nRight - nLeft), VR.AxleDistance)
  VR.Direction = VR.Direction + nDeltaAngle
  If VR.Direction > TWO_PI Then
    VR.Direction = 0
  ElseIf VR.Direction < 0 Then
    VR.Direction = TWO_PI
  End If
  
  'velocity is average of two wheels
  VR.Velocity = (nRight + nLeft) / 2
  nDeltaXY = GetVectorXY(VR.Direction, VR.Velocity)
  VR.Center.X = VR.Center.X + nDeltaXY.X  'calc new psotion
  VR.Center.Y = VR.Center.Y + nDeltaXY.Y
  
  'calculate new center and orientation values
  VR.Reference.X = VR.Center.X + ((VR.Length / 2) * Cos(VR.Direction))
  VR.Reference.Y = VR.Center.Y + ((VR.Length / 2) * Sin(VR.Direction))
  VR.LeftFront.X = VR.Center.X + (VR.Hypotenuse * Cos(VR.Direction + VR.Offset))
  VR.LeftFront.Y = VR.Center.Y + (VR.Hypotenuse * Sin(VR.Direction + VR.Offset))
  VR.RightFront.X = VR.Center.X + (VR.Hypotenuse * Cos(VR.Direction - VR.Offset))
  VR.RightFront.Y = VR.Center.Y + (VR.Hypotenuse * Sin(VR.Direction - VR.Offset))
  VR.LeftBack.X = VR.Center.X + (VR.Hypotenuse * Cos(VR.Direction + (3.14 - VR.Offset)))
  VR.LeftBack.Y = VR.Center.Y + (VR.Hypotenuse * Sin(VR.Direction + (3.14 - VR.Offset)))
  VR.RightBack.X = VR.Center.X + (VR.Hypotenuse * Cos(VR.Direction - (3.14 - VR.Offset)))
  VR.RightBack.Y = VR.Center.Y + (VR.Hypotenuse * Sin(VR.Direction - (3.14 - VR.Offset)))
  EnforceFieldBoundaries
End Sub

'**********************************************************
' Each time the virtual bot is updated (corners) they must
' be verified to be inside field..if they are outside of
' field then stop the motor pushing that corner
'**********************************************************
Public Sub EnforceFieldBoundaries()
  
  If VR.LeftFront.X < FIELD_LEFT Then VR.LeftMotor = 127
  If VR.LeftFront.X > FIELD_RIGHT Then VR.LeftMotor = 127
  If VR.RightFront.X < FIELD_LEFT Then VR.RightMotor = 127
  If VR.RightFront.X > FIELD_RIGHT Then VR.RightMotor = 127
  If VR.LeftBack.X < FIELD_LEFT Then VR.LeftMotor = 127
  If VR.LeftBack.X > FIELD_RIGHT Then VR.LeftMotor = 127
  If VR.RightBack.X < FIELD_LEFT Then VR.RightMotor = 127
  If VR.RightBack.X > FIELD_RIGHT Then VR.RightMotor = 127
  
  If VR.LeftFront.Y > FIELD_TOP Then VR.LeftMotor = 127
  If VR.LeftFront.Y < FIELD_BOTTOM Then VR.LeftMotor = 127
  If VR.RightFront.Y > FIELD_TOP Then VR.RightMotor = 127
  If VR.RightFront.Y < FIELD_BOTTOM Then VR.RightMotor = 127
  If VR.LeftBack.Y > FIELD_TOP Then VR.LeftMotor = 127
  If VR.LeftBack.Y < FIELD_BOTTOM Then VR.LeftMotor = 127
  If VR.RightBack.Y > FIELD_TOP Then VR.RightMotor = 127
  If VR.RightBack.Y < FIELD_BOTTOM Then VR.RightMotor = 127
  
End Sub

'**********************************************************
' Returns sample Virtual Code for loading into text box
'**********************************************************
Public Function LoadSampleVirtualCode() As String
  Dim sOut As String
  
  sOut = sOut & "CVAR i,static int,0" & vbCrLf
  sOut = sOut & "INC i,1" & vbCrLf
  sOut = sOut & "GLR i<80,4" & vbCrLf
  sOut = sOut & "SVAR pwm01,200" & vbCrLf
  sOut = sOut & "SVAR pwm02,200" & vbCrLf
  sOut = sOut & "JMP 3" & vbCrLf
  sOut = sOut & "SVAR pwm01,127" & vbCrLf
  sOut = sOut & "SVAR pwm02,127" & vbCrLf
  sOut = sOut & "END" & vbCrLf
  'sOut = sOut & "" & vbCrLf
  'sOut = sOut & "" & vbCrLf
  'sOut = sOut & "" & vbCrLf
  'sOut = sOut & "" & vbCrLf
  LoadSampleVirtualCode = sOut
End Function

'**********************************************************
' Loads program (string) into array g_sProg()
' Eliminates white spaces (trim)
' Does no code checking...assumes all is okay.
'**********************************************************
Public Function LoadVirtualCodeIntoArray(sInput As String) As Boolean
  Dim i As Integer
  Dim sLines() As String
  Dim sOut As String
  
  LoadVirtualCodeIntoArray = False
  
  'program loaded into array sLines()
  sLines = Split(sInput, vbCrLf)
  g_nMaxLines = -1
  
  'trim lines, check for substance, load into global program array g_sProg()
  For i = 0 To UBound(sLines) - 1
    sOut = Trim(sLines(i))
    If Len(sOut) > 0 Then
      g_nMaxLines = g_nMaxLines + 1
      ReDim Preserve g_sProg(g_nMaxLines)
      g_sProg(g_nMaxLines) = sOut
    End If
  Next i
  If g_nMaxLines < 0 Then Exit Function
  
  'get variables initialized - done only 1x after program loaded into array
  InitializeVCVariables
  
  LoadVirtualCodeIntoArray = True
End Function

'********************************************************
' This code processes the virtual code in g_sProg()
' every 26 mS in order to update robot behavior
'********************************************************
Public Sub ProcessVirtualCode()
  Dim i As Integer
  Dim sLine As String 'line to process
  
  'clear all variables that are not static
  ResetVCVariables
  
  'go through one line at a time
  For i = 0 To g_nMaxLines
    sLine = g_sProg(i)
  
    If Left(sLine, 5) = "CVAR " Then
    
    ElseIf Left(sLine, 4) = "INC " Then
    
    ElseIf Left(sLine, 4) = "DEC " Then
    
    ElseIf Left(sLine, 4) = "GLR " Then
    
    ElseIf Left(sLine, 4) = "JMP " Then
    
    ElseIf Left(sLine, 5) = "SVAR " Then
  
    ElseIf Left(sLine, 4) = "END " Then
  
    Else
    
    End If
  Next i
  
  'now set robot controller outputs based upon specified values
  
End Sub


'****************************************************
' Intializes all VC variables in g_uVCVar( )
' This is called one time only for autonomous mode.
'*****************************************************
Private Sub InitializeVCVariables()
  
  ReDim g_uVCVar(ROBOT_MAX_VARIABLES)
  
  'loads needed variables that will persist through auto mode
  g_uVCVar(0).Symbol = "pwm01"
  g_uVCVar(0).Scope = "static"
  g_uVCVar(0).Type = "unsigned char"
  g_uVCVar(0).Value = "127"
  g_uVCVar(1).Symbol = "pwm02"
  g_uVCVar(1).Scope = "static"
  g_uVCVar(1).Type = "unsigned char"
  g_uVCVar(1).Value = "127"
  
End Sub

'******************************************************
' Intializes ONLY non-static VC variables in g_uVCVar( )
' This is called every 26 mSec
'*******************************************************
Private Sub ResetVCVariables()
  Dim i As Integer
  Dim uEmptyVariable As VIRTUAL_CODE_VARIABLE
  
  For i = 0 To ROBOT_MAX_VARIABLES
    If g_uVCVar(i).Scope <> "static" Then
      g_uVCVar(i) = uEmptyVariable
    End If
  Next i
End Sub


'****************************************************
' Creates a variable in g_uVCVar()
' sInput contains code from
' CVAR i,static int,0
'*****************************************************
Private Sub CreateVCVariable(sInput As String)

End Sub

'****************************************************
' Sets value of particular variable in g_uVCVar()
' sInput contains code from
' SVAR i,127
'*****************************************************
Private Sub SetVCVariableValue(sInput As String)

End Sub

'****************************************************
' Gets value of particular variable in g_uVCVar()
'*****************************************************
Private Function GetVCVariableValue(sVarName As String) As String

End Function

'****************************************************
' Increment a variable value
' INC a,2
'*****************************************************
Private Sub IncrementVCVariableValue()

End Sub

'****************************************************
' Decrement a variable value
' DEC a,2
'*****************************************************
Private Sub DecrementVCVariableValue()

End Sub



