Attribute VB_Name = "Module1"
Option Explicit

Public Const PI = 3.14159
Public Const MAX_VELOCITY = 500
Public Const MIN_VELOCITY = -500
Public Const STEP_VELOCITY = 20
Public Const SHIP_LENGTH = 300

Public Type MOVING_OBJECT
  X As Single
  Y As Single
  Velocity As Single
  Angle As Single
End Type

Public S As MOVING_OBJECT


