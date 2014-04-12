Attribute VB_Name = "Global"
Option Explicit

Public Type Joint
  X As Single
  Y As Single
  VX As Single
  VY As Single
End Type

Public j(10) As Joint

Public Const PI = 3.14159
Public g_X As Single
Public g_Y As Single
Public g_segment As Single

