Attribute VB_Name = "Module1"
Option Explicit

'stores four values of same pattern at four different angles 0, 90, 180 and 270
Public Type Pattern
  N As Long
  E As Long
  W As Long
  S As Long
  color As Integer
  width As Integer
  height As Integer
End Type

Public p(30) As Pattern 'stores all values

'game stats
Public gintPatterns As Integer
Public gintRows As Integer
Public gintLevel As Integer
Public glngScore As Long
Public gintSeconds As Integer

Public Type Winner
  Name As String
  Score As Long
  Level As Integer
  Rows As Integer
  Patterns As Integer
End Type

Public win(3) As Winner

