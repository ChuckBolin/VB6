Attribute VB_Name = "Module1"
Option Explicit

'stores four values of same pattern at four different angles 0, 90, 180 and 270
Public Type Pattern
  N As Long
  E As Long
  W As Long
  S As Long
  color As Integer
End Type

Public p(20) As Pattern 'stores all values
