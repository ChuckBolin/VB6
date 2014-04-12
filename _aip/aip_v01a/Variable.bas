Attribute VB_Name = "Variable"
Option Explicit

'global variables
Public gsRules As String 'stores all rules
Public gnRows As Integer 'total number of rows in grid
Public gnCols As Integer 'total number of columns in grid

'create type to hold patterns and other info
Public Type Pattern
  word  As Long '32 bit can hold a 5x5 array
  rel As Long 'stores equivalent relative pattern
  wins As Integer 'number of wins attributed to this pattern
  proven As Boolean 'TRUE if pattern is legal , FALSE if pattern is not legal
End Type
