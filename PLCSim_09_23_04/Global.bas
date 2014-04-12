Attribute VB_Name = "Global"
Option Explicit

Public Type BOOLEAN_PATTERNS
  Pattern As String
  Replacement As String
End Type

Public Type CODE_DATA
  Code As String            'original typed in code
  CleanCode As String       'cleaned up code...correct spacing
  Substitute As String      'all input/outputs substituted with 1's and 0's
  ResultString As String    'symbol of result. I.e. OUT2, OUT4
  Result As Integer          'boolean result of equation. I.e. 1, 0
  
End Type

Public Type BYTE_VALUES
  Symbol As String  'i.e. IN1
  Value As Integer  '       1
End Type

'stores 8 bits of each input and output
'Public g_byIn As Byte
'Public g_byOut As Byte
Public g_nLines As Integer 'number of lines of code

Public Const MAX_PATTERNS = 15

Public bp(MAX_PATTERNS) As BOOLEAN_PATTERNS
Public g_sCode(10) As CODE_DATA
Public g_uIn(7) As BYTE_VALUES
Public g_uOut(7) As BYTE_VALUES


