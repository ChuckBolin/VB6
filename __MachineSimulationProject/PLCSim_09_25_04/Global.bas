Attribute VB_Name = "Global"
Option Explicit

'public types
'this stores possible boolean fragments and their simplified equivalent
Public Type BOOLEAN_PATTERNS
  Pattern As String
  Replacement As String
End Type

'this stores a line of code
Public Type CODE_DATA
  Code As String            'original typed in code
  CleanCode As String       'cleaned up code...correct spacing
  Substitute As String      'all input/outputs substituted with 1's and 0's
  ResultString As String    'symbol of result. I.e. OUT2, OUT4
  Result As Integer          'boolean result of equation. I.e. 1, 0
End Type

'stores info for inputs, outputs and bits
Public Type BYTE_VALUES
  Absolute As String  'i.e. IN1
  Symbol As String    'i.e. S12
  Value As Integer  '       1
End Type

'mode for running or editing
Public Enum PLC_MODE
  Edit = 1
  Run
  StopPLC
End Enum

'global constants
Public Const MAX_PATTERNS = 17  'most patterns for boolean solutions
Public Const MAX_INPUTS = 8     'max inputs available
Public Const MAX_OUTPUTS = 8    'max outputs available
Public Const MAX_BITS = 8       'max markers or bits available
Public Const MAX_LINES_OF_CODE = 10

'global variables
Public g_nLines As Integer 'number of lines of code
Public bp(MAX_PATTERNS) As BOOLEAN_PATTERNS
Public g_sCode(MAX_LINES_OF_CODE) As CODE_DATA
Public g_uIn(MAX_INPUTS) As BYTE_VALUES
Public g_uOut(MAX_OUTPUTS) As BYTE_VALUES
Public g_uBit(MAX_BITS) As BYTE_VALUES
Public g_sOperators As String 'list of all valid operators
Public g_sLegalCharacters As String 'list of all non-operator valid characters


