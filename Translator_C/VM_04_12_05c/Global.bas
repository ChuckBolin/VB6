Attribute VB_Name = "Global"
'****************************************************************************
' Global.bas - Written by Chuck Bolin, April 2005
' All global constants, types, enums and variables are here as well as
' global functions and subs not covered in other .bas files.
' Functions:
' LoadVariables() Loads system variables
'****************************************************************************

Option Explicit

'public constants
Public Const MAX_CODE_LINES = 300
Public Const MAX_VARIABLES = 300

'public types
Public Type ROBOT_AGENT
  X As Single
  Y As Single
  Direction As Single
  Speed As Single
End Type

Public Type VARIABLES
  Symbol As String
  Type As String
  Value As String
  Scope As String 'automatic (created/deleted) or static (never deleted)
End Type

'public variables
Public robot As ROBOT_AGENT   'stores all robot info
Public var(MAX_VARIABLES) As VARIABLES  'stores all VM code variables
Public g_sProgram As String 'program name
Public g_sVersion As String 'current version (changes with each revision)
Public g_sTeam As String    'team 342
Public g_nTimeLeft As Integer
Public g_sCode() As String 'this is C source code
Public g_sVM() As String 'this is virtual machine (VM) code

'Public g_sCode(MAX_CODE_LINES) As String  'stores program
Public g_nMaxLines As Integer

'loads global variables
Public Sub LoadVariables()
  g_sProgram = "F.I.R.S.T. RC Virtual Machine"
  g_sVersion = "Version 0.07"
  g_sTeam = "Team 342"
End Sub

