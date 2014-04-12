Attribute VB_Name = "Global"
Option Explicit

'stores class information for name and header remarks
Public Type CLASS_HEADER
  ClassName As String
  Author As String
  Date As String
  Purpose As String
End Type

'defines name of variable and type
Public Type CLASS_VAR_TYPE
  DataType As Integer  '0=integer, 1=single, 2=double,3=boolean,
                      '4=string, 5=byte
  Name As String 'name of variable
End Type

'an object array stores all info about properties
'and procedures so user can delete items
Public Type CLASS_CONTENT
  PropProc As Integer '0=property, 1=sub, 2=function
  Scope As Integer '0=Private, 1=Public
  DataType As Integer '0=integer, 1=single, 2=double,3=boolean,
                      '4=string, 5=byte
  Name As String  'name of property, sub or function
  ReadValue As Boolean 'true to read
  WriteValue As Boolean 'true to write
  InitialValue As String
  MinimumValue As String '
  MaximumValue As String
  Abstract As Boolean 'true if abstract
  Comment As String 'allows user comment
  NumOfArgs As Integer '0=none, number of args for function
  Arg1 As CLASS_VAR_TYPE 'stores data for four arguments
  Arg2 As CLASS_VAR_TYPE
  Arg3 As CLASS_VAR_TYPE
  Arg4 As CLASS_VAR_TYPE
  Delete As Boolean 'false means do not delete
End Type

'global variables
Public g_uClassHeader As CLASS_HEADER
Public g_uClassContent() As CLASS_CONTENT
Public g_nIndex As Integer 'pointer to g_uClassContent
Public g_sQuote As String

'global constants
Public Const G_INTEGER = 0
Public Const G_SINGLE = 1
Public Const G_DOUBLE = 2
Public Const G_BOOLEAN = 3
Public Const G_STRING = 4
Public Const G_BYTE = 5
Public Const G_PRIVATE = 0
Public Const G_PUBLIC = 1
Public Const G_PROPERTY = 0
Public Const G_SUB = 1
Public Const G_FUNCTION = 2




