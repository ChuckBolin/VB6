Attribute VB_Name = "Module1"
Option Explicit

Public Enum SEQ_DIAG
  SD_NOTHING = 0
  SD_ADD_NODE = 1
End Enum

Public Enum ENTITY_TYPE
  Chain = 0
  Node
  PLCInput
  PLCInputNot
  PLCOutput
  PLCOutputNot
  PLCBit
  PLCBitNot
  TiePoint
End Enum

Public Type ENTITY_INFO
  ChainNumber As Integer
  NodeNumber As Integer
  EntityType As ENTITY_TYPE
  Name As String
End Type

Public Type NODE_INFO
  x As Integer
  y As Integer
  Width As Integer
  Height As Integer
  Name As String
  PLCInput As String 'comma separated list of input symbols
  PLCInputNot As String 'comma separated list of input NOT symbols
  PLCOutput As String 'comma separated list of output symbols
  PLCOutputNot As String 'comma separated list of output NOT symbols
End Type

'global variables
Public g_eMode As SEQ_DIAG 'sets drawing mode when tool button is clicked
Public g_nNodeMax As Integer
Public g_uNode(10) As NODE_INFO
Public g_uEntity(100) As ENTITY_INFO
Public g_nNodeCount As Integer
  
'load variables
Public Sub LoadVariables()
  g_eMode = SD_NOTHING
  g_nNodeCount = 0
End Sub
