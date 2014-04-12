Attribute VB_Name = "Module1"
Option Explicit

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

Public g_nNodeMax As Integer
Public g_uNode(10) As NODE_INFO
Public g_uEntity(100) As ENTITY_INFO


  'm_nEntIndex = 1
  'g_uEntity(m_nEntIndex).EntityType = Node
  'g_uEntity(m_nEntIndex).ChainNumber = 1
  'g_uEntity(m_nEntIndex).NodeNumber = 1
  'g_uEntity(m_nEntIndex).Name = "1N0"
  
  'm_nEntIndex = 2
  'g_uEntity(m_nEntIndex).EntityType = PLCInput
  'g_uEntity(m_nEntIndex).ChainNumber = 1
  'g_uEntity(m_nEntIndex).NodeNumber = 1
  'g_uEntity(m_nEntIndex).Name = "S1"
