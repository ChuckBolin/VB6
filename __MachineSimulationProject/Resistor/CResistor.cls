VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResistor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************
' Resistor Class - Written by Chuck Bolin, October 2004

'***********************************************************************
Option Explicit

'enumerations and types
Public Enum RESISTOR_FAULT
  RESISTOR_OK = 0
  RESISTOR_OPEN = 1
  RESISTOR_SHORT = 2
End Enum

Private Type Resistor
  Resistor As Single
  Value As Single
  Fault As RESISTOR_FAULT
End Type

'private variables
Private m_nRes() As Resistor
Private m_nTotalRes As Integer
Private m_sResNetwork As String 'stores boolean expression of resistor network
Private m_bValidNetwork As Boolean 'true if network is a valid string
Private m_sFilter(9) As String 'stores filter strings that are legal
Private m_nTotalResistance As Single
'initial setup of class
Private Sub Class_Initialize()
  ReDim m_nRes(0)
  m_nTotalRes = 0
  m_bValidNetwork = False
  
  'load filters
  m_sFilter(0) = "*&*"
  m_sFilter(1) = "*+*"
  m_sFilter(2) = "*&(*"
  m_sFilter(3) = "*)"
  m_sFilter(4) = "(*)"
  m_sFilter(5) = "+*)"
  m_sFilter(6) = "&*)"
  m_sFilter(7) = "+*"
  m_sFilter(8) = "&*"
  m_sFilter(9) = "*&)"
  
End Sub

'********************************************* SetResistorNetwork
'user enters boolean string to evaluate
'i.e.  "*&*",  "*+*", "*&(*+*)", etc.
Public Function SetResistorNetwork(sIn As String) As Boolean
  SetResistorNetwork = False
  Dim sTemp As String
  Dim i, nCt As Integer
  
  'remove spaces and count asterisks (*)
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 1) <> " " Then
      sTemp = sTemp & Mid(sIn, i, 1)
      If Mid(sIn, i, 1) = "*" Then nCt = nCt + 1
    End If
  Next i
    
  If Len(sTemp) < 1 Then Exit Function 'nothing in string
  If nCt < 1 Then Exit Function 'no asterisks
  
  'do validity test here<<<<<<<<<<<<<<<<<<<
  m_bValidNetwork = True
  m_sResNetwork = sTemp
  m_nTotalRes = nCt  'number of resistors
  SetResistorNetwork = True
  
  'create resistors to go into network
  ReDim m_nRes(nCt)
  For i = 1 To nCt
    m_nRes(i).Resistor = 100
    m_nRes(i).Fault = RESISTOR_OK
    m_nRes(i).Value = m_nRes(i).Resistor * m_nRes(i).Fault
  Next i
  
End Function

'********************************************* GetResistorNetwork
Public Function GetResistorNetwork() As String
 GetResistorNetwork = m_sResNetwork
End Function

Public Function GetTotalResistors() As Integer
  GetTotalResistors = m_nTotalRes
End Function

Public Function GetTotalResistance() As Single
  CalcTotalResistance
  GetTotalResistance = m_nTotalResistance
End Function

Private Function CalcTotalResistance() As Single
  Dim sNet As String
  Dim i, nIndex As Integer
  
  'replaces * with a number
  nIndex = 0
  For i = 1 To Len(m_sResNetwork)
  
    If Mid(m_sResNetwork, i, 1) = "*" Then
      nIndex = nIndex + 1
      sNet = sNet & CStr(nIndex)
    Else
      sNet = sNet & Mid(m_sResNetwork, i, 1)
    End If
  Next i
  
  MsgBox sNet
  
  m_nTotalResistance = 0
End Function

'********************************************* EvaluateExpression
'Ensures m_sResNetwork if logically correct
'using m_sFilter( )
Private Function EvaluateExpression() As Boolean
  'EvaluateExpression = False
  Dim i, j As Integer
  Dim bFound As Boolean
  
  i = 1
  '*&(*&*+*)&*&(*&*+*+*)&*
  Do While i < Len(m_sResNetwork)
    bFound = False
    For j = 0 To UBound(m_sFilter)
      If Mid(m_sResNetwork, i, Len(m_sFilter(j))) = m_sFilter(j) Then
        bFound = True
        'MsgBox m_sFilter(j) & "    " & Mid(m_sResNetwork, i, Len(m_sFilter(j))) & "    " & m_sResNetwork
        i = i + Len(m_sFilter(j))
        'MsgBox i
        Exit For
      End If
    Next j
    If bFound = False Then
      'MsgBox i & "  " & Len(m_sResNetwork) & " is not logically correct!"
      EvaluateExpression = False
      Exit Function
    End If
  Loop
     
  EvaluateExpression = True
End Function
