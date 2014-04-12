VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBooleanEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Processes boolean expressions or a collection of boolean expressions and sets outputs accordingly."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
' CLASS CBooleanExpression - Written by Chuck Bolin, September 27, 2004
' Purpose:
' Reads boolean expressions and solves problems.   Expressions consist of
' various inputs and bit markers logically expressed. Outputs are turned
' on and off based upon logical rules.
'********************************************************************************
Option Explicit
Option Base 1

'class constants
Private Const MAX_INPUT_BITS = 16
Private Const MAX_MARKER_BITS = 32
Private Const MAX_OUTPUT_BITS = 16

'class types
Private Type BIT_ARRAY
  Bit As Boolean
  Symbol As String
  Comment As String
End Type

'class arrays - first position is 1, not 0
Private m_nIn(MAX_INPUT_BITS) As BIT_ARRAY
Private m_nMarker(MAX_MARKER_BITS) As BIT_ARRAY
Private m_nOut(MAX_OUTPUT_BITS) As BIT_ARRAY

'class private variable declarations...initialized in Class_Initialize
Private m_nErrorCode As Integer
Private m_sProgramString As String 'stores entire program in one string
Private m_sProgramArray() As String 'empty array that stores parsed program (vbcrlf) delimited

'initialize variables
Private Sub Class_Initialize()
  Dim i As Integer
  
  'initialize variables
  m_nErrorCode = 0
  m_sProgramString = ""
  
  'initialize arrays
  For i = 1 To MAX_INPUT_BITS
    m_nIn(i).Bit = False
    m_nIn(1).Symbol = "INP" & CStr(i)
  Next i
  
  For i = 1 To MAX_OUTPUT_BITS
    m_nOut(i).Bit = False
    m_nOut(1).Symbol = "OUT" & CStr(i)
  Next i
  
  For i = 1 To MAX_MARKER_BITS
    m_nMarker(i).Bit = False
    m_nMarker(1).Symbol = "BIT" & CStr(i)
  Next i
  
  'set for testing
  m_nIn(1).Bit = True
  m_nIn(7).Bit = True
  m_nOut(2).Bit = True
  m_nOut(12).Bit = True
  
End Sub


Public Function GetErrorDescription() As String
End Function

Public Function LoadProgram(sProgram As String) As Integer
End Function

Public Function GetOutput(nBit As Integer) As Boolean
  If nBit < 1 Or nBit > MAX_OUTPUT_BITS Then Exit Function
  GetOutput = m_nOut(nBit).Bit
End Function

Public Function GetInput(nBit As Integer) As Boolean
  If nBit < 1 Or nBit > MAX_INPUT_BITS Then Exit Function
  GetInput = m_nIn(nBit).Bit
End Function

Public Function GetMaxInputBits() As Integer
  GetMaxInputBits = MAX_INPUT_BITS
End Function

Public Function GetMaxoutputBits() As Integer
  GetMaxoutputBits = MAX_OUTPUT_BITS
End Function

Public Sub SetInput(nBit As Integer)
End Sub


