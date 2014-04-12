Attribute VB_Name = "PLC"
'***************************************************
' PLC.BAS
'
' May 12, 2005 - Removed cards..use just bit array
' Solve latch, unlatch and assigning of bits I,O,B
'***************************************************
Option Explicit

'enumerations
Public Enum PLC_BIT_TYPE
  Bit_Input = 0
  Bit_Output = 1
  Bit_Bit = 2 'like a marker bit
End Enum

'types
'this defines a bit used by
Public Type PLC_BIT
  Symbol As String         'text name for this bit
  Value As Boolean         'true or false
  Latch As Boolean         'true if latched, false means not latched
  ModuleNumber As Integer  '1 through 10, or 0 for bits/markers
  Type As PLC_BIT_TYPE     'input, output, bit marker
End Type

'model for a 16-bit word
Public Type PLC_BIT_ARRAY
  Bit(15) As PLC_BIT
  Type As PLC_BIT_TYPE
End Type

Public Type PLC_CODE_SEGMENT
  LValue As String
  Operation As String
  RValue As String
  Valid As Boolean
End Type

'public constants
Public Const PLC_MAX_BITS = 1000

'public variables
Public g_uCard(10) As PLC_BIT_ARRAY 'stores all bits associated with PLC rack
Public g_uBit(PLC_MAX_BITS) As PLC_BIT
Public g_nLastBit As Integer 'stores last used bit in array
Public g_sOperationList As String 'legal operations
Public g_sLogicList As String 'legal logic operators
Public g_sMathList As String 'legal math operators
Public g_sConditionList As String 'legal condition operators
Public g_sProg() As String 'stores program

'initialize PLC
Public Sub InitializePLC()
    
  'loads operation list
  g_sOperationList = "\RT,\CU,\CD,\RC,\=,\L,\U,\T"
  g_sLogicList = "&|!"
  g_sMathList = "*/+-"
  g_sConditionList = "<>=="
  
  'array
  ReDim g_sProg(1)
  g_nLastBit = 0
End Sub

'sets or latches a bit
Public Function LatchBit(nBit As Integer) As Boolean
  LatchBit = False
  
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
  g_uBit(nBit).Latch = True
  g_uBit(nBit).Value = True
  
  LatchBit = True
End Function

'resets or unlatches a bit
Public Function UnlatchBit(nBit As Integer) As Boolean
  UnlatchBit = False
    
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
    
  g_uBit(nBit).Latch = False
  g_uBit(nBit).Value = False
  
  UnlatchBit = True
End Function

'assigns bit as true
Public Function AssignBit(nBit As Integer) As Boolean
  AssignBit = False
    
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  If g_uBit(nBit).Latch = True Then
    'do nothing
  Else
    g_uBit(nBit).Value = True
  End If
    
  AssignBit = True
End Function

'unassigns bit as true
Public Function UnAssignBit(nBit As Integer) As Boolean
  UnAssignBit = False
    
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
  
  'only unassign a bit if it is not latched
  If g_uBit(nBit).Latch = False Then
    g_uBit(nBit).Value = False
  End If
  
  UnAssignBit = True
End Function

'sets a symbolic name for a particular bit
Public Function SetBitSymbolName(nBit As Integer, sSymbol As String) As Boolean
  SetBitSymbolName = False
  
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  g_uBit(nBit).Symbol = sSymbol
  
  SetBitSymbolName = True
End Function

'clears a symbol name for a particular bit
Public Function ClearBitSymbolName(nBit As Integer) As Boolean
  ClearBitSymbolName = False
  
  If nBit < 0 Or nBit > PLC_MAX_BITS Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  g_uBit(nBit).Symbol = ""
  
  ClearBitSymbolName = True
End Function

'returns a line of code broken into three segments
Public Function ExtractLineOfCode(sInput As String) As PLC_CODE_SEGMENT
  Dim nBS As Integer 'position of backslash
  Dim nSpace As Integer 'position of space after backslash
  Dim sOper As String 'stores operations
  Dim sLValue As String
  Dim sRValue As String
  Dim sTemp As String
  Dim i As Integer
  
  'grabs operations i.e. \=
  nBS = InStr(1, sInput, "\")
  If nBS < 1 Then Exit Function
  nSpace = InStr(nBS, sInput, " ")
  If nSpace < 1 Then Exit Function
  
  'extracts three segments
  sLValue = Left(sInput, nBS - 1)
  sOper = Mid(sInput, nBS, nSpace - nBS)
  sRValue = Mid(sInput, nSpace + 1)
  
  'removes spaces from all three segments and checks length
  'does not validate left side of expression
  sTemp = ""
  For i = 1 To Len(sLValue)
    If Mid(sLValue, i, 1) <> " " Then sTemp = sTemp & Mid(sLValue, i, 1)
  Next i
  sLValue = sTemp
  If Len(sLValue) < 1 Then Exit Function
  
  'does not validate operation
  sTemp = ""
  For i = 1 To Len(sOper)
    If Mid(sOper, i, 1) <> " " Then sTemp = sTemp & Mid(sOper, i, 1)
  Next i
  sOper = sTemp
  If Len(sOper) < 1 Then Exit Function
  
  'does not validate right side of expression
  sTemp = ""
  For i = 1 To Len(sRValue)
    If Mid(sRValue, i, 1) <> " " Then sTemp = sTemp & Mid(sRValue, i, 1)
  Next i
  sRValue = sTemp
  If Len(sRValue) < 1 Then Exit Function
  
  'return information
  ExtractLineOfCode.LValue = sLValue
  ExtractLineOfCode.Operation = sOper
  ExtractLineOfCode.RValue = sRValue
  ExtractLineOfCode.Valid = True
End Function

'returns true if operation is legal operation
Public Function IsOperationExpression(sInput As String) As Boolean
  Dim i As Integer
  Dim nCt As Integer
  Dim sOper() As String
  Dim bFound As Boolean
  
  IsOperationExpression = False
  sOper = Split(g_sOperationList, ",")
  bFound = False
  
  For i = 0 To UBound(sOper)
    If sOper(i) = sInput Then
      bFound = True
      Exit For
    End If
  Next i
  If bFound = False Then Exit Function
  
  IsOperationExpression = True
End Function

'returns true if a purely logic expression
Public Function IsLogicExpression(sInput As String) As Boolean
  Dim i As Integer
  Dim nCt As Integer 'counter
  
  IsLogicExpression = False
  
  'look for at least one logic operator
  For i = 1 To Len(sInput)
    If InStr(i, g_sLogicList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
  Next i
  If nCt < 1 Then Exit Function 'false..not a logic expression
  
  'verify there are no math operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sMathList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
  
  'verify there are no conditional operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sConditionList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
    
  IsLogicExpression = True
End Function

'returns true if a purely math expression
Public Function IsMathExpression(sInput As String) As Boolean
  Dim i As Integer
  Dim nCt As Integer 'counter
  
  IsMathExpression = False
  
  'look for at least one math operator
  For i = 1 To Len(sInput)
    If InStr(i, g_sMathList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
  Next i
  If nCt < 1 Then Exit Function 'false..not a math expression
  
  'verify there are no logic operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sLogicList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
  
  'verify there are no conditional operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sConditionList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
    
  IsMathExpression = True
End Function

'returns true if a purely math expression
Public Function IsConditionExpression(sInput As String) As Boolean
  Dim i As Integer
  Dim nCt As Integer 'counter
  
  IsConditionExpression = False
  
  'look for at least one condition operator
  For i = 1 To Len(sInput)
    If InStr(i, g_sConditionList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
  Next i
  If nCt < 1 Then Exit Function 'false..not a condition expression
  
  'verify there are no logic operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sLogicList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
  
  'verify there are no math operators
  For i = 1 To Len(sInput)
    If InStr(i, g_sMathList, Mid(sInput, i, 1)) > 0 Then Exit Function
  Next i
    
  IsConditionExpression = True
End Function

'reads text file into array
Public Function LoadProgram(sInput As String) As Boolean
  Dim i As Integer
  Dim sCode() As String
  
  ReDim g_sProg(1)
  
  LoadProgram = False
  sCode = Split(sInput, vbCrLf)
  
  'loads text file into array
  For i = 0 To UBound(sCode) - 1
    sCode(i) = Trim(sCode(i))
    If Len(sCode(i)) > 0 Then
      g_sProg(UBound(g_sProg)) = sCode(i)
      ReDim Preserve g_sProg(UBound(g_sProg) + 1)
    End If
  Next i
        
  LoadProgram = True
End Function

'evaluates program
Public Function EvaluateProgram() As Boolean
  Dim i As Integer
  Dim uRet As PLC_CODE_SEGMENT
  Dim nTypeLValue As Integer '1 = numeric, 2=logic, 3=condition, 4=math
    
  EvaluateProgram = False
  
  'parse code
  For i = 1 To UBound(g_sProg) - 1
    
    'grab one line of code at a time
    uRet = ExtractLineOfCode(g_sProg(i))
    
    'verify operations is legal: i.e. \L, \=, \U
    If IsOperationExpression(uRet.Operation) = False Then
      MsgBox "Illegal operation! " & uRet.Operation
      Exit Function
    End If
    
    'determine left side of expression type
    nTypeValue = 0
    If IsNumeric(uRet.LValue) = True Then
      nTypeLValue = 1
    ElseIf IsLogicExpression(uRet.LValue) = True Then
      nTypeLValue = 2
    ElseIf IsConditionExpression(uRet.LValue) = True Then
      nTypeLValue = 3
    ElseIf IsMathExpression(uRet.LValue) = True Then
      nTypeLValue = 4
    End If
    
    If nTypeLValue < 1 Then
      MsgBox "Left side of expression: " & uRet.LValue & " is illegal!"
      Exit Function
    End If
    
    
    MsgBox uRet.LValue & "    " & uRet.Operation & "    " & uRet.RValue & "    " & uRet.Valid
  Next i
  
  EvaluateProgram = True
End Function
