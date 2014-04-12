Attribute VB_Name = "PLC"
'***************************************************
' PLC.BAS
'
' May 14, 2005 -  Work on assigning a bit, latching
' unlatching based upon logic expression
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
  Absolute As String       'construct of type, module and bit
  Value As Boolean         'true or false
  Latch As Boolean         'true if latched, false means not latched
  Word As Integer          '1 through 10, or 0 for bits/markers
  Bit As Integer           '0 through 15
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
'Public g_uCard(10) As PLC_BIT_ARRAY 'stores all bits associated with PLC rack
Public g_uBit(PLC_MAX_BITS) As PLC_BIT
Public g_nLastBit As Integer 'stores last used bit in array
Public g_sOperationList As String 'legal operations
Public g_sLogicList As String 'legal logic operators
Public g_sMathList As String 'legal math operators
Public g_sConditionList As String 'legal condition operators
Public g_sProg() As String 'stores program

'initialize PLC
'Input modules 1.0 - 1.15, 2.0 - 2.15
'Output modules 3.0 - 3.15, 4.0 - 4.15
'Bit markers 0.0 - 0.15,...5.0 - 5.15
Public Sub InitializePLC()
  Dim i As Integer
  Dim j As Integer
  
  'loads operation list
  g_sOperationList = "\RT,\CU,\CD,\RC,\=,\L,\U,\T"
  g_sLogicList = "&|!"
  g_sMathList = "*/+-"
  g_sConditionList = "<>=="
  
  'add 2 input modules - cards 1 and 2
  For j = 1 To 2
    For i = 0 To 15
      g_uBit((j - 1) * 16 + i).Type = Bit_Input
      g_uBit((j - 1) * 16 + i).Absolute = "I" & CStr(j) & "." & CStr(i)
    Next i
  Next j
  
  'add 2 output modules - cards 3 and 4
  For j = 3 To 4
    For i = 0 To 15
      g_uBit((j - 1) * 16 + i).Type = Bit_Output
      g_uBit((j - 1) * 16 + i).Absolute = "O" & CStr(j) & "." & CStr(i)
    Next i
  Next j
  
  'add 6 bit markers
  For j = 5 To 10
    For i = 0 To 15
      g_uBit((j - 1) * 16 + i).Type = Bit_Bit
      g_uBit((j - 1) * 16 + i).Absolute = "B" & CStr(j - 5) & "." & CStr(i)
    Next i
  Next j
  
  g_nLastBit = 160 'there are 160 addressable bits
    
  'array
  ReDim g_sProg(1)
  
End Sub

'reads value of a bit
Public Function ReadBitValue(nBit As Integer) As Boolean
  If nBit < 0 Then Exit Function
  If nBit > g_nLastBit Then Exit Function

  ReadBitValue = g_uBit(nBit).Value
End Function

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
    If InStr(1, g_sLogicList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
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
    If InStr(1, g_sMathList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
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
    If InStr(1, g_sConditionList, Mid(sInput, i, 1)) > 0 Then nCt = nCt + 1
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
'************************************************************************
'                                          E V A L U A T E  P R O G R A M
'************************************************************************
Public Function EvaluateProgram() As Boolean
  Dim i As Integer
  Dim uRet As PLC_CODE_SEGMENT
  Dim nTypeLValue As Integer '1 = numeric, 2=logic, 3=condition, 4=math
  Dim nRet As Integer
  
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
    nTypeLValue = 0
    
    If IsNumeric(uRet.LValue) = True Then
      nTypeLValue = 1
    ElseIf IsLogicExpression(uRet.LValue) = True Then
      nTypeLValue = 2
    ElseIf IsConditionExpression(uRet.LValue) = True Then
      nTypeLValue = 3
    ElseIf IsMathExpression(uRet.LValue) = True Then
      nTypeLValue = 4
    End If
    
    'process expression based upon type
    If nTypeLValue < 1 Then      'unknown expression
      MsgBox "Left side of expression: " & uRet.LValue & " is illegal!"
      Exit Function
    ElseIf nTypeLValue = 2 Then  'logic expression
      nRet = GetLogicResult(uRet.LValue) 'holds 1 or 0 or -1 error
      If nRet < 0 Then
        MsgBox "Illegal logic expression: " & uRet.LValue
        Exit Function
      End If
      MsgBox uRet.LValue & " = " & nRet  '<<<<<<<<<<<<<<<<<<
    End If
    
  
    'MsgBox uRet.LValue & "    " & uRet.Operation & "    " & uRet.RValue & "    " & uRet.Valid
  Next i
  
  EvaluateProgram = True
End Function
'********************************************************************

'reads the bit array g_uBit( ) looking for a match for absolute or symbolic
Private Function GetBitNumber(sInput As String) As Integer
  Dim i As Integer
  
  GetBitNumber = -1 'default..not found
  
  'go through table and look for match, return index of array if found
  For i = 0 To g_nLastBit
    If UCase(sInput) = UCase(g_uBit(i).Symbol) Or UCase(sInput) = UCase(g_uBit(i).Absolute) Then
      GetBitNumber = i
      Exit Function
    End If
  Next i
  
End Function

'evaluates a logic expression and returns a 0, 1 or -1. -1 is an error
' i.e. I1.12 & B3.4
Private Function GetLogicResult(sInput As String) As Integer
  Dim i As Integer
  Dim sToken As String 'holds an absolute/symbolic I/O/B
  Dim sTemp As String 'holds new string of operators and bit values 1 or 0
  Dim nRet As Integer 'grabs return value
  Dim sBit As String
  Dim bFound As Boolean
  Dim nPos As Integer
  
  GetLogicResult = -1 'default ERROR
  
  'convert string of identifiers and operators to a modified
  'string version such as 1&0|!1
  'MsgBox "Input: " & sInput
  For i = 1 To Len(sInput)
  
    'this is a legal logic operator
    If InStr(1, g_sLogicList, Mid(sInput, i, 1)) > 0 Or Mid(sInput, i, 1) = "(" Or Mid(sInput, i, 1) = ")" Then
      If Len(sToken) > 0 Then
        'MsgBox sToken
        nRet = GetBitNumber(sToken)
        If nRet < 0 Then 'illegal identifier I/O/B
          MsgBox "Unknown identifier: " & sToken
          Exit Function
        End If
        'MsgBox "Here" & ReadBitValue(nRet)
        If ReadBitValue(nRet) = True Then
          sBit = "1"
        Else
          sBit = "0"
        End If
        sTemp = sTemp & sBit
        sToken = ""
      End If
      
      sTemp = sTemp & Mid(sInput, i, 1)
    Else
      sToken = sToken & Mid(sInput, i, 1)
    End If
  Next i
  
  'any trailing token
  If Len(sToken) > 0 Then
    nRet = GetBitNumber(sToken)
    If nRet < 0 Then 'illegal identifier I/O/B
      MsgBox "Unknown identifier: " & sToken
      Exit Function
    End If
    'MsgBox ReadBitValue(nRet)
    If ReadBitValue(nRet) = True Then
      sBit = "1"
    Else
      sBit = "0"
    End If
    sTemp = sTemp & sBit
  End If
  
  'now, sTemp holds a modified string of operators and 1 or 0
  
  'must pass through string several times until len is 1 with a
  'value of 0 or 1...else exit because of illegal format
  Do
    bFound = False
    
    nPos = InStr(1, sTemp, "!0")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 2)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "!1")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 2)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, ")(")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & ")&(" & Mid(sTemp, nPos + 2)
      bFound = True
    End If

    nPos = InStr(1, sTemp, "(0)")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 3)
      bFound = True
    End If

    nPos = InStr(1, sTemp, "(1)")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
    
    nPos = InStr(1, sTemp, "0&0")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "0&1")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "1&0")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "1&1")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
    
    nPos = InStr(1, sTemp, "0|0")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "0" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "0|1")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "1|0")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
  
    nPos = InStr(1, sTemp, "1|1")
    If bFound = False And nPos > 0 Then
      sTemp = Left(sTemp, nPos - 1) & "1" & Mid(sTemp, nPos + 3)
      bFound = True
    End If
    
    If bFound = False Then Exit Do 'nothing match..illegal
  
  Loop Until Len(sTemp) = 1
  
  If sTemp = "0" Then
    GetLogicResult = 0
  ElseIf sTemp = "1" Then
    GetLogicResult = 1
  Else
    'error...oops!
  End If
  
  'MsgBox sTemp
  
  
End Function


