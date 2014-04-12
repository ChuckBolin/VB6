Attribute VB_Name = "PLC"
'****************************************
' PLC.BAS
' Fix IsOperationExpression OK with \A but not \AZ
'****************************************
Option Explicit

'enumerations
Public Enum PLC_BIT_TYPE
  Bit_Input = 0
  Bit_Output = 1
  Bit_Bit = 2
End Enum

'types
Public Type PLC_BIT
  Symbol As String
  Value As Boolean
  Latch As Boolean
End Type

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

'public variables
Public g_uCard(10) As PLC_BIT_ARRAY 'stores all bits associated with PLC rack
Public g_sOperationList As String 'legal operations
Public g_sLogicList As String 'legal logic operators
Public g_sMathList As String 'legal math operators
Public g_sConditionList As String 'legal condition operators
Public g_sProg() As String 'stores program

'initialize PLC
Public Sub InitializePLC()
  
  'configures cards in PLC as I or O
  g_uCard(1).Type = Bit_Input  'input modules
  g_uCard(2).Type = Bit_Input
  g_uCard(3).Type = Bit_Input
  g_uCard(4).Type = Bit_Input
  g_uCard(5).Type = Bit_Input
  g_uCard(6).Type = Bit_Output 'output modules
  g_uCard(7).Type = Bit_Output
  g_uCard(8).Type = Bit_Output
  g_uCard(9).Type = Bit_Output
  g_uCard(10).Type = Bit_Output
  
  'loads operation list
  g_sOperationList = "\=,\L,\U,\T,\RT,\CU,\CD,\RC"
  g_sLogicList = "&|!"
  g_sMathList = "*/+-"
  g_sConditionList = "<>=="
  
  'array
  ReDim g_sProg(1)
End Sub

'sets or latches a bit
Public Function LatchCardBit(nCard As Integer, nBit As Integer) As Boolean
  LatchCardBit = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  If g_uCard(nCard).Type = Bit_Input Then Exit Function 'can't set inputs
  
  g_uCard(nCard).Bit(nBit).Latch = True
  g_uCard(nCard).Bit(nBit).Value = True
  
  LatchCardBit = True
End Function

'resets or unlatches a bit
Public Function UnlatchCardBit(nCard As Integer, nBit As Integer) As Boolean
  UnlatchCardBit = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  If g_uCard(nCard).Type = Bit_Input Then Exit Function 'can't reset inputs
  
  g_uCard(nCard).Bit(nBit).Latch = False
  g_uCard(nCard).Bit(nBit).Value = False
  
  UnlatchCardBit = True
End Function

'assigns bit as true
Public Function AssignCardBit(nCard As Integer, nBit As Integer) As Boolean
  AssignCardBit = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  
  g_uCard(nCard).Bit(nBit).Value = True
  
  AssignCardBit = True
End Function

'unassigns bit as true
Public Function UnAssignCardBit(nCard As Integer, nBit As Integer) As Boolean
  UnAssignCardBit = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be unassigned if equal to a sensor or switch
  
  'only unassign a bit if it is not latched
  If g_uCard(nCard).Bit(nBit).Latch = False Then
    g_uCard(nCard).Bit(nBit).Value = False
  End If
  
  UnAssignCardBit = True
End Function

'sets a symbolic name for a particular bit
Public Function SetCardBitSymbolName(nCard As Integer, nBit As Integer, sSymbol As String) As Boolean
  SetCardBitSymbolName = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  g_uCard(nCard).Bit(nBit).Symbol = sSymbol
  
  SetCardBitSymbolName = True
End Function

'clears a symbol name for a particular bit
Public Function ClearCardBitSymbolName(nCard As Integer, nBit As Integer) As Boolean
  ClearCardBitSymbolName = False
  
  If nCard < 1 Or nCard > 10 Then Exit Function  'need legal card 1 - 10
  If nBit < 0 Or nBit > 15 Then Exit Function    'need legal bit 0 - 15
  
  'inputs can be assigned if equal to a sensor or switch
  g_uCard(nCard).Bit(nBit).Symbol = ""
  
  ClearCardBitSymbolName = True
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
  sTemp = ""
  For i = 1 To Len(sLValue)
    If Mid(sLValue, i, 1) <> " " Then sTemp = sTemp & Mid(sLValue, i, 1)
  Next i
  sLValue = sTemp
  If Len(sLValue) < 1 Then Exit Function
  
  sTemp = ""
  For i = 1 To Len(sOper)
    If Mid(sOper, i, 1) <> " " Then sTemp = sTemp & Mid(sOper, i, 1)
  Next i
  sOper = sTemp
  If Len(sOper) < 1 Then Exit Function
  
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
  
  IsOperationExpression = False

  For i = 1 To Len(sInput)
    If InStr(1, g_sOperationList, Mid(sInput, i, 3)) > 0 Then
      nCt = nCt + 1
    ElseIf InStr(1, g_sOperationList, Mid(sInput, i, 2)) > 0 Then
      nCt = nCt + 1
    End If
  Next i
  If nCt < 1 Then Exit Function
  
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
  
  For i = 1 To UBound(g_sProg) - 1
    uRet = ExtractLineOfCode(g_sProg(i))
    If IsOperationExpression(uRet.Operation) = False Then
      MsgBox "Illegal operation! " & uRet.Operation
      'Exit Function
    End If
    'MsgBox uRet.LValue & "    " & uRet.Operation & "    " & uRet.RValue & "    " & uRet.Valid
  Next i
  
  EvaluateProgram = True
End Function
