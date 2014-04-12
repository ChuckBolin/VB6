Attribute VB_Name = "MathExpression"
'************************************************************************
' MathExpression.bas  - Written by Chuck Bolin
' Date: April 28,2005
' Purpose:  Allows for creation of math variables,
' and solves complex expressions
' Functions and Subs:
'
' Public Sub InitializeMathExpression()
' Private Sub ClearVariableArray()
' Private Function AddConstant(sInput As String) As String
' Private Function AddVariable(sInput As String) As String
' Private Function SetVariableValue(sSymbol As String, sValue As String)
' Private Function GetValue(sSymbol As String) As StringAs String
' Private Function CreatePattern(sInput As String) As String
'************************************************************************
Option Explicit

Public Type MATH_VARIABLES
  Symbol As String
  VariableName As String
  VariableValue As String
  VariableType As String
  VariableScope As String
End Type

'used only within this module
Private g_uVar(26) As MATH_VARIABLES

'******************************************
'must call this to initialize this module
Public Sub InitializeMathExpression()
  Dim i As Integer
  
  'loads symbols A through Z...used in program
  For i = 65 To 90
    g_uVar(i - 64).Symbol = Chr(i)
  Next i
End Sub

'**********************************************
' returns answer or ERROR
Public Function ProcessMathExpression(sInput As String) As String
  Dim i As Integer
  Dim sTemp As String
  Dim sNum As String
  Dim sReturn As String
  Dim bFound As Boolean
  Dim nPos As Integer
  Dim sVarA As String
  Dim sVarB As String
  
  ProcessMathExpression = "ERROR"

  For i = 1 To Len(sInput)
    If IsNumeric(Mid(sInput, i, 1)) Or Mid(sInput, i, 1) = "." Then
      sNum = sNum & Mid(sInput, i, 1)
    ElseIf Mid(sInput, i, 1) = " " Then
  
    Else
      If Len(sNum) > 0 Then
        sReturn = AddConstant(sNum)
        If Left(sReturn, 5) = "ERROR" Then Exit Function
        sTemp = sTemp & sReturn
        sNum = ""
      End If
      sTemp = sTemp & Mid(sInput, i, 1)
    End If
  
  Next i

  'grab last number if expression ends with a number
  If Len(sNum) > 0 Then
    sReturn = AddConstant(sNum)

    If Left(sReturn, 5) = "ERROR" Then Exit Function
    sTemp = sTemp & sReturn
    sNum = ""
  End If
  sInput = sTemp
  
  'now expression has replaced all numbers with letters
  'must now pattern match and reduce this down to
  'a single letter
  sTemp = ""
  While Len(sInput) > 1
    bFound = False
    sTemp = CreatePattern(sInput) 'this is pattern matching string
    
    'MsgBox sInput
    
    'check each possible match..in order of operations
    nPos = InStr(1, sTemp, ")(")
    If nPos > 0 And bFound = False Then
      bFound = True
      sInput = Left(sInput, nPos) & "*" & Mid(sInput, nPos + 1)
    End If
   
    nPos = InStr(1, sTemp, "(%)")
    If nPos > 0 And bFound = False Then
      bFound = True
      sInput = Left(sInput, nPos - 1) & Mid(sInput, nPos + 1, 1) & Mid(sInput, nPos + 3)
    End If
    
    nPos = InStr(1, sTemp, "(%*%)")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos + 1, 1))
      sVarB = GetValue(Mid(sInput, nPos + 3, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos + 1, 1), CStr(Val(sVarA) * Val(sVarB)))
      sInput = Left(sInput, nPos + 1) & Mid(sInput, nPos + 4)
    End If
    
    nPos = InStr(1, sTemp, "(%/%)")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos + 1, 1))
      sVarB = GetValue(Mid(sInput, nPos + 3, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos + 1, 1), CStr(Val(sVarA) / Val(sVarB)))
      sInput = Left(sInput, nPos + 1) & Mid(sInput, nPos + 4)
    End If
      
    nPos = InStr(1, sTemp, "(%+%)")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos + 1, 1))
      sVarB = GetValue(Mid(sInput, nPos + 3, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos + 1, 1), CStr(Val(sVarA) + Val(sVarB)))
      sInput = Left(sInput, nPos + 1) & Mid(sInput, nPos + 4)
    End If
    
    nPos = InStr(1, sTemp, "(%-%)")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos + 1, 1))
      sVarB = GetValue(Mid(sInput, nPos + 3, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos + 1, 1), CStr(Val(sVarA) - Val(sVarB)))
      sInput = Left(sInput, nPos + 1) & Mid(sInput, nPos + 4)
    End If
    
    nPos = InStr(1, sTemp, "%*%")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos, 1))
      sVarB = GetValue(Mid(sInput, nPos + 2, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos, 1), CStr(Val(sVarA) * Val(sVarB)))
      sInput = Left(sInput, nPos) & Mid(sInput, nPos + 3)
    End If
  
    nPos = InStr(1, sTemp, "%/%")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos, 1))
      sVarB = GetValue(Mid(sInput, nPos + 2, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos, 1), CStr(Val(sVarA) / Val(sVarB)))
      sInput = Left(sInput, nPos) & Mid(sInput, nPos + 3)
    End If
  
    nPos = InStr(1, sTemp, "%+%")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos, 1))
      sVarB = GetValue(Mid(sInput, nPos + 2, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos, 1), CStr(Val(sVarA) + Val(sVarB)))
      sInput = Left(sInput, nPos) & Mid(sInput, nPos + 3)
    End If
  
    nPos = InStr(1, sTemp, "%-%")
    If nPos > 0 And bFound = False Then
      bFound = True
      sVarA = GetValue(Mid(sInput, nPos, 1))
      sVarB = GetValue(Mid(sInput, nPos + 2, 1))
      sReturn = SetVariableValue(Mid(sInput, nPos, 1), CStr(Val(sVarA) - Val(sVarB)))
      sInput = Left(sInput, nPos) & Mid(sInput, nPos + 3)
    End If
  
  Wend
  
  ProcessMathExpression = GetValue(sInput)
  
End Function


'******************************************
'clears g_uVar( ) array
Private Sub ClearVariableArray()
  Dim i As Integer
  
  'clears all values
  For i = 65 To 90
    g_uVar(i - 64).Symbol = Chr(i)
    g_uVar(i - 64).VariableName = ""
    g_uVar(i - 64).VariableValue = ""
    g_uVar(i - 64).VariableType = ""
    g_uVar(i - 64).VariableScope = ""
  Next i
  
End Sub

'******************************************
' adds a numerical value, no associated
' variable in this expression
' i.e. 3 * 6
' Returns symbol associated with constant
' or an ERROR
Private Function AddConstant(sInput As String) As String
  Dim i As Integer
  
  'default...assume the worst
  AddConstant = "ERROR"
  
  'reasons to bail
  If Len(sInput) < 1 Then Exit Function
  If Not IsNumeric(sInput) Then Exit Function
  
  'go through array...look for empty value
  For i = 1 To 26
    If g_uVar(i).VariableName = "" Then 'no variable name assigned
      If g_uVar(i).VariableValue = "" Then 'no constant assigned either
        g_uVar(i).VariableValue = sInput
        AddConstant = g_uVar(i).Symbol
        Exit Function
      End If
    End If
  Next i
  
End Function

'******************************************
' adds a variable name
' i.e. nNumber * nInterest
' Returns symbol associated with variable
' or an ERROR
Private Function AddVariable(sInput As String) As String
  Dim i As Integer
  
  'default...assume the worst
  AddVariable = "ERROR"
  
  'reasons to bail
  If Len(sInput) < 1 Then Exit Function
  If IsNumeric(sInput) Then Exit Function
  
  'go through array...look for empty value
  For i = 1 To 26
    If g_uVar(i).VariableName = "" Then 'no variable name assigned
      If g_uVar(i).VariableValue = "" Then 'no constant assigned either
        g_uVar(i).VariableName = sInput
        AddVariable = g_uVar(i).Symbol
        Exit Function
      End If
    End If
  Next i
  
End Function

'******************************************
' Sets a variable to a specific value
' i.e. nNumber * nInterest
' Returns symbol associated with variable
' or an ERROR
Private Function SetVariableValue(sSymbol As String, sValue As String) As String
  Dim i As Integer
  
  'default...assume the worst
  SetVariableValue = "ERROR"
  
  'reasons to bail
  If Len(sSymbol) < 1 Then Exit Function
  If Len(sValue) < 1 Then Exit Function
  If Not IsNumeric(sValue) Then Exit Function
  
  'go through array...look for empty value
  For i = 1 To 26
    If g_uVar(i).Symbol = sSymbol Then '
      g_uVar(i).VariableValue = sValue
      SetVariableValue = g_uVar(i).Symbol
      Exit Function
    End If
  Next i
  
End Function

'******************************************
' Sets a variable to a specific value
' i.e. nNumber * nInterest
' Returns symbol associated with variable
' or an ERROR
Private Function GetValue(sSymbol As String) As String
  Dim i As Integer
  
  'default...assume the worst
  GetValue = "ERROR"
  
  'reasons to bail
  If Len(sSymbol) < 1 Then Exit Function
    
  'go through array...look for empty value
  For i = 1 To 26
    If g_uVar(i).Symbol = sSymbol Then '
      If g_uVar(i).VariableValue = "" Then g_uVar(i).VariableValue = "0"
      GetValue = g_uVar(i).VariableValue
      Exit Function
    End If
  Next i
  
End Function

'**********************************************
' Generates pattern matching expression where
' % represents each letter A through Z
Private Function CreatePattern(sInput As String) As String
  Dim i As Integer
  Dim sTemp As String
  
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) >= "A" And Mid(sInput, i, 1) <= "Z" Then
      sTemp = sTemp & "%"
    Else
      sTemp = sTemp & Mid(sInput, i, 1)
    End If
  Next i
  
  CreatePattern = sTemp
End Function
