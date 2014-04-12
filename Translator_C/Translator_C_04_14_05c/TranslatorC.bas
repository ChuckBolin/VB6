Attribute VB_Name = "TranslatorC"
'****************************************************************************
' Translator.bas - Written by Chuck Bolin Team 342, April 2005
' Contains functions and subs that parse and compile "C" source code
' into Virtual Machine (VM) code.
' Functions:

' Convert_C_VM(sInput) standalones lines converted to VM
' BuildSystemVariables() adds all RC default I/Os to g_uVar() table
' C_DecrementVariable(sIn) or returns ERROR
' C_IncrementVariable(sIn) or returns ERROR
' GetVariableValue(sIn) returns value of variable or ERROR
' CorrectLineTermination(sInput)splices two lines together ending in letter
' TranslateVariables_VM(sInput) finds all variables and code -> VM code
' AddVariable(sName, sType, sValue, Optional sScope) adds variable to g_uVar()
' AlignBraces(sInput) ensures each brace is on it's own line
' IsAutoFunction(sInput) returns string of C code if essential autonomous
'                        lines of code are present.
' RemoveBlankLines(sInput) returns string of C w/o tabs and leading/trailing
'                          white spaces
' RemoveComments(sInput) returns string of C w/o comments
'
'****************************************************************************
Option Explicit

'public constants
'****************
Public Const MAX_VARIABLES = 500

'public types
'************
Public Type VARIABLE_TABLE
  Name As String
  Type As String
  Value As String
  Scope As String
End Type

'public variables
'****************
Public g_sTranslatorC_Version As String

Public g_uVar(MAX_VARIABLES) As VARIABLE_TABLE
Public g_nNumVar As Integer 'number of variables actually in use

'*******************************************
' Convert various C standalones to VM code
' Once the C code has been formatted and
' cleaned up then individual lines may
' be considered.  This routine may have
' to be repeated until all braces have
' been eliminated
'*******************************************
Public Function Convert_C_VM(sInput As String) As String
  Dim sLines() As String
  Dim sOut As String
  Dim i As Integer
  Dim sCode As String
  Dim sReturn As String
    
  sLines = Split(sInput, vbCrLf)
  
  'loads program into an array
  For i = 0 To UBound(sLines) - 1  'fetch one line of code
    sCode = Trim(sLines(i))        'trim line
    If Right(sCode, 1) = ";" Then sCode = Left(sCode, Len(sCode) - 1) 'remove semicolons
    ' MsgBox sIn & "  " & Len(sIn)
    
    If InStr(1, sCode, "==") Then
    
    ElseIf InStr(1, sCode, "++") Then  'post-fix operator ++
      
      sOut = sOut & C_IncrementVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        Convert_C_VM = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    ElseIf InStr(1, sCode, "--") Then  'post-fix operator --
      sOut = sOut & C_DecrementVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        Convert_C_VM = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    ElseIf InStr(1, sCode, "*=") Then
    
    ElseIf InStr(1, sCode, "/=") Then
    
    ElseIf InStr(1, sCode, "+=") Then
    
    ElseIf InStr(1, sCode, "-=") Then
    
    ElseIf InStr(1, sCode, "=") Then
           
    End If
  
  Next i
  
  'construct string from array
  For i = 0 To UBound(sLines) - 1
   sOut = sOut & sLines(i) & vbCrLf
  Next i

  Convert_C_VM = sOut
End Function

Public Function BuildSystemVariables() As String
  Dim sOut As String
    
  'sOut = "ERROR!" & vbCrLf & "   Unknown!"
  
  'adds to variable table
  sOut = sOut & AddVariable("pwm01", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm02", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm03", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm04", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("User_Mode_byte", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Switch1_LED", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Switch2_LED", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Switch3_LED", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Pwm1_red", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Pwm2_red", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Pwm1_green", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Pwm2_green", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Relay1_red", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Relay2_red", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Relay1_green", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("Relay2_green", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay1_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay2_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay3_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay4_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay1_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay2_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay3_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay4_rev", "unsigned char", "0", "static")
  
  'constructs VM Code
  sOut = sOut & "CVAR pwm01,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm02,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm03,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm04,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR Pwm1_red,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Pwm2_red,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Pwm1_green,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Pwm2_green,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Relay1_red,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Relay2_red,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Relay1_green,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Relay2_green,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR relay1_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay1_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay2_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay2_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay3_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay3_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay4_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay4_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR User_Mode_byte,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Switch1_LED,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Switch2_LED,unsigned char,0" & vbCrLf
  sOut = sOut & "CVAR Switch3_LED,unsigned char,0" & vbCrLf
  
  BuildSystemVariables = sOut
End Function


'*******************************************
'decrements variable by 1
'*******************************************
Public Function C_DecrementVariable(sIn) As String
  Dim nPlus As Integer 'position of first + in ++
  Dim sName As String 'var name
  Dim sReturn As String 'catches variable value
  
  C_DecrementVariable = "ERROR!" & vbCrLf & "   Invalid use of '--' operator!"
  
  nPlus = InStr(1, sIn, "--")  'must have --
  If nPlus > 1 Then  '-- found
    sName = Trim(Left(sIn, nPlus - 1))
    sOut = sOut & GetVariableValue(sName)
    If Left(sReturn, 5) = "ERROR" Then  'oops, no variable by this name
      C_DecrementVariable = "ERROR!" & vbCrLf & "   Variable not declared!"
      Exit Function
    Else  'this is OK
      C_DecrementVariable = "DEC " & sName & ",1" & vbCrLf
    End If
  Else  '-- not found
    C_DecrementVariable = "ERROR!" & vbCrLf & "   Missing '--' operator!"
    Exit Function
  End If
End Function

'*******************************************
'increments variable by 1
'*******************************************
Public Function C_IncrementVariable(sIn) As String
  Dim nPlus As Integer 'position of first + in ++
  Dim sName As String 'var name
  Dim sReturn As String 'catches variable value
  
  C_IncrementVariable = "ERROR!" & vbCrLf & "   Invalid use of '++' operator!"
  
  nPlus = InStr(1, sIn, "++")  'must have ++
  If nPlus > 1 Then  '++ found
    sName = Trim(Left(sIn, nPlus - 1))
    sOut = sOut & GetVariableValue(sName)
    If Left(sReturn, 5) = "ERROR" Then  'oops, no variable by this name
      C_IncrementVariable = "ERROR!" & vbCrLf & "   Variable not declared!"
      Exit Function
    Else  'this is OK
     
      If Right(sIn, 2) = "++" Then
        C_IncrementVariable = "INC " & sName & ",1" & vbCrLf
      Else
        C_IncrementVariable = sIn
      End If
    End If
  Else  '++ not found
    C_IncrementVariable = "ERROR!" & vbCrLf & "   Missing '++' operator!"
    Exit Function
  End If
End Function

'*******************************************
'get value assigned to a variable
'*******************************************
Public Function GetVariableValue(sIn As String) As String
  Dim i As Integer
    
  GetVariableValue = "ERROR"   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Name = LTrim(RTrim(sIn)) Then
      GetVariableValue = g_uVar(i).Value
      Exit Function
    End If
  Next i
End Function

'*******************************************
' This ensures that a line does not end in
' a number or letter. This would suggest
' code is spread out across two lines or
' more.
'*******************************************
Public Function CorrectLineTermination(sInput As String) As String
  Dim sLines() As String
  Dim sOut As String
  Dim i As Integer
  Dim nUpper As Integer
  
  sLines = Split(sInput, vbCrLf)
  nUpper = UBound(sLines) - 1
  
  'loads program into an array
  For i = 0 To UBound(sLines) - 1
    If Len(Trim(sLines(i))) > 0 Then
      If IsSymbol(Right(sLines(i), 1)) = False And i < nUpper Then 'does end in a symbol
        sLines(i) = sLines(i) & " " & sLines(i + 1)
        sLines(i + 1) = ""
      End If
    End If
  Next i
  
  'construct string from array
  For i = 0 To UBound(sLines) - 1
   sOut = sOut & sLines(i) & vbCrLf
  Next i
  
  CorrectLineTermination = sOut
End Function

'*******************************************
' Finds variable declarations and translates
' them to VM code.  Adds to g_uVar() table.
'*******************************************
Public Function TranslateVariables_VM(sInput As String) As String
  Dim sOut As String
  Dim sLines() As String
  Dim i, j As Integer
  Dim nPos As Integer
  Dim sTemp As String 'holds stuff to right of variable type
  Dim sVar(13) As String
  Dim bFound As Boolean
  Dim sVMString As String 'holds VM code built from C code
  Dim nEqual As Integer 'position of equal sign
  Dim sName As String 'name of variable
  Dim sType As String 'type of variable
  Dim sValue As String 'value fo variable
  Dim sReturn As String 'catches results of AddVariable() calls
  
  'allowable variable combinations - must be followed by a space
  sVar(0) = "static unsigned char "
  sVar(1) = "static unsigned int "
  sVar(2) = "static unsigned long "
  sVar(3) = "static char "
  sVar(4) = "static int "
  sVar(5) = "static long "
  sVar(6) = "static float "
  sVar(7) = "unsigned char "
  sVar(8) = "unsigned int "
  sVar(9) = "unsigned long "
  sVar(10) = "char "
  sVar(11) = "int "
  sVar(12) = "long "
  sVar(13) = "float "
  
  sLines = Split(sInput, vbCrLf) 'loads program into array
  
  'evaluates each line of code and looks for a variable declaration
  For i = 0 To UBound(sLines) - 1
    sVMString = "": bFound = False
    
    'consider each possible variable
    For j = 0 To 12
      nPos = InStr(1, sLines(i), sVar(j))
      If nPos > 0 Then  'variable type found
        If Right(RTrim(sLines(i)), 1) <> ";" Then
          TranslateVariables_VM = "ERROR!" & vbCrLf & "Missing semicolong..." & vbCrLf
          TranslateVariables_VM = TranslateVariables_VM & sOut
          Exit Function
        Else
          sLines(i) = Left(sLines(i), Len(sLines(i)) - 1)
        End If
        
        bFound = True
        sVMString = "," & sVar(j) & ","
        sType = sVar(j)
        sTemp = Mid(sLines(i), Len(sVar(j))) 'holds rest of declaration
        nEqual = InStr(1, sTemp, "=")       'position of equal sign
        
        If nEqual > 0 Then  '= sign, could be initialized
          sName = Trim(Left(sTemp, nEqual - 1))
          sValue = Trim(Mid(sTemp, nEqual + 1))
          If Len(sName) < 1 Then
            TranslateVariables_VM = "ERROR!" & vbCrLf & "Incorrect variable declaration..." & vbCrLf
            TranslateVariables_VM = TranslateVariables_VM & sOut
            Exit Function
          End If
          If Len(sValue) < 1 Then
            TranslateVariables_VM = "ERROR!" & vbCrLf & "No value assigned in variable declaration..." & vbCrLf
            TranslateVariables_VM = TranslateVariables_VM & sOut
            Exit Function
          End If
          sVMString = "CVAR " & sName & sVMString & sValue
        Else 'variable created, not intialized
          sName = Trim(sTemp)
          
          If Len(sName) < 1 Then
            TranslateVariables_VM = "ERROR!" & vbCrLf & "Incorrect variable declaration..." & vbCrLf
            TranslateVariables_VM = TranslateVariables_VM & sOut
            Exit Function
          End If
          
          'constructs VM code
          sVMString = "CVAR " & sName & sVMString & "0"
          
          'adds variable to g_uVar() table
          sOut = sOut & AddVariable(sName, sType, sValue)
          If Left(sReturn, 5) = "ERROR" Then  'not enough space for variables
            TranslateVariables_VM = sReturn
            Exit Function
          End If
          
        End If
        Exit For
      End If
    
    Next j
    
    'a variable declaration was found, add VM to sOut
    If bFound = True Then
      sOut = sOut & sVMString & vbCrLf
    Else
      sOut = sOut & sLines(i) & vbCrLf
    End If
  Next i
  
  TranslateVariables_VM = sOut
End Function

'******************************************
' Adds variables to g_uVar() array
'******************************************
Public Function AddVariable(sName As String, sType As String, sValue As String, Optional sScope As String)
  Dim i As Integer
    
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Name = "" Then
      g_uVar(i).Name = sName
      g_uVar(i).Type = sType
      g_uVar(i).Value = sValue
      If Left(sType, 6) = "static" Then 'implied
        g_uVar(i).Scope = "static"
      ElseIf sScope = "static" Then     'passed as an option
        g_uVar(i).Scope = "static"
      Else                              'default auto
        g_uVar(i).Scope = "auto"
      End If
      g_nNumVar = g_nNumVar + 1
      'AddVariable = "OK"
      Exit Function
    End If
  Next i
  
  AddVariable = "ERROR!" & vbCrLf & "  Out of memory for variables. Reduce the number of variables."

End Function

'******************************************
' Aligns code based upon braces and
' semicolons. C program can have
' multiple commands separated by ; on a
' single line of code.
'******************************************
Public Function AlignBraces(sInput As String) As String
  Dim sOut As String
  Dim i, j As Integer
    
  sOut = ""
  
  'find braces
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) = "{" Then
        sOut = sOut & vbCrLf & "{" & vbCrLf
    ElseIf Mid(sInput, i, 1) = "}" Then
        sOut = sOut & vbCrLf & "}" & vbCrLf
    ElseIf Mid(sInput, i, 1) = ";" Then
      sOut = sOut & ";" & vbCrLf
    Else
      sOut = sOut & Mid(sInput, i, 1)
    End If
  Next i
  
  AlignBraces = sOut
End Function

'******************************************
' Authenticates that the code is an
' autonomous function.  Ensures all of
' the necessary lines are present and
' verifies that braces are accounted for.
'******************************************
Public Function IsAutoFunction(sInput As String) As String
  Dim sLines() As String
  Dim sOut As String
  Dim i, j As Integer
  Dim sCode1 As String  'stores required lines of code for auto function
  Dim sCode2 As String
  Dim sCode2b As String 'variant of sCode1
  Dim sCode3 As String
  Dim sCode3b As String 'variant of sCode3
  Dim sCode4 As String
  Dim sCode5 As String
  Dim nCode1 As Integer 'position required lines of code
  Dim nCode2 As Integer
  Dim nCode3 As Integer
  Dim nCode4 As Integer
  Dim nCode5 As Integer
  Dim nBrace As Integer 'used for checking sequence
  Dim nBraceCt As Integer 'counts pairs of braces
    
  'pre-load necessary lines of code
  sCode1 = "void User_Autonomous_Code(void)"
  sCode2 = "while (autonomous_mode)"
  sCode2b = "while(autonomous_mode)"
  sCode3 = "if (statusflag.NEW_SPI_DATA)"
  sCode3b = "if(statusflag.NEW_SPI_DATA)"
  sCode4 = "Getdata(&rxdata)"
  sCode5 = "Putdata(&txdata)"
  
  'verify braces
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) = "{" Then
      nBrace = nBrace + 1
      nBraceCt = nBraceCt + 1
    ElseIf Mid(sInput, i, 1) = "}" Then
      nBrace = nBrace - 1
      If nBrace < 0 Then
        IsAutoFunction = "ERROR!" & vbCrLf & "Incorrect placement of braces '}'"
        Exit Function
      End If
    End If
  Next i
  
  'braces must be matched
  If nBrace > 0 Then
    IsAutoFunction = "ERROR!" & vbCrLf & "Incorrect number of braces '{' or '}'"
    Exit Function
  End If
  
  If nBraceCt < 3 Then
    IsAutoFunction = "ERROR!" & vbCrLf & "Insufficient number of braces '{' for autonomous mode."
    Exit Function
  End If
  
  'verify correct sequence of essential lines
  nCode1 = InStr(1, sInput, sCode1)
  nCode2 = InStr(1, sInput, sCode2)
  nCode3 = InStr(1, sInput, sCode3)
  nCode4 = InStr(1, sInput, sCode4)
  nCode5 = InStr(1, sInput, sCode5)
  If nCode2 < 1 Then nCode2 = InStr(1, sInput, sCode2b)
  If nCode3 < 1 Then nCode3 = InStr(1, sInput, sCode3b)
  
  'MsgBox nCode1 & " " & nCode2 & " " & nCode3 & " " & nCode4 & " " & nCode5
  
  If (nCode1 > 0) And (nCode1 < nCode2) And (nCode2 < nCode3) And (nCode3 < nCode4) And (nCode4 < nCode5) Then
    'essential lines of code in correct order
  Else
    IsAutoFunction = "ERROR!" & vbCrLf & "Essential autonomous lines are missing or out of order." & vbCrLf
    IsAutoFunction = IsAutoFunction & "Verify the following lines are include in this sequence." & vbCrLf
    IsAutoFunction = IsAutoFunction & vbTab & sCode1 & vbCrLf
    IsAutoFunction = IsAutoFunction & vbTab & sCode2 & vbCrLf
    IsAutoFunction = IsAutoFunction & vbTab & sCode3 & vbCrLf
    IsAutoFunction = IsAutoFunction & vbTab & sCode4 & vbCrLf
    IsAutoFunction = IsAutoFunction & vbTab & sCode5 & vbCrLf
    Exit Function
  End If
  
  'IsAutoFunction = sInput
  'Exit Function
  
  'let's get rid of function name and its braces
  nCode1 = InStr(1, sInput, sCode1)
  Mid(sInput, nCode1, Len(sCode1)) = String(Len(sCode1), " ")
  nBrace = InStr(1, sInput, "{")
  Mid(sInput, nBrace, 1) = " "
  
  nCode2 = InStr(1, sInput, sCode2)
  If nCode2 < 1 Then nCode2 = InStr(1, sInput, sCode2b)
  Mid(sInput, nCode2, Len(sCode2)) = String(Len(sCode2), " ")
  nBrace = InStr(1, sInput, "{")
  Mid(sInput, nBrace, 1) = " "
  
  nCode3 = InStr(1, sInput, sCode3)
  If nCode3 < 1 Then nCode3 = InStr(1, sInput, sCode3b)
  Mid(sInput, nCode3, Len(sCode3)) = String(Len(sCode3), " ")
  nBrace = InStr(1, sInput, "{")
  Mid(sInput, nBrace, 1) = " "
  
  nCode4 = InStr(1, sInput, sCode4)
  Mid(sInput, nCode4, Len(sCode4)) = String(Len(sCode4), " ")
  'nBrace = InStr(1, sInput, "{")
  'Mid(sInput, nBrace, 1) = " "
  
  nCode5 = InStr(1, sInput, sCode5)
  Mid(sInput, nCode5, Len(sCode5)) = String(Len(sCode5), " ")
  'nBrace = InStr(1, sInput, "{")
  'Mid(sInput, nBrace, 1) = " "
  
  nBraceCt = 0
  For i = Len(sInput) To 1 Step -1
    If Mid(sInput, i, 1) = "}" Then nBraceCt = nBraceCt + 1
    If nBraceCt = 3 Then
      'MsgBox nBraceCt & " " & i
      sInput = Left(sInput, i - 1)
      Exit For
    End If
  Next i
  
  'legal auto mode
  IsAutoFunction = sInput
End Function

'******************************************
' Removes all lines that are blank and
' trims whitespaces.
'******************************************
Public Function RemoveBlankLines(sInput As String) As String
  Dim sLines() As String
  Dim sOut As String
  Dim i, j As Integer
  Dim nTab As Integer
  
  If Len(sInput) < 1 Then
    RemoveBlankLines = "ERROR!" & vbCrLf & "Nothing to compile"
    Exit Function
  End If
  
  'put program into array
  sLines = Split(sInput, vbCrLf)
  
  'exclude blank rows and trim the rest
  For i = 0 To UBound(sLines) - 1
    
    'gets rid of leading tabs
    nTab = 0
    For j = 1 To Len(sLines(i))
      If Mid(sLines(i), j, 1) = vbTab Then nTab = nTab + 1
    Next j
    If nTab > 0 Then
      sLines(i) = Mid(sLines(i), nTab + 1)
    End If
    
    sLines(i) = Trim(sLines(i))  'get rid of leading/trailing whitespaces
    
    'sometimes a leftover semicolon remains...get rid of it here
    If Len(sLines(i)) > 0 And sLines(i) <> ";" Then  'there is something here...keep this
      sOut = sOut & sLines(i) & vbCrLf
    End If
  Next i
    
  RemoveBlankLines = sOut
End Function

'******************************************
' Removes all comments and leading/lagging
' whitespaces.
'******************************************
Public Function RemoveComments(sInput As String) As String
 
  Dim nComment1 As Integer  'tracks position of comments /* and //
  Dim nComment2 As Integer  'tracks position of comment */
  Dim nCrLf As Integer        'position of vbCrLf
  Dim i, j As Integer
  Dim nCt As Integer  'counts /*
  Dim nCt2 As Integer 'counts */
  Dim sOut As String
  Dim sLines() As String
  
  'count open comments /*
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 2) = "/*" Then
      nCt = nCt + 1
    ElseIf Mid(sInput, i, 2) = "*/" Then
      nCt2 = nCt2 + 1
    End If
  Next i
  
  If nCt <> nCt2 Then  'error with comments /* and */
    sOut = "ERROR!" & vbCrLf
    sOut = sOut & "...Unequal number of /* and */. Verify comments" & vbCrLf
    RemoveComments = sOut
    Exit Function
  End If
    
  'strips out comments /* and */
  If nCt > 0 Then
    For j = 1 To nCt
      nComment1 = InStr(1, sInput, "/*")
      If nComment1 > 0 Then nComment2 = InStr(nComment1 + 1, sInput, "*/")
      If nComment2 > nComment1 And nComment1 > 0 Then
        sInput = Left(sInput, nComment1 - 1) & Mid(sInput, nComment2 + 2)
      End If
    Next j
  End If
  
  'strips out comments //
  nComment1 = 0
  For i = 1 To Len(sInput)
    nComment1 = InStr(nComment1 + 1, sInput, "//")
    nCrLf = InStr(nComment1 + 1, sInput, vbCrLf)
    If nComment1 > 0 And nCrLf > 0 Then
      sOut = Left(sInput, nComment1 - 1)
      sOut = sOut & Mid(sInput, nCrLf)
      sInput = sOut
    End If
  Next i
    
  'strips out leading/trailing white spaces
  sLines = Split(sInput, vbCrLf)
  
  If UBound(sLines) < 0 Then
    sOut = "ERROR!" & vbCrLf
    sOut = sOut & "...No code" & vbCrLf
    RemoveComments = sOut
    Exit Function
  End If
  
  sOut = ""
  For i = 0 To UBound(sLines) - 1
    sLines(i) = Trim(sLines(i))
    sOut = sOut & sLines(i) & vbCrLf
  Next i
    
  RemoveComments = sOut
End Function

'****************************************
' Loads global variables for this
' module TranslatorC.Bas
'****************************************
Public Sub LoadTranslatorCVariables()
  g_sTranslatorC_Version = "0.01"
End Sub

