Attribute VB_Name = "Translator"
'****************************************************************************
' Translator.bas - Written by Chuck Bolin, April 2005
' Contains functions and subs that parse and compile "C" source code
' into Virtual Machine (VM) code.
' Functions:
' Translate(sInput) return VM Code string
' CleanCCode(sInput) returns string with consecutive spaces or spaces before/
'                     after symbols
' AlignBraces(sInput) returns string with { } on separate lines
' RemoveCommentsWhitespaces(sInput) returns string without comments, leading/
'                                   trailing whitespaces
' ClearCodeArray()  'clears g_sCode() array holding "C" source code
' BuildSystemVariables() adds system variables to var() and creates required
'                        VM code. This happens for all programs automatically.
' IncrementVariable(sIn) returns VM code for "C" ++ operator.
' AddForLoop(sIn) creates a string of VM code for "C" for loop
' CreateVariable(sIn) creates a string of VM code for "C" variable creation.
' ClearVariables() clears all variables in var() array
' AddVariable(sName, sType, sValue, sScope) adds variable to var()
' GetVariableValue(sIn) returns value of existing variable or ERROR
' GetVariableType(sIn) returns type of existing variable or ERROR
' SetVariable(sName, sVal) sets variable with a value or returns ERROR
'****************************************************************************

Option Explicit

'*****************************************
' T R A N S L A T E
' Converts C code into VM code. This is
' main function that calls others.
'*****************************************
Public Function Translate(sInput) As String
  Dim sOut As String  'stores VM code
  Dim sName As String 'variable name(symbol)
  Dim sValue As String 'value of variable
  Dim nPos As Integer 'position of something...used many times
  Dim i, j As Integer
  Dim nStart, nEnd As Integer 'start and stop markers for multiple lines of related code
  Dim nBrace As Integer 'counts braces. If {, count++, else }, count--. Should be equal number of braces
  Dim nMaxElements As Integer 'max elements in array
  Dim nCt As Integer 'counter of things
  Dim sTemp As String
  Dim nPassCount As Integer
  Dim sReturn As String
  Dim sVar As String 'stores sytem variables
  
  Translate = "Nothing to Load/Compile!"
  
  'build system RC variables
  sVar = BuildSystemVariables
  
  If Len(sInput) < 1 Then Exit Function
  nPassCount = 0
PassTo:
  sOut = ""
  nPassCount = nPassCount + 1
  g_sCode = Split(sInput, vbCrLf) 'stores C code into array
  nMaxElements = UBound(g_sCode)
  
  'read and evaluate each line in array
  'this stuff is case sensitive
  For i = 0 To nMaxElements - 1
    g_sCode(i) = CleanCCode(g_sCode(i))
    
    'VM code - do to multiple passes, it is important that these VM code
    'lines remain untouched
    If Left(g_sCode(i), 5) = "SVAR " Or Left(g_sCode(i), 5) = "CVAR " Then
      sOut = sOut & g_sCode(i) & vbCrLf
    ElseIf Left(g_sCode(i), 4) = "GLR " Or Left(g_sCode(i), 4) = "JMP " Then
      sOut = sOut & g_sCode(i) & vbCrLf
    ElseIf Left(g_sCode(i), 4) = "END " Or Left(g_sCode(i), 5) = "SOUT " Then
      sOut = sOut & g_sCode(i) & vbCrLf
    ElseIf Left(g_sCode(i), 4) = "INC " Or Left(g_sCode(i), 4) = "DEC " Then
      sOut = sOut & g_sCode(i) & vbCrLf
    ElseIf g_sCode(i) = "{" Or g_sCode(i) = "}" Then
        
    'automatic variables - not static, these values are deleted every 26 mSec
    ElseIf Left(g_sCode(i), 4) = "int " Then  'variable 'int'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 5) = "char " Then 'variable 'char'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 5) = "long " Then 'variable 'long'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 13) = "unsigned int " Then 'variable 'unsigned int'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 14) = "unsigned char " Then 'variable 'unsigned char'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 14) = "unsigned long " Then 'variable 'unsigned long'
      sOut = sOut & CreateVariable(g_sCode(i))
      
    'static variables - these variables persist during the entire program time
    ElseIf Left(g_sCode(i), 11) = "static int " Then  'variable 'int'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 12) = "static char " Then 'variable 'char'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 12) = "static long " Then 'variable 'long'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 20) = "static unsigned int " Then 'variable 'unsigned int'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 21) = "static unsigned char " Then 'variable 'unsigned char'
      sOut = sOut & CreateVariable(g_sCode(i))
    ElseIf Left(g_sCode(i), 21) = "static unsigned long " Then 'variable 'unsigned long'
      sOut = sOut & CreateVariable(g_sCode(i))
      
    'somewhat messy.  a for() can have braces or no braces. If no braces, then keep
    'only one line after for(), otherwise find all lines up to '}'
    ElseIf Left(g_sCode(i), 4) = "for " Or Left(g_sCode(i), 4) = "for(" Then 'for loop
      'i represents element with for
      nCt = 0
      nBrace = 0
      nStart = 0
      nEnd = 0
      sTemp = ""
      
      'let's find the braces
      For j = i To nMaxElements - 1
        If InStr(1, g_sCode(j), "{") > 0 Then
          If nStart = 0 Then nStart = j  'stores starting brace
          nBrace = nBrace + CountChar(g_sCode(j), "{")
        End If
        If InStr(1, g_sCode(j), "}") > 0 Then
          nCt = CountChar(g_sCode(j), "}")
          nBrace = nBrace - nCt
        End If
        If nCt > 0 And nBrace = 0 Then 'means a } was found and they are matching
          nEnd = j
          Exit For
        End If
      Next j
     
      If nEnd > nStart And nStart > 0 Then 'means there is a starting/ending brace
        For j = i To nEnd
          sTemp = sTemp & g_sCode(j) & vbCrLf
        Next j
      
      'no braces
      ElseIf InStr(1, g_sCode(i), "{") < 1 And InStr(1, g_sCode(i + 1), "{") < 1 Then
        sTemp = sTemp & g_sCode(i) & vbCrLf & g_sCode(i + 1) & vbCrLf
        i = i + 1 'skip over this line
      End If
      
      'something to process
      If Len(sTemp) > 0 Then
        i = i + (nEnd - nStart)  'must advance pointer to jump over 'for' code
        sOut = sOut & AddForLoop(sTemp)
      End If
     
    ElseIf Left(g_sCode(i), 3) = "if " Or Left(g_sCode(i), 4) = "if(" Then 'if loop
    
    ElseIf InStr(1, g_sCode(i), "=") > 0 And InStr(1, g_sCode(i), "==") < 1 Then 'assume it is an assignment
      nPos = InStr(1, g_sCode(i), "=")
      sName = LTrim(RTrim(Left(g_sCode(i), nPos - 1)))
      sValue = LTrim(RTrim(Mid(g_sCode(i), nPos + 1)))
      If Right(sValue, 1) = ";" Then sValue = Left(sValue, Len(sValue) - 1)
      If Len(sName) > 0 And Len(sValue) > 0 Then
        sReturn = SetVariable(sName, sValue)
        If sReturn = "ERROR" Then
          MsgBox "Identifier " & sName & " does not exist!", vbOKOnly, "Unknown identifier"
          Exit Function
        End If
        sOut = sOut & "SVAR " & Trim(sName) & "," & Trim(sValue) & vbCrLf
      End If
    
    'this is increment by one function
    ElseIf InStr(1, g_sCode(i), "++") > 0 Then
      
      sOut = sOut & IncrementVariable(g_sCode(i))
    End If
  
  Next i
  If nPassCount <= 3 Then
    sInput = sOut
    
    GoTo PassTo
  End If
  'sOut = "Intermediate file..." & vbCrLf & "--------------------" & vbCrLf & sOut
 
  Translate = sVar & sOut & "END" & vbCrLf
  'MsgBox sOut
  
End Function

'this eliminates leading, trailing spaces, and other spaces
'that are two or more consecutive spaces. Deletes spaces before
'and after ( ) ; and math operators
Public Function CleanCCode(sInput As String) As String
  Dim i As Integer
  Dim bFound As Boolean
  Dim sTemp As String 'holds new cleaned up string
  Dim sBefore As String
  
  bFound = False
  
  '
  '//test code
  '//**************************
  '       //Sample 2
  'int i        ;
  'int           m;
  'for(i  =  0  ;    i   <    10  ;  i++  )
  '{
  'm++        ;
  '      }
  '//****************

  sInput = LTrim(RTrim(sInput))  'deletes leading and trailing spaces
  sBefore = sInput
  'narrows the string to single spaces where found
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) = " " And bFound = True Then
      'nothing
    ElseIf Mid(sInput, i, 1) = " " And bFound = False Then
      bFound = True
      sTemp = sTemp & " "
    Else
      sTemp = sTemp & Mid(sInput, i, 1)
      bFound = False
    End If
  Next i
  sInput = sTemp
  
  'now get rid of leading spaces before special characters
  sTemp = ""
  'sBefore = sInput
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) = " " And IsSymbol(Mid(sInput, i + 1, 1)) = True Then
      'do nothing...don't keep this space
      sTemp = sTemp & Mid(sInput, i + 1, 1)
      i = i + 1
      'MsgBox "YES"
    Else
      'sTemp = sTemp & Mid(sInput, i, 1)
    End If
  Next i
  'MsgBox "|" & sBefore & "|" & sInput & "|"

  'now get rid of trailing spaces after special characters
  sTemp = ""
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 2) = " " And IsSymbol(Mid(sInput, i, 1)) = True Then
      sTemp = sTemp & Mid(sInput, i, 1)
      i = i + 1
      'do nothing...don't keep this space
    Else
      sTemp = sTemp & Mid(sInput, i, 1)
    End If
  Next i
  
    
  CleanCCode = sTemp
End Function

'*****************************************
' A L I G N  B R A C E S
' Places a brace on a single line by
' itself.
'*****************************************
Public Function AlignBraces(sInput As String) As String
  Dim sOut As String
  Dim i As Integer
  
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 1) = "{" Then
      If Mid(sInput, i - 2, 2) = vbCrLf Then
        sOut = sOut & "{"
      Else
        sOut = sOut & vbCrLf & "{"
      End If
    ElseIf Mid(sInput, i, 1) = "}" Then
      If Mid(sInput, i - 2, 2) = vbCrLf Then
        sOut = sOut & "}"
      Else
        sOut = sOut & vbCrLf & "}"
      End If
    Else
      sOut = sOut & Mid(sInput, i, 1)
    End If
  Next i
  
  AlignBraces = sOut
End Function

'******************************************
' Removes all comments and leading/lagging
' whitespaces.
'******************************************
Public Function RemoveCommentsWhitespaces(sInput As String) As String
 
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
    RemoveCommentsWhitespaces = sOut
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
    RemoveCommentsWhitespaces = sOut
    Exit Function
  End If
  
  sOut = ""
  For i = 0 To UBound(sLines) - 1
    sLines(i) = LTrim(RTrim(sLines(i)))
    sOut = sOut & sLines(i) & vbCrLf
  Next i
    
  RemoveCommentsWhitespaces = sOut
End Function

Public Sub ClearCodeArray()
  Dim i As Integer
  For i = 0 To MAX_CODE_LINES - 1
    g_sCode(i) = ""
  Next i
  g_nMaxLines = 0
End Sub

'***********************************************
' This creates a string of VM Code representing
' robot controller variables such as digital
' I/O, pwms, etc.
'***********************************************
Public Function BuildSystemVariables() As String
  Dim sOut As String
  
  'adds to variable table
  AddVariable "pwm01", "unsigned char", "127", "static"
  AddVariable "pwm02", "unsigned char", "127", "static"
  AddVariable "pwm03", "unsigned char", "127", "static"
  AddVariable "pwm04", "unsigned char", "127", "static"
  AddVariable "User_Mode_byte", "unsigned char", "0", "static"
  AddVariable "Switch1_LED", "unsigned char", "0", "static"
  AddVariable "Switch2_LED", "unsigned char", "0", "static"
  AddVariable "Switch3_LED", "unsigned char", "0", "static"
  AddVariable "Pwm1_red", "unsigned char", "0", "static"
  AddVariable "Pwm2_red", "unsigned char", "0", "static"
  AddVariable "Pwm1_green", "unsigned char", "0", "static"
  AddVariable "Pwm2_green", "unsigned char", "0", "static"
  AddVariable "Relay1_red", "unsigned char", "0", "static"
  AddVariable "Relay2_red", "unsigned char", "0", "static"
  AddVariable "Relay1_green", "unsigned char", "0", "static"
  AddVariable "Relay2_green", "unsigned char", "0", "static"
  AddVariable "relay1_fwd", "unsigned char", "0", "static"
  AddVariable "relay2_fwd", "unsigned char", "0", "static"
  AddVariable "relay3_fwd", "unsigned char", "0", "static"
  AddVariable "relay4_fwd", "unsigned char", "0", "static"
  AddVariable "relay1_rev", "unsigned char", "0", "static"
  AddVariable "relay2_rev", "unsigned char", "0", "static"
  AddVariable "relay3_rev", "unsigned char", "0", "static"
  AddVariable "relay4_rev", "unsigned char", "0", "static"
  
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

'increments variable by 1
Public Function IncrementVariable(sIn) As String
  Dim nPlus As Integer 'position of first + in ++
  Dim sName As String 'var name
  nPlus = InStr(1, sIn, "++")
  If nPlus > 1 Then
    sName = Left(sIn, nPlus - 1)
  End If
  IncrementVariable = "INC " & sName & ",1" & vbCrLf
  'MsgBox "IV:" & vbCrLf & vbCrLf & vbCrLf & IncrementVariable
  
End Function

'process a for loop
'returns IL format
Public Function AddForLoop(sIn As String) As String
  Dim sLines() As String 'stores received string as an array
  Dim i As Integer
  Dim nSemiColon As Integer 'number of semicolons
  Dim nLP, nRP As Integer 'positions of parenthesis
  Dim nPos1, nPos2 As Integer 'positions of semicolons in parameters
  Dim sInit As String  'components of for() parameters
  Dim sCond As String
  Dim sOp As String
  Dim nOp, nLen As Integer
  Dim sExpr As String
  Dim nEqual As Integer 'position of first equal sign..initial
  Dim sInitValue As String 'var value
  Dim sName As String 'var name
  Dim sJump As String 'number to jump to skip all for() code
  Dim sOut As String
  Dim sCheck As String
  
  AddForLoop = ""
  
  sLines = Split(sIn, vbCrLf) 'load array
  
  'find parens with parameters
  nLP = InStr(1, sLines(0), "(")
  nRP = InStr(1, sLines(0), ")")
  If nLP < 1 Or nRP < 1 Or nRP < nLP Then Exit Function
  
  nSemiColon = CountChar(sLines(0), ";")
  If nSemiColon <> 2 Then Exit Function
  
  nPos1 = InStr(1, sLines(0), ";")
  nPos2 = InStr(nPos1 + 1, sLines(0), ";")
  
  sInit = LTrim(RTrim(Mid(sLines(0), nLP + 1, nPos1 - nLP - 1)))
  sCond = LTrim(RTrim(Mid(sLines(0), nPos1 + 1, nPos2 - nPos1 - 1)))
  sExpr = LTrim(RTrim(Mid(sLines(0), nPos2 + 1, nRP - nPos2 - 1)))

  If Len(sInit) < 1 And Len(sCond) < 1 And Len(sExpr) < 1 Then
    AddForLoop = "************" & vbCrLf & "Special Feature" & vbCrLf & "*************" & vbCrLf
  Else
    
    'find out our jump value..you'll see how this works
    If UBound(sLines) > 2 Then
      sJump = CStr(UBound(sLines)) '+ 1
    Else
      sJump = "4"
    End If
    
    'verify variable exists
    nEqual = InStr(1, sInit, "=")
    If nEqual < 1 Then Exit Function
    sName = Left(sInit, nEqual - 1) 'initial variable and value
    sCheck = GetVariableValue(sName)
    If sCheck = "ERROR" Then
      MsgBox "Identifier " & sName & " does not exist!", vbOKOnly, "Unknown identifier"
      Exit Function
    End If
       
    sInitValue = Mid(sInit, nEqual + 1)
    sOut = sOut & "SVAR " & Trim(sName) & "," & Trim(sInitValue) & vbCrLf
    
    'get operator < or > or <= or >=
    nLen = 2
    nOp = InStr(1, sCond, "<=")
    If nOp < 1 Then nOp = InStr(1, sCond, ">=")
    If nOp < 1 Then nOp = InStr(1, sCond, "==")
    If nOp < 1 Then nLen = 1
    If nOp < 1 Then nOp = InStr(1, sCond, ">")
    If nOp < 1 Then nOp = InStr(1, sCond, "<")
    
    If nOp > 1 Then
      sOp = Mid(sCond, nOp, nLen)
      sOut = sOut & "GLR " & Trim(Left(sCond, nOp - 1)) & "," & Trim(sOp) & "," & Trim(Mid(sCond, nOp + nLen)) & "," & "[" & Trim(sJump) & "]" & vbCrLf
      For i = 1 To UBound(sLines)
        If Left(sLines(i), 1) = "{" Then
          'ignore this
        ElseIf Left(sLines(i), 1) = "}" Then
          'ignore this also
        ElseIf Len(LTrim(RTrim(sLines(i)))) < 1 Then
          'ignore blanks
        Else  'can't ignore this stuff
          sOut = sOut & sLines(i) & vbCrLf
        End If
      Next i
    End If
    'MsgBox sExpr
    sOut = sOut & sExpr & vbCrLf
    sOut = sOut & "JMP " & "[-" & CStr(CInt(sJump) - 1) & "]" & vbCrLf
    'MsgBox sOut
    'AddForLoop = "************" & vbCrLf & sOut & "*************" & vbCrLf
    AddForLoop = sOut
  End If
End Function

'********************************************** VARIABLES *******************
'adds variable to var( ) look up table and
'generates equivalent Virtual Machine (VM) code
'Two types allowed:
'1)  type var;
'2)  type var = value;
Public Function CreateVariable(sIn As String) As String
  Dim nEqual As Integer 'position of equal sign
  Dim nSpace As Integer 'first space after variable type
  Dim nSpace2 As Integer 'second space after variable type
  Dim sType As String 'stores variable type
  Dim sName As String 'stores variable name
  Dim sValue As String 'stores variable value
  Dim sScope As String 'stores auto or static
  Dim sCheck As String 'verifies variable doesn't already exist
  Dim nOffset As Integer 'used for determining auto or static
  Dim nStatic As Integer 'holds position of static word
  Dim nUnsigned As Integer 'holds position of unsigned word
    
  nEqual = InStr(1, sIn, "=") 'get position of equal sign
  nStatic = InStr(1, sIn, "static")
  nUnsigned = InStr(1, sIn, "unsigned")
  'MsgBox "Static: " & nStatic
  'several combinations of static and unsigned
  If nUnsigned < 1 And nStatic < 1 Then  'not static or unsigned
    nSpace = InStr(1, sIn, " ")
    sType = Trim(Mid(sIn, 1, nSpace - 1))
  ElseIf nUnsigned > 0 And nStatic < 1 Then 'not static but unsigned
    nSpace = InStr(nUnsigned + 10, sIn, " ") 'get position of first space
    sType = Trim(Mid(sIn, 1, nSpace - 1))
  ElseIf nUnsigned < 1 And nStatic > 0 Then 'static but not unsigned
    nSpace = InStr(nStatic + 8, sIn, " ") 'get position of first space
    sType = Trim(Mid(sIn, 1, nSpace - 1))
  ElseIf nUnsigned > 0 And nStatic > 0 Then 'static and unsigned
    nSpace = InStr(17, sIn, " ") 'get position of first space
    sType = Trim(Mid(sIn, 1, nSpace - 1))
  End If
  
  nOffset = 0
  CreateVariable = ""

  If Right(sIn, 1) = ";" Then sIn = Left(sIn, Len(sIn) - 1) 'strip off semicolon
 
  If nSpace < 1 Then Exit Function 'needs a space
 
  If nEqual < 1 Then 'no value assigned, use default
    sName = Trim(Mid(sIn, nSpace))
  Else    'value assigned
    sName = Trim(Mid(sIn, nSpace, nEqual - nSpace))
    sValue = Trim(Mid(sIn, nEqual + 1))
  End If
  'MsgBox sName & " " & sType
  
  sCheck = GetVariableValue(sName)
  'MsgBox Len(sCheck)
  
  If sCheck = "ERROR" Then  'variable not in table, OK to add now
    'MsgBox "OK"
    If nEqual < 1 Then 'no value assigned, use default
      sValue = "0"
    End If
     
    'MsgBox sName & ": " & sType & " " & Len(sType)
    
     
    If sType = "char" And CLng(sValue) >= -128 And CLng(sValue) <= 127 Then AddVariable sName, sType, sValue
    If sType = "int" And CLng(sValue) >= -32768 And CLng(sValue) <= 32768 Then AddVariable sName, sType, sValue
    If sType = "long" And CLng(sValue) >= -2147483648# And CLng(sValue) <= 2147483647 Then AddVariable sName, sType, sValue
    If sType = "unsigned char" And CLng(sValue) >= 0 And CLng(sValue) <= 127 Then AddVariable sName, sType, sValue
    If sType = "unsigned int" And CLng(sValue) >= 0 And CLng(sValue) <= 32768 Then AddVariable sName, sType, sValue
    If sType = "unsigned long" And CLng(sValue) >= 0 And CLng(sValue) <= 2147483647 Then AddVariable sName, sType, sValue
    If sType = "static char" And CLng(sValue) >= -128 And CLng(sValue) <= 127 Then AddVariable sName, sType, sValue, sScope
    If sType = "static int" And CLng(sValue) >= -32768 And CLng(sValue) <= 32768 Then AddVariable sName, sType, sValue, sScope
    If sType = "static long" And CLng(sValue) >= -2147483648# And CLng(sValue) <= 2147483647 Then AddVariable sName, sType, sValue, sScope
    If sType = "static unsigned char" And CLng(sValue) >= 0 And CLng(sValue) <= 127 Then AddVariable sName, sType, sValue, sScope
    If sType = "static unsigned int" And CLng(sValue) >= 0 And CLng(sValue) <= 32768 Then AddVariable sName, sType, sValue, sScope
    If sType = "static unsigned long" And CLng(sValue) >= 0 And CLng(sValue) <= 2147483647 Then AddVariable sName, sType, sValue, sScope

    sCheck = GetVariableValue(sName)
     
    If sCheck <> "ERROR" Then
      'MsgBox sName & "  " & sType & "   " & sValue & "  " & sScope
      CreateVariable = "CVAR " & Trim(sName) & "," & Trim(sType) & "," & Trim(sValue) & vbCrLf
    End If
  End If
  
End Function

'clears all variables from var() array
Public Sub ClearVariables()
  Dim i As Integer
  
  For i = 0 To MAX_VARIABLES
    var(i).Symbol = ""
    var(i).Type = ""
    var(i).Value = ""
  Next i
End Sub

'adds variable
Public Sub AddVariable(sName As String, sType As String, sValue As String, Optional sScope As String)
  Dim i As Integer
    
  For i = 0 To MAX_VARIABLES
    If var(i).Symbol = "" Then
      var(i).Symbol = sName
      var(i).Type = sType
      var(i).Value = sValue
      If Len(sScope) < 1 Then
        var(i).Scope = "auto"
      Else
        var(i).Scope = "static"
      End If
      Exit Sub
    End If
  Next i

End Sub

'get value assigned to a variable
Public Function GetVariableValue(sIn As String) As String
  Dim i As Integer
    
  GetVariableValue = "ERROR"   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If var(i).Symbol = LTrim(RTrim(sIn)) Then
      GetVariableValue = var(i).Value
      Exit Function
    End If
  Next i
End Function

'gets variable type
Public Function GetVariableType(sIn As String) As String
  Dim i As Integer
    
  GetVariableType = "ERROR"   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If var(i).Symbol = LTrim(RTrim(sIn)) Then
      GetVariableType = var(i).Type
      Exit Function
    End If
  Next i

End Function

'sets current variable to a value
Public Function SetVariable(sName As String, sVal As String) As String
  Dim i As Integer
  SetVariable = "ERROR"
  For i = 0 To MAX_VARIABLES
    If var(i).Symbol = Trim(sName) Then
      'MsgBox Len(sName) & " " & sName
      var(i).Value = Trim(sVal)
      SetVariable = ""
      Exit Function
    End If
  Next i
End Function


