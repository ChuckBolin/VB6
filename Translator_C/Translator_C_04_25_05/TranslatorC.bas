Attribute VB_Name = "TranslatorC"
'****************************************************************************
' Translator.bas - Written by Chuck Bolin Team 342, April 2005
' Contains functions and subs that parse and compile "C" source code
' into Virtual Machine (VM) code.
' Functions:

' TranslationC_VM(sInput) returns translated code
' IsVMCode(sInput) returns true if line is VM code
' C_For_Oneliner(sInput) standalones lines converted to VM
' BuildSystemVariables() adds all RC default I/Os to g_uVar() table
' C_DecrementVariable(sIn) or returns ERROR
' C_IncrementVariable(sIn) or returns ERROR
' C_For_Twoliner(sInput) converts 'for' to VM without braces
' GetVariableValue(sIn) returns value of variable or ERROR
' CorrectLineTermination(sInput)splices two lines together ending in letter
' TranslateVariables_VM(sInput) finds all variables and code -> VM code
' AddVariable(sName, sType, sValue, Optional sScope) adds variable to g_uVar()
' AlignSemicolons(sInput) splits lines with semicolons
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

Public Type FUNCTION_FOR 'used for returning the results of a parsed 'for'
  Variable As String
  Initial As String
  Condition As String
  Increment As String
  Valid As Boolean 'true if valid format
End Type

'public variables
'****************
Public g_sTranslatorC_Version As String
Public g_bUpperPWM As Boolean 'true if Generate_Pwms(pwm13,pwm14,pwm15,pwm16);
Public g_uVar(MAX_VARIABLES) As VARIABLE_TABLE
Public g_nNumVar As Integer 'number of variables actually in use
Public g_sErrorNum As String

'*******************************************
' Translate is the function that performs
' all of the magic of converting "C" code
' to VM code.
'*******************************************
Public Function Translate(sInput As String) As String
  Dim nLength As Long 'length of file
  Dim uPair As GENERIC_PAIR
  Dim sTemp As String 'temp file storage
  Dim bTryAgain As Boolean 'use for repeating code
  Dim nPos As Long 'position of any character
  Dim nLBrace As Long 'position braces
  Dim nRBrace As Long '
  Dim nPos2 As Long 'another position marker in the string
  Dim nPos3 As Long
  Dim uff As FUNCTION_FOR
  Dim nLast As Long
  Dim sLines() As String
  Dim nJump As Integer
  Dim sMiddle As String
  Dim i As Long
  Dim nPass As Integer
  
  'remove comments
  i = 1
  sInput = RemoveComments(sInput)
  Do
    uPair = GetStringInnerBraces(Mid(sInput, i))
    If uPair.First > 0 And uPair.Second > 0 Then
      'MsgBox uPair.First & "  " & uPair.Second & "  " '&  Mid(sInput, uPair.First)
      If CheckForJumpCode(Mid(sInput, uPair.First, uPair.Second - uPair.First - 1), "for") = True Then
        MsgBox "for" & "  " & uPair.First & "  " & uPair.Second & "  " & Mid(sInput, uPair.First, uPair.Second - uPair.First - 1)
      ElseIf CheckForJumpCode(Mid(sInput, uPair.First, uPair.Second - uPair.First - 1), "while") = True Then
        MsgBox "while"
      End If
      i = uPair.Second + 1
      bTryAgain = True
    Else
      bTryAgain = False
    End If
  
  Loop Until bTryAgain = False
PassBegin:
  
  
  'find any 'for' loops without braces
  '************************************
  nLast = 1
RepeatFor:
  nPos = InStr(nLast, sInput, "for")
  nLast = nPos + 1
  'MsgBox nPos & "  " & Len(sInput) & "  " & Mid(sInput, nPos)
  If nPos > 0 Then
    
    If VerifySequence(sInput, "for$($;$;$)") = True Then  'found valid 'for'
      
      nPos2 = InStr(nPos, sInput, ")")
      If VerifyNextCharacter(nPos2 + 1, "{", sInput) = False Then 'not found left brace
        nPos3 = InStr(nPos2 + 2, sInput, vbCrLf)
        uff = ParseFor(Mid(sInput, nPos, nPos2 - nPos + 1))
        sTemp = ""
        sTemp = Left(sInput, nPos - 1)
        sTemp = sTemp & "SVAR " & uff.Variable & "," & uff.Initial & vbCrLf
        sTemp = sTemp & "GLR " & uff.Condition & ",3" & vbCrLf
        sTemp = sTemp & uff.Increment & vbCrLf
        sTemp = sTemp & "JMP -2" & vbCrLf
        sTemp = sTemp & Mid(sInput, nPos3)
        sInput = sTemp
      
      End If
    End If
  End If
  sInput = RemoveBlankLines(sInput)
  If nPos > 0 Then GoTo RepeatFor
  
 
  
  
  'find 'for' with braces
  '************************
  nLast = 1
RepeatFor2:
  nPos = InStr(nLast, sInput, "for")
  nLast = nPos + 1
  'MsgBox nPos & "  " & Len(sInput) & "  " & Mid(sInput, nPos)
  If nPos > 0 Then
    
    If VerifySequence(sInput, "for$($;$;$)") = True Then  'found valid 'for'
      
      nPos2 = InStr(nPos, sInput, ")")
      'If VerifyNextCharacter(nPos2 + 1, "{", sInput) = True Then 'not found left brace
      
      If VerifyNextCharacter(nPos2 + 1, "{", sInput) = True Then 'found left brace
        uPair = FindMatchingPair(nPos2, "{", "}", sInput)
      
        'this 'for' has matching braces and no braces in between
        If uPair.First > 0 And uPair.Second > 0 And uPair.InBetween = False Then
          uff = ParseFor(Mid(sInput, nPos, nPos2 - nPos + 1))
          If uff.Valid = True Then
      
            'let's figure out the lines between the two braces
            sLines = Split(Mid(sInput, uPair.First, uPair.Second - uPair.First - 1), vbCrLf)
            nJump = 0
      
            'parse this section
            sMiddle = ""
            For i = 0 To UBound(sLines)
              sLines(i) = Trim(sLines(i))
              'MsgBox nPass & "  " & sLines(i)
              If Len(sLines(i)) > 1 Then
                sMiddle = sMiddle & sLines(i) & vbCrLf
                nJump = nJump + 1
              End If
            Next i
            sTemp = ""
            sTemp = Left(sInput, nPos - 1)
            sTemp = sTemp & "SVAR " & uff.Variable & "," & uff.Initial & vbCrLf
            sTemp = sTemp & "GLR " & uff.Condition & "," & CStr(nJump + 3) & vbCrLf
            sTemp = sTemp & uff.Increment & vbCrLf
            sTemp = sTemp & sMiddle
            sTemp = sTemp & "JMP -" & CStr(nJump + 2) & vbCrLf
            sTemp = sTemp & Mid(sInput, uPair.Second + 1)
            sInput = sTemp
          Else
            sInput = Mid(sInput, nPos, nPos2 - nPos - 1) & vbCrLf & "ERROR!" & vbCrLf
            sInput = sInput & "...invalid for format"
            Translate = sInput
            Exit Function
          End If
      
        'there are other braces inside of these two...do nothing
        ElseIf uPair.First > 0 And uPair.Second > 0 And uPair.InBetween = True Then
          'do nothing here for now
        Else  'no right brace found...error
          sInput = Left(sInput, nPos2) & vbCrLf & "ERROR!" & vbCrLf
          sInput = sInput & "...missing right brace '}'"
          Translate = sInput
          Exit Function
        End If
     End If
    End If
  End If
  sInput = RemoveBlankLines(sInput)
   
  If nPos > 0 Then GoTo RepeatFor2
  'Translate = sInput
   'Exit Function
   
  nPass = nPass + 1
  If nPass < 3 Then GoTo PassBegin
   
   
  Translate = sInput
End Function

'*******************************************
' returns true if one of these multiline
' functions exist (requires jumps)
' The goal is not to translate code inside
' of braces if it contains an if, else, for
' or while, or switch
'*******************************************
Public Function CheckForJumpCode(sInput As String, sFilter As String) As Boolean
  Dim i As Integer
  
  sInput = Trim(sInput)
  If Len(sInput) < 1 Then Exit Function
  If Len(sFilter) < 1 Then Exit Function
  
  Select Case sFilter
    Case "for":
      If VerifySequence(sInput, "for$($;$;$)") = True Then CheckForJumpCode = True
    Case "while":
      If VerifySequence(sInput, "while$($)") = True Then CheckForJumpCode = True
    Case "if":
      If VerifySequence(sInput, "if$($)") = True Then CheckForJumpCode = True
    Case "else":
      If VerifySequence(sInput, "else") = True And VerifySequence(sInput, "else$if$($)") = False Then CheckForJumpCode = True
    Case "elseif":
      If VerifySequence(sInput, "else$if$($)") = True Then CheckForJumpCode = True
  End Select
  
End Function

'*******************************************
' This parses the stuff in for(    ) and
' returns results.
'*******************************************
Public Function ParseFor(sInput As String) As FUNCTION_FOR
  Dim uff As FUNCTION_FOR
  Dim nPosLeftParen As Integer
  Dim nPosRightParen As Integer
  Dim nPosFirstSemi As Integer
  Dim nPosSecondSemi As Integer
  Dim nPosEqual As Integer
      
  uff.Valid = False
  'MsgBox sInput
  
  If Len(sInput) < 1 Then Exit Function
  If VerifySequence(sInput, "($;$;$)") = True Then
    nPosLeftParen = InStr(1, sInput, "(")
    nPosFirstSemi = InStr(nPosLeftParen + 1, sInput, ";")
    nPosSecondSemi = InStr(nPosFirstSemi + 1, sInput, ";")
    nPosRightParen = InStr(nPosSecondSemi + 1, sInput, ")")
    'MsgBox nPosLeftParen & "  " & nPosFirstSemi & "  " & nPosSecondSemi & "  " & nPosRightParen
    'MsgBox (nPosRightParen > nPosSecondSemi)
    'MsgBox (nPosSecondSemi > nPosFirstSemi)
    'MsgBox (nPosFirstSemi > nPosLeftParen)
    
    'must be valid values
    If (nPosRightParen > nPosSecondSemi) And (nPosSecondSemi > nPosFirstSemi) And (nPosFirstSemi > nPosLeftParen) Then
      uff.Valid = True
      uff.Initial = Trim(Mid(sInput, nPosLeftParen + 1, nPosFirstSemi - nPosLeftParen - 1))
      nPosEqual = InStr(1, uff.Initial, "=")
      If nPosEqual > 0 Then
        uff.Variable = Trim(Left(uff.Initial, nPosEqual - 1))
        uff.Initial = Trim(Mid(uff.Initial, nPosEqual + 1))
      End If
      uff.Condition = Trim(Mid(sInput, nPosFirstSemi + 1, nPosSecondSemi - nPosFirstSemi - 1))
      uff.Increment = Trim(Mid(sInput, nPosSecondSemi + 1, nPosRightParen - nPosSecondSemi - 1))
      ParseFor = uff
      Exit Function
    Else
      Exit Function
    End If
  Else
    Exit Function
  End If
  
End Function

'*******************************************
' Translates C code into VM, Returns VM
'*******************************************
Public Function TranslationC_VM(sInput As String) As String

  Dim sReturn As String
  Dim sSystem As String 'holds system variables
  Dim uPair As GENERIC_PAIR
  
  'get rid of all comments /*, */ and //
  g_sErrorNum = "100"
  sReturn = sInput
  sReturn = RemoveComments(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum:  Exit Function
  
  'put all braces { and } on separate lines in array
  g_sErrorNum = "110"
  sReturn = AlignBraces(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
  
  'could be multiple commands on one line, or one line is
  'spread out across two or more lines
  g_sErrorNum = "120"
  sReturn = AlignSemicolons(sReturn)
  
  'remove all blank lines
  g_sErrorNum = "130"
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
  
  'verify code written adheres to Autonomous Code formatting and then removes
  'these lines of code with their braces
  g_sErrorNum = "140"
  sReturn = IsAutoFunction(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
  
  'creates system variables (these are robot inputs and outputs
  sSystem = BuildSystemVariables()
  sReturn = sSystem & sReturn
   
  'creates variables
  g_sErrorNum = "150"
  sReturn = TranslateVariables_VM(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
    
  'remove all blank lines
  g_sErrorNum = "160"
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
  
  'now deal with inner braces..for...else..if..else if..while
  g_sErrorNum = "170"
  sReturn = C_Braces(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
  
  'remove blank lines in array
  g_sErrorNum = "180"
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
   
  'finds all lines of code that require only two lines. I.e. a 'for' w/o braces
  g_sErrorNum = "190"
  sReturn = C_For_Twoliner(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
   
  'commands on only one line are processed here
  g_sErrorNum = "200"
  sReturn = C_For_Oneliner(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
   
  'remove blank lines in array
  g_sErrorNum = "210"
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then TranslationC_VM = sReturn & vbCrLf & g_sErrorNum: Exit Function
      
  'return translated code
  TranslationC_VM = sReturn & "END" & vbCrLf

End Function

'*******************************************
' Gets rid of innerbraces and loops
'*******************************************
Function C_Braces(sInput As String) As String
  Dim uPair As GENERIC_PAIR
  Dim sOut As String
  Dim sLines() As String
  Dim i As Integer
  Dim bFor As Boolean
  Dim bIf As Boolean
  Dim bElse As Boolean
  Dim bElseIf As Boolean
  Dim bWhile As Boolean
  Dim nCt As Integer
  
Repeat:
  uPair = GetInnerBraces(sInput)

  
  'a pair of inner braces have been found
  If uPair.First > 0 And uPair.Second > 0 Then
    
    'create array
    sLines = Split(sInput, vbCrLf)
    
    'for loop
    If VerifySequence(sLines(uPair.First - 1), "for$($;$;$)") = True Then
      sInput = C_Process_For_Loop(sInput, uPair.First - 1, uPair.Second)
    
    'while loop
    'ElseIf VerifySequence(sLines(uPair.First - 1), "while$($)") = True Then
    
    'else if section
    'ElseIf VerifySequence(sLines(uPair.First - 1), "else$if$($)") = True Then
        
    'else section
    'ElseIf VerifySequence(sLines(uPair.First - 1), "else") = True Then
    
    'if section
    ElseIf VerifySequence(sLines(uPair.First - 1), "if$($)") = True Then
      
      
      'If VerifySequence(sLines(uPair.Second + 1), "else$if$($)") = True Then
      
    
      sInput = C_Process_If_Loop(sInput, uPair.First - 1, uPair.Second)
      MsgBox uPair.First & "  " & uPair.Second
    Else
      sInput = sInput
    End If
  End If
  
  sInput = C_For_Oneliner(sInput)
  If Left(sInput, 5) = "ERROR" Then
    C_Braces = sInput
    Exit Function
  End If
  sInput = RemoveBlankLines(sInput)

  'uPair = GetInnerBraces(sInput)
  'If uPair.First > 0 Or uPair.Second > 0 Then GoTo Repeat
  'MsgBox sInput
  nCt = nCt + 1
  If nCt < 3 Then GoTo Repeat
  
  C_Braces = sInput
   
End Function

'*******************************************
' Handles 'if' functions
'*******************************************
Public Function C_Process_If_Loop(sInput As String, ByVal nBegin As Integer, ByVal nEnd As Integer) As String
  Dim i As Integer
  Dim sLines() As String
  Dim sOutFirst As String
  Dim sOutLast As String
  Dim sOut As String
  Dim nA As Integer  'position of key symbols
  Dim nB As Integer
  Dim nC As Integer
  Dim nD As Integer
  Dim sP1 As String  'for parameters
  Dim sP2 As String
  Dim sP3 As String
  Dim sVar As String 'variable name
  Dim sInit As String 'initial value
  Dim sExpr As String 'expression
  Dim sCond As String 'conditional check
  Dim nEqual As Integer
  
  C_Process_For_Loop = sInput
  sLines = Split(sInput, vbCrLf)
  
  'grab code before for loop
  For i = 0 To nBegin - 1
    sOutFirst = sOutFirst & sLines(i) & vbCrLf
  Next i
  
  'grab code after for loop
  For i = nEnd + 1 To UBound(sLines)
    sOutLast = sOutLast & sLines(i) & vbCrLf
  Next i
  
  'let's parse this for code
  nA = InStr(1, sLines(nBegin), "(")
  nB = InStr(nA, sLines(nBegin), ")")
  
  sExpr = Trim(Mid(sLines(nBegin), nA + 1, nB - nA - 1))
  
  For i = nBegin To nEnd - 1
   'MsgBox i & "  " & sLines(i)
  Next i
  
  sOut = sOut & "SVAR " & sVar & "," & sInit & vbCrLf
  sOut = sOut & "GLR " & sCond & "," & CStr((nEnd - nBegin) + 1) & vbCrLf
  sOut = sOut & sExpr & vbCrLf
  For i = nBegin + 2 To nEnd - 1
    sOut = sOut & sLines(i) & vbCrLf
  Next i
  
  sOut = sOut & "JMP -" & CStr((nEnd - nBegin)) & vbCrLf
    
  C_Process_If_Loop = sOutFirst & sOut & sOutLast
End Function

'*******************************************
' Returns start and end of inner most
' braces { and } in an array.
'*******************************************
Public Function GetInnerBraces(sInput As String) As GENERIC_PAIR
  Dim i As Integer
  Dim sLines() As String
  Dim nBraceCt As Integer
  Dim nBrace As Integer 'tracks sequence
  Dim nFirst As Integer
  Dim nSecond As Integer
    
  GetInnerBraces.First = 0
  GetInnerBraces.Second = 0
  If Len(sInput) < 1 Then Exit Function
  nBraceCt = CountChar(sInput, "{")
  If nBraceCt < 1 Then Exit Function
  
  'load string into an array
  sLines = Split(sInput, vbCrLf)
  
  'examine each line for braces
  For i = 0 To UBound(sLines) - 1
    
    If Left(sLines(i), 1) = "{" Then
      'MsgBox i & " " & sLines(i)
      nBrace = nBrace + 1
      nFirst = i
    ElseIf Left(sLines(i), 1) = "}" Then
      'MsgBox i & " " & sLines(i)
      nBrace = nBrace - 1
      nSecond = i
      GetInnerBraces.First = nFirst
      GetInnerBraces.Second = nSecond
      Exit Function
    Else
      'do nothing
    End If
  Next i
  
End Function

'*******************************************
' Returns start and end of inner most
' braces { and } in an array.
'*******************************************
Public Function GetStringInnerBraces(sInput As String) As GENERIC_PAIR
  Dim i As Long
  Dim sLines() As String
  Dim nBraceCt As Integer
  Dim nBrace As Integer 'tracks sequence
  Dim nFirst As Long
  Dim nSecond As Long
    
  GetStringInnerBraces.First = 0
  GetStringInnerBraces.Second = 0
  If Len(sInput) < 1 Then Exit Function
  nBraceCt = CountChar(sInput, "{")
  If nBraceCt < 1 Then Exit Function
  
  
  'examine each line for braces
  For i = 1 To Len(sInput)
    
    If Mid(sInput, i, 1) = "{" Then
      'MsgBox i & " " & sLines(i)
      nBrace = nBrace + 1
      nFirst = i
    ElseIf Mid(sInput, i, 1) = "}" Then
      'MsgBox i & " " & sLines(i)
      nBrace = nBrace - 1
      nSecond = i
      GetStringInnerBraces.First = nFirst
      GetStringInnerBraces.Second = nSecond
      Exit Function
    Else
      'do nothing
    End If
  Next i
  
End Function


'*******************************************
' Return true if entire code is VM
' as opposed to  'C' code
'*******************************************
Public Function IsTotalVMCode(sInput As String) As Boolean
  Dim i As Integer
  Dim sLines() As String
  
  IsTotalVMCode = False
  sInput = RemoveBlankLines(sInput)
  sLines = Split(sInput, vbCrLf)
  For i = 0 To UBound(sLines) - 1
    If IsVMCode(sLines(i)) = False Then
      Exit Function
    End If
  Next i
  IsTotalVMCode = True
End Function
  

'*******************************************
' Return true if line is VM as opposed to
' 'C' code
'*******************************************
Public Function IsVMCode(sInput As String) As Boolean
  IsVMCode = False
  
  'all VM codes need to be added here
  If Left(sInput, 5) = "CVAR " Then
    IsVMCode = True
  ElseIf Left(sInput, 5) = "SVAR " Then
    IsVMCode = True
  ElseIf Left(sInput, 4) = "GLR " Then
    IsVMCode = True
  ElseIf Left(sInput, 4) = "INC " Then
    IsVMCode = True
  ElseIf Left(sInput, 4) = "DEC " Then
    IsVMCode = True
  ElseIf Left(sInput, 4) = "JMP " Then
    IsVMCode = True
  ElseIf Left(sInput, 3) = "END" Then
    IsVMCode = True
  End If
End Function

'*******************************************
' Convert various C standalones to VM code
' Once the C code has been formatted and
' cleaned up then individual lines may
' be considered.  This routine may have
' to be repeated until all braces have
' been eliminated
'*******************************************
Public Function C_For_Oneliner(sInput As String) As String
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
    
    
    'relational test
    If InStr(1, sCode, "==") Then
    
    ElseIf InStr(1, sCode, "Generate_Pwms(pwm13,pwm14,pwm15,pwm16)") Then
      g_bUpperPWM = True
      sLines(i) = ""
    ElseIf InStr(1, sCode, "++") Then  'post-fix operator ++
      
      sReturn = C_IncrementVariable(sCode)
      'MsgBox sOut
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    ElseIf InStr(1, sCode, "--") Then  'post-fix operator --
      sReturn = C_DecrementVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    ElseIf InStr(1, sCode, "*=") Then
      sReturn = C_OperateVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    
    ElseIf InStr(1, sCode, "/=") Then
      sReturn = C_OperateVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    
    ElseIf InStr(1, sCode, "+=") Then
      sReturn = C_OperateVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    
    ElseIf InStr(1, sCode, "-=") Then
      sReturn = C_OperateVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
    
    ElseIf InStr(1, sCode, "=") Then    'assignment operator
      sReturn = C_AssignVariable(sCode)
      If Left(sReturn, 5) = "ERROR" Then
        C_For_Oneliner = sReturn
        Exit Function
      Else
        sLines(i) = sReturn
      End If
           
    End If
  
  Next i
  
  'construct string from array
  For i = 0 To UBound(sLines) - 1
   sOut = sOut & sLines(i) & vbCrLf
  Next i

  C_For_Oneliner = sOut
End Function

'*******************************************
' Performs various operations upon variables
' such as *=, /=, +=, -=
'*******************************************
Public Function C_OperateVariable(sInput As String) As String
  Dim sVar As String
  Dim sValue As String
  Dim nOpPos As Integer
  Dim sOp As String
  Dim sVar2 As String 'other variable
  Dim sReturn As String
  Dim sExpr As String
  
  
  C_OperateVariable = "ERROR" & vbCrLf

  If InStr(1, sInput, "*=") > 0 Then
    nOpPos = InStr(1, sInput, "*=")
    sOp = "MPY"
  ElseIf InStr(1, sInput, "/=") > 0 Then
    nOpPos = InStr(1, sInput, "/=")
    sOp = "DIV"
  ElseIf InStr(1, sInput, "+=") > 0 Then
    nOpPos = InStr(1, sInput, "+=")
    sOp = "ADD"
  ElseIf InStr(1, sInput, "-=") > 0 Then
    nOpPos = InStr(1, sInput, "-=")
    sOp = "SUB"
  Else
  
  End If
  
  If Len(sOp) < 1 Then
    C_OperateVariable = "ERROR" & vbCrLf & "  Unrecognized math operation!" & sInput
    Exit Function
  End If
  
  If nOpPos < 2 Then
    C_OperateVariable = "ERROR" & vbCrLf & "  Missing math operator!" & sInput
    Exit Function
  End If
  
  sVar = Trim(Left(sInput, nOpPos - 1))
  sExpr = Trim(Mid(sInput, nOpPos + 2))
 
  'could be assigning to a variable
  If Not IsNumeric(sExpr) Then
    If VariableExists(sValue) = True Then
      sVar2 = sExpr
      'sValue = GetVariableValue(sValue)
    Else
    End If
    
  End If
  
  'same type
  'If GetVariableType(sVar) = GetVariableType(sVar2) Then
    
  'Else
  
  'End If
    
  If sOp = "MPY" Then
    C_OperateVariable = "SVAR " & sVar & "," & sVar & "*" & sExpr
    Exit Function
  ElseIf sOp = "DIV" Then
    C_OperateVariable = "SVAR " & sVar & "," & sVar & "/" & sExpr
    Exit Function
  ElseIf sOp = "ADD" Then
    C_OperateVariable = "SVAR " & sVar & "," & sVar & "+" & sExpr
    Exit Function
  ElseIf sOp = "SUB" Then
    C_OperateVariable = "SVAR " & sVar & "," & sVar & "-" & sExpr
    Exit Function
  Else
  
  End If
    
  
End Function

'*******************************************
' assigns values to existing variables
'*******************************************
Public Function C_AssignVariable(sInput As String) As String
  Dim sName As String 'var name
  Dim sReturn As String 'catches variable value
  Dim sOut As String
  Dim sVars() As String 'in case there are multiple assignments
  Dim nEqualCt As Integer
  Dim i As Integer
  Dim sValue As String
  
  If InStr(1, sInput, "==") Then
    C_AssignVariable = "ERROR!" & vbCrLf & "   Invalid use of '==' operator in " & sInput
    Exit Function
  End If
  C_AssignVariable = "ERROR!" & vbCrLf & "   Invalid use of '=' operator in " & sInput
  
  'split string , last element in array is the value
  sVars = Split(sInput, "=")
  sValue = sVars(UBound(sVars))
  'MsgBox "Value: " & sValue
  
  'if this value is not a number, it may be a variable...check and see
  If Not IsNumeric(sValue) Then
    sValue = GetVariableValue(sValue) 'get value of variable
    If Left(sValue, 5) = "ERROR" Then
      C_AssignVariable = "ERROR!" & vbCrLf & "   Unknown identifier in " & sVars(i)
      Exit Function
    End If
  End If
  
  'work through assigning all variables
  For i = 0 To UBound(sVars) - 1
    sVars(i) = Trim(sVars(i)) 'nothing here...error
    If Len(sVars(i)) < 1 Then
      C_AssignVariable = "ERROR!" & vbCrLf & "   Missing identifier in " & sVars(i)
      Exit Function
    End If

    'it should exist
    If VariableExists(sVars(i)) = True Then
      sReturn = SetVariable(sVars(i), sValue)
      If Left(sReturn, 5) = "ERROR" Then
        C_AssignVariable = "ERROR!" & vbCrLf & "   Could not assign value to identifier. " & sVars(i)
        Exit Function
      End If
      sOut = sOut & "SVAR " & sVars(i) & "," & sValue & vbCrLf
    Else
      MsgBox sVars(i)
      C_AssignVariable = "ERROR!" & vbCrLf & "   Identifier doesn't exist " & sVars(i)
      Exit Function
    End If
  Next i
  
C_AssignVariable = sOut

End Function


'*******************************************
' robot controller I/O variables, static
'*******************************************
Public Function BuildSystemVariables() As String
  Dim sOut As String
    
  'sOut = "ERROR!" & vbCrLf & "   Unknown!"
  
  'adds to variable table
  sOut = sOut & AddVariable("pwm01", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm02", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm03", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm04", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm05", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm06", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm07", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm08", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm09", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm10", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm11", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm12", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm13", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm14", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm15", "unsigned char", "127", "static")
  sOut = sOut & AddVariable("pwm16", "unsigned char", "127", "static")
  
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
  sOut = sOut & AddVariable("relay5_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay6_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay7_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay8_fwd", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay5_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay6_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay7_rev", "unsigned char", "0", "static")
  sOut = sOut & AddVariable("relay8_rev", "unsigned char", "0", "static")

  
  'constructs VM Code
  sOut = sOut & "CVAR pwm01,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm02,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm03,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm04,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm05,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm06,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm07,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm08,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm09,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm10,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm11,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm12,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm13,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm14,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm15,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR pwm16,unsigned char,127" & vbCrLf
  
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
  
  sOut = sOut & "CVAR relay5_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay5_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay6_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay6_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay7_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay7_rev,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay8_fwd,unsigned char,127" & vbCrLf
  sOut = sOut & "CVAR relay8_rev,unsigned char,127" & vbCrLf
  
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
  Dim sOut As String
  
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
' Processes a 'for' loop with braces
'*******************************************
Public Function C_Process_For_Loop(sInput As String, ByVal nBegin As Integer, ByVal nEnd As Integer) As String
  Dim i As Integer
  Dim sLines() As String
  Dim sOutFirst As String
  Dim sOutLast As String
  Dim sOut As String
  Dim nA As Integer  'position of key symbols
  Dim nB As Integer
  Dim nC As Integer
  Dim nD As Integer
  Dim sP1 As String  'for parameters
  Dim sP2 As String
  Dim sP3 As String
  Dim sVar As String 'variable name
  Dim sInit As String 'initial value
  Dim sExpr As String 'expression
  Dim sCond As String 'conditional check
  Dim nEqual As Integer
  
  C_Process_For_Loop = sInput
  sLines = Split(sInput, vbCrLf)
  
  'grab code before for loop
  For i = 0 To nBegin - 1
    sOutFirst = sOutFirst & sLines(i) & vbCrLf
  Next i
  
  'grab code after for loop
  For i = nEnd + 1 To UBound(sLines)
    sOutLast = sOutLast & sLines(i) & vbCrLf
  Next i
  
  'let's parse this for code
  nA = InStr(1, sLines(nBegin), "(")
  nB = InStr(nA, sLines(nBegin), ";")
  nC = InStr(nB + 1, sLines(nBegin), ";")
  nD = InStr(nC, sLines(nBegin), ")")
  sP1 = Trim(Mid(sLines(nBegin), nA + 1, nB - nA - 1))
  sP2 = Trim(Mid(sLines(nBegin), nB + 1, nC - nB - 1))
  sP3 = Trim(Mid(sLines(nBegin), nC + 1, nD - nC - 1))
  
  'get initial variable and value
  nEqual = InStr(1, sP1, "=")
  If nEqual > 0 Then
    sVar = Left(sP1, nEqual - 1)
    sInit = Mid(sP1, nEqual + 1)
  Else
    
  End If
  
  'get conditional check
  sCond = sP2
  
  'get expression
  sExpr = sP3
 
  For i = nBegin To nEnd - 1
   'MsgBox i & "  " & sLines(i)
  Next i
  
  sOut = sOut & "SVAR " & sVar & "," & sInit & vbCrLf
  sOut = sOut & "GLR " & sCond & "," & CStr((nEnd - nBegin) + 1) & vbCrLf
  sOut = sOut & sExpr & vbCrLf
  For i = nBegin + 2 To nEnd - 1
    sOut = sOut & sLines(i) & vbCrLf
  Next i
  
  sOut = sOut & "JMP -" & CStr((nEnd - nBegin)) & vbCrLf
    
  C_Process_For_Loop = sOutFirst & sOut & sOutLast
End Function
'*******************************************
' Finds all 'for' functions with only one
' line of code (no braces). Converts to
' VM
'*******************************************
Public Function C_For_Twoliner(sInput As String) As String
  Dim i As Integer
  Dim nCount As Integer
  Dim sOut As String
  Dim sLines() As String
  Dim nA As Integer  'position of key symbols
  Dim nB As Integer
  Dim nC As Integer
  Dim nD As Integer
  Dim sP1 As String  'for parameters
  Dim sP2 As String
  Dim sP3 As String
  Dim sVar As String 'variable name
  Dim sInit As String 'initial value
  Dim sExpr As String 'expression
  Dim sCond As String 'conditional check
  Dim nEqual As Integer
  Dim bReturn As Boolean
  Dim bFound As Boolean
    
  nCount = InStr(1, sInput, "for") 'look for 'for' words
  If nCount < 1 Then  'no found in this code
    C_For_Twoliner = sInput
    Exit Function
  End If
  
  'something found, lets examine
  sLines = Split(sInput, vbCrLf)
  
  For i = 0 To UBound(sLines) - 1
    If InStr(1, sLines(i), "for") > 0 Then 'for found
      bReturn = VerifySequence(sLines(i), "for$($;$;$)")
      'MsgBox sLines(i) & " " & bReturn
      If bReturn = True Then
        If Left(sLines(i + 1), 1) = "{" Then
          'not a two liner
          sOut = sOut & sLines(i) & vbCrLf
        Else  'two liner - lets parse
          bFound = True
          'let's parse this for code
          nA = InStr(1, sLines(i), "(")
          nB = InStr(nA, sLines(i), ";")
          nC = InStr(nB + 1, sLines(i), ";")
          nD = InStr(nC, sLines(i), ")")
          sP1 = Trim(Mid(sLines(i), nA + 1, nB - nA - 1))
          sP2 = Trim(Mid(sLines(i), nB + 1, nC - nB - 1))
          sP3 = Trim(Mid(sLines(i), nC + 1, nD - nC - 1))
          
          'get initial variable and value
          nEqual = InStr(1, sP1, "=")
          If nEqual > 0 Then
            sVar = Left(sP1, nEqual - 1)
            sInit = Mid(sP1, nEqual + 1)
          Else
            
          End If
          
          'get conditional check
          sCond = sP2
          
          'get expression
          sExpr = sP3
          
          sOut = sOut & "SVAR " & sVar & "," & sInit & vbCrLf
          sOut = sOut & "GLR " & sCond & ",4" & vbCrLf
          sOut = sOut & sExpr & vbCrLf
        End If
      
      Else
        sOut = sOut & sLines(i + 1) & vbCrLf
      End If
      sOut = sOut & sLines(i + 1) & vbCrLf & "JMP -3" & vbCrLf
      
    Else 'not a for loop
      If bFound = True Then
        bFound = False
      Else
        sOut = sOut & sLines(i) & vbCrLf
      End If
      'MsgBox sOut
    End If
  
  Next i
 
  C_For_Twoliner = sOut ' & "JMP  3" & vbCrLf
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
'set variable value
'*******************************************
Public Function SetVariableValue(sIn As String, sValue As String) As String
  Dim i As Integer
    
  SetVariableValue = "ERROR"   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Name = Trim(sIn) Then
      g_uVar(i).Value = sValue
      SetVariableValue = "OK"
      Exit Function
    End If
  Next i
End Function


'*******************************************
'gets variable type
'*******************************************
Public Function GetVariableType(sIn As String) As String
  Dim i As Integer
    
  GetVariableType = "ERROR"   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Symbol = Trim(sIn) Then
      GetVariableType = g_uVar(i).Type
      Exit Function
    End If
  Next i

End Function

'*******************************************
' Returns true if variable exists
'*******************************************
Public Function VariableExists(sIn As String) As Boolean
  Dim i As Integer
    
  VariableExists = False   'default doesn't exist
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Name = Trim(sIn) Then
      VariableExists = True
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
          TranslateVariables_VM = "ERROR!" & vbCrLf & "   Missing semi colon..." & vbCrLf
          TranslateVariables_VM = TranslateVariables_VM & sOut
          Exit Function
        Else
          sLines(i) = Left(sLines(i), Len(sLines(i)) - 1)
        End If
        
        bFound = True
        sVMString = "," & Trim(sVar(j)) & ","
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
            TranslateVariables_VM = "ERROR!" & vbCrLf & "  No value assigned in variable declaration..." & vbCrLf
            TranslateVariables_VM = TranslateVariables_VM & sOut
            Exit Function
          End If
          sVMString = "CVAR " & Trim(sName) & Trim(sVMString) & sValue
        Else 'variable created, not intialized
          sName = Trim(sTemp)
          
          If Len(sName) < 1 Then
            TranslateVariables_VM = "ERROR!" & vbCrLf & "  Incorrect variable declaration..." & vbCrLf
            TranslateVariables_VM = TranslateVariables_VM & sOut
            Exit Function
          End If
          
          'constructs VM code
          sVMString = "CVAR " & Trim(sName) & Trim(sVMString) & "0"
          
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
' Looks for several lines of code on
' the same line delimited by semicolons
'******************************************
Public Function AlignSemicolons(sInput As String) As String
  Dim sOut As String
  Dim i, j As Integer
  Dim sLines() As String
  Dim nCount As Integer
  Dim nPos As Integer
  Dim sSemi() As String
  
  sOut = ""
  
  sLines = Split(sInput, vbCrLf)
  
  'exclude blank rows and trim the rest
  For i = 0 To UBound(sLines) - 1
    nCount = CountChar(sLines(i), ";")
    
    'more than one semicolon
    If nCount > 1 Then
      nPos = 0
      nPos = InStr(1, sLines(i), "for")
      If nPos < 1 Then  'no 'for' here...lets split these up
        sSemi = Split(sLines(i), ";")
        For j = 0 To UBound(sSemi)
          sOut = sOut & sSemi(j) & ";" & vbCrLf
        Next j
      Else  'assume for...keep all
        sOut = sOut & sLines(i) & vbCrLf
      End If
    Else
      sOut = sOut & sLines(i) & vbCrLf
    End If
  Next i
  
  AlignSemicolons = sOut
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
    'ElseIf Mid(sInput, i, 1) = ";" Then
    '  sOut = sOut & ";" & vbCrLf
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
        IsAutoFunction = "ERROR!" & vbCrLf & "  Incorrect placement of braces '}'"
        Exit Function
      End If
    End If
  Next i
  
  'braces must be matched
  If nBrace > 0 Then
    IsAutoFunction = "ERROR!" & vbCrLf & "  Incorrect number of braces '{' or '}'"
    Exit Function
  End If
  
  If nBraceCt < 3 Then
    IsAutoFunction = "ERROR!" & vbCrLf & "  Insufficient number of braces '{' for autonomous mode."
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
    IsAutoFunction = "ERROR!" & vbCrLf & "  Essential autonomous lines are missing or out of order." & vbCrLf
    IsAutoFunction = IsAutoFunction & "  Verify the following lines are include in this sequence." & vbCrLf
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
    RemoveBlankLines = "ERROR!" & vbCrLf & "  Nothing to compile"
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
 
  Dim nLength As Long 'length of file
  Dim uPair As GENERIC_PAIR
  Dim sTemp As String 'temp file storage
  Dim bTryAgain As Boolean 'use for repeating code
  Dim nPos As Long 'position of any character
  Dim nLBrace As Long 'position braces
  Dim nRBrace As Long '
  Dim nPos2 As Long 'another position marker in the string
  Dim nPos3 As Long
    
  'let's find and removecomments /*  and */
  '*****************************************
  Do
    uPair = FindMatchingPair(1, "/*", "*/", sInput)
    
    If uPair.First > 0 And uPair.Second > 0 Then
      bTryAgain = True
      sTemp = ""
      sTemp = Left(sInput, uPair.First - 1) & vbCrLf
      sTemp = sTemp & Mid(sInput, uPair.Second + 2)
      sInput = sTemp
    Else
      bTryAgain = False
    End If
  Loop Until bTryAgain = False
  
  'let's find and remove comments //
  '*********************************
  Do
    uPair = FindMatchingPair(1, "//", vbCrLf, sInput)
    
    If uPair.First > 0 And uPair.Second > 0 Then
      bTryAgain = True
      sTemp = ""
      sTemp = Left(sInput, uPair.First - 1)
      sTemp = sTemp & Mid(sInput, uPair.Second + 2)
      sInput = sTemp
    Else
      bTryAgain = False
    End If
  Loop Until bTryAgain = False
  
  RemoveComments = sInput
End Function

'****************************************
' Loads global variables for this
' module TranslatorC.Bas
'****************************************
Public Sub LoadTranslatorCVariables()
  g_sTranslatorC_Version = "0.03"
  g_bUpperPWM = False
End Sub

'****************************************
'clears all variables from var() array
' except 'static' variables
'****************************************
Public Sub ClearVariables()
  Dim i As Integer
  
  For i = 0 To MAX_VARIABLES
    If g_uVar(i).Scope <> "static" Then
      g_uVar(i).Name = ""
      g_uVar(i).Type = ""
      g_uVar(i).Value = ""
    End If
  Next i
End Sub


