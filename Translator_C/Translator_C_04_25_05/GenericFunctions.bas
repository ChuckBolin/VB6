Attribute VB_Name = "GenericFunctions"
'****************************************************************************
' GenericFunctions.bas - Written by Chuck Bolin, April 2005
' Stores functions and subs that can be used in other
' programs and does not point to any specific global variables or is related
' to this program. Can be used without conversion in other programs.
' Functions:
' FindMatchingPair(nBegin, sFirst, sSecond , sInput) returns string pointers
'                  to a matching set of character sFirst and sSecond
' VerifySequence(sInput, sSeq) returns true if all characters are in
'    correct sequence
' IsAlpha(sIn) return true if all characters are letters
' CountChar(sIn, sChar ) returns number of specified characters
' IsSymbol (sChar) returns boolean true if char is not alpha-numeric
' GetFileContents (sFilename) returns contents of text file
'****************************************************************************
Option Explicit

Public Type GENERIC_PAIR
  First As Long
  Second As Long
  InBetween As Boolean 'for use with matching pair
End Type

'****************************************************
' Checks for the occurence of a specific character
' not counting vbCrLf, spaces or vbtabs at a
' certain position within a string
'****************************************************
Public Function VerifyNextCharacter(nBegin As Long, sChar As String, sInput As String) As Boolean
  Dim i As Long
  VerifyNextCharacter = False
  If Len(sInput) <= nBegin Then Exit Function
  If Len(sChar) <> 1 Then Exit Function
  
  For i = nBegin To Len(sInput)
    If Mid(sInput, i, 1) = " " Then
      'ignore spaces
    ElseIf Mid(sInput, i, 1) = vbTab Then
      'ignore tabs
    ElseIf Mid(sInput, i, 1) = vbCr Then
      'ignore carriage returns
    ElseIf Mid(sInput, i, 1) = vbLf Then
      'ignore line feeds
    ElseIf Mid(sInput, i, 1) = sChar Then 'found
      VerifyNextCharacter = True
      Exit Function
    Else 'not found
      Exit Function
    End If
  Next i
  
End Function


'****************************************************
' This returns the position of the first and second
' character that is matching.  I.e. {},[], (), <>
' /* */. Returns a 0 if not found. Helps to find
' matched pairs of characters.
'****************************************************
Public Function FindMatchingPair(nBegin As Long, sFirst As String, sSecond As String, sInput As String) As GENERIC_PAIR
  Dim i As Long
  Dim nCount As Integer
  Dim nPair As GENERIC_PAIR
  
  FindMatchingPair = nPair
  
  sFirst = Trim(sFirst): sSecond = Trim(sSecond)
  If Len(sFirst) < 1 Or Len(sSecond) < 1 Then Exit Function
  If nBegin > Len(sInput) Then Exit Function
  
  For i = nBegin To Len(sInput)
    If Mid(sInput, i, Len(sFirst)) = sFirst Then
      nCount = nCount + 1
      
      If nCount = 1 Then
        nPair.First = i
      End If
    ElseIf Mid(sInput, i, Len(sSecond)) = sSecond And nCount > 0 Then
      nCount = nCount - 1
      If nCount = 0 Then
        nPair.Second = i
        FindMatchingPair = nPair
        Exit Function
      Else
        nPair.InBetween = True 'this means there are similar characters in between
      End If
    End If
    
  Next i
  FindMatchingPair = nPair
End Function


'****************************************************
' determines correct sequence of characters in
' a string.  sSeq is a string literal that is '$'
' delimited.  For example:
'  bReturn = VerifySequence(sLines(i), "for$($;$;$)")
'****************************************************
Public Function VerifySequence(sInput As String, sSeq As String) As Boolean
  Dim i As Integer
  Dim sToken() As String 'stores sequential stuff
  Dim nPos1 As Integer 'position of items
  Dim nPos2 As Integer
  VerifySequence = False
  
  sToken = Split(sSeq, "$")
  
  nPos1 = 1
  For i = 0 To UBound(sToken) - 1
    nPos2 = InStr(nPos1, sInput, sToken(i))
    If nPos2 < nPos1 Then
      VerifySequence = False
      Exit Function
    End If
    nPos1 = nPos2
  Next i
  
  VerifySequence = True
End Function

'****************************************************
' determines if a charaacter is alpha
'****************************************************
Public Function IsAlpha(sChar As String) As Boolean
  IsAlpha = False
  Dim i As Integer
  
  If Len(sChar) < 1 Then Exit Function
    For i = 1 To Len(sChar)
      If Asc(UCase(Mid(sChar, i, 1))) < 65 Or Asc(UCase(Mid(sChar, i, 1))) > 90 Then
        IsAlpha = False
        Exit Function
      End If
    End If
  Next i
End Function

'counts the number of a certain character
Public Function CountChar(ByVal sIn As String, ByVal sChar As String) As Integer
  Dim i As Integer
  
  CountChar = 0
  If Len(sIn) < 1 Then Exit Function
  
  'iterate through string and look for and count sChar
  For i = 1 To Len(sIn)
    If Mid(sIn, i, Len(sChar)) = sChar Then
      CountChar = CountChar + 1
    End If
  Next i
    
End Function

'****************************************************
'returns true if character is not alphanumeric
Public Function IsSymbol(sChar As String) As Boolean
  IsSymbol = False
  
  If Len(sChar) <> 1 Then Exit Function
  
  If IsNumeric(sChar) = True Then Exit Function
  
  If (Asc(sChar) >= 65 And Asc(sChar) <= 90) Or (Asc(sChar) >= 97 And Asc(sChar) <= 122) Then
    IsSymbol = False
  Else
    IsSymbol = True
  End If
  
End Function

'Returns the contents of a text ("C") file as a string
Public Function GetFileContents(sFile As String) As String
  Dim nFile As Integer
  Dim sInput As String
  Dim sOut As String
    
  nFile = FreeFile
  
  If Dir(sFile) = "" Then
    GetFileContents = "Bad File Name: " & sFile
  End If
  
  Open sFile For Input As nFile
    Do
      Line Input #nFile, sInput
      sOut = sOut & sInput & vbCrLf
    Loop Until EOF(nFile)
  Close nFile
  
  GetFileContents = sOut
End Function

