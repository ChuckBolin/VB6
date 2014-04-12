Attribute VB_Name = "GenericFunctions"
'****************************************************************************
' GenericFunctions.bas - Written by Chuck Bolin, April 2005
' Stores functions and subs that can be used in other
' programs and does not point to any specific global variables or is related
' to this program. Can be used without conversion in other programs.
' Functions:
' CountChar(sIn, sChar ) returns number of specified characters
' IsSymbol (sChar) returns boolean true if char is not alpha-numeric
' GetFileContents (sFilename) returns contents of text file
'****************************************************************************
Option Explicit

'****************************************************
' determines if a charaacter is alpha
'****************************************************
Public Function IsAlpha(sChar As String) As Boolean
  IsAlpha = False
  If Len(sChar) < 1 Then Exit Function
  If Len(sChar) > 1 Then sChar = Mid(sChar, 1, 1)
  If (Asc(sChar) >= 65 And Asc(sChar) <= 90) Or (Asc(sChar) >= 97 And Asc(sChar) <= 122) Then
    IsAlpha = True
  End If
End Function

'counts the number of a certain character
Public Function CountChar(sIn As String, sChar As String) As Integer
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

