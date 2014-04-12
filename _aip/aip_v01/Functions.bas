Attribute VB_Name = "Functions"
Option Explicit

'*****************************************************************************
'COUNTCHAR( )
'counts the number of occurences of a particular
'character found in a string and returns the value
'*****************************************************************************
Public Function CountChar(sIn As String, sChar As String) As Integer
  Dim x As Integer
  Dim nCount As Integer
  On Error GoTo MyError
  
  If Len(sIn) < 1 Then Exit Function
  If Len(sChar) < 1 Then Exit Function
  sIn = UCase(sIn)
  sChar = UCase(sChar)
  For x = 1 To Len(sIn)
    If Mid(sIn, x, 1) = sChar Then nCount = nCount + 1
  Next x
  CountChar = nCount
  Exit Function
  
MyError:
  gsForm = "Module Functions"
  gsProcedure = "CountChar"
  ErrorHandler
End Function

