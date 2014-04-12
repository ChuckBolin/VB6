Attribute VB_Name = "GenSubs"
Option Explicit

'*********************************************************
'REMOVE WHITESPACE
'Removes spaces and tabs from strings
'*********************************************************
Public Sub RemoveWhitespaces(sString As String)
  On Error GoTo MyError
  Dim x As Integer
  Dim sTemp As String
  Dim sChar As String
    
  'exit sub if length of string <1
  If Len(sString) < 1 Then Exit Sub
  
  'extract white spaces and place results in sTemp
  For x = 1 To Len(sString)
    sChar = Mid(sString, x, 1)
    If sChar = vbKeySpace Or sChar = vbKeyTab Then 'do not keep character
    Else 'keep this character
      sTemp = sTemp & sChar
    End If
  Next x

  'swap string data
  sString = sTemp
  Exit Sub

MyError:
  gsForm = "GenSubs"
  gsProcedure = "RemoveWhitespaces"
  ErrorHandler
End Sub
