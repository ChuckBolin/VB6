Attribute VB_Name = "error"
Option Explicit

'**************************************
' E R R O R H A N D L E R
' Written by Chuck Bolin, March 1, 2000
' Requires System.bas
'**************************************
Public Sub ErrorHandler()
  Dim strPhrase As String
  Dim intResult As Integer
  
  'this routine requires System.bas mod
  If zblnRecordEventOn = False Then Exit Sub
  
  'specific action depending upon error numbers
  If Err.Number = 0 Then Exit Sub
  
  'display error message
  strPhrase = "Error has occurred: " _
    & vbCrLf & "Error Number: " & Err.Number _
    & vbCrLf & "Description: " & Err.Description _
    & vbCrLf & "Location in code: " & zstrFlag _
    & vbCrLf _
    & vbCrLf & "Send Eventlog.dat file to Programmer for Debug"
  
  PushError Date, Time
  
  intResult = MsgBox(strPhrase, vbAbortRetryIgnore, "Illegal Operation")
  Select Case intResult
    Case vbAbort:
      PushEvent "Error Abort", Time
      End
    Case vbRetry:
      PushEvent "Error Retry", Time
      'Resume Next
    Case vbIgnore:
      PushEvent "Error Ignore", Time
      'Resume
  End Select
End Sub

'**************************
' C R E A T E R R O R
'**************************
Public Sub CreateError(lngError As Long)

  On Error GoTo MyError
  PushEvent "ErrorMod - CreateError", Time
  zstrFlag = "EM_CE_" & CStr(lngError)
    
  Err.Raise lngError
  Exit Sub
  
MyError:
  ErrorHandler
End Sub
