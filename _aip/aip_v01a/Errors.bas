Attribute VB_Name = "Errors"
Option Explicit

'all errors provide this info that can be saved for debugging and/or BETA testing
Public glErrorNumber As Integer
Public gsErrorDescription As String
Public gsForm As String
Public gsProcedure As String
Public gsLocation As String

Public Sub ErrorHandler()
  Dim sMsg As String
  Dim nFile As Integer
  
  'collect error specific info
  glErrorNumber = Err.Number
  gsErrorDescription = Err.Description
  
  'construct error message
  sMsg = "E R R O R ! ! !" & vbCrLf
  sMsg = sMsg & "No.: " & CStr(glErrorNumber) & vbCrLf
  sMsg = sMsg & "Description: " & gsErrorDescription & vbCrLf
  sMsg = sMsg & "Form: " & gsForm & vbCrLf
  sMsg = sMsg & "Procedure: " & gsProcedure & vbCrLf
  sMsg = sMsg & "Location in Code: " & gsLocation
  MsgBox sMsg
  
  'save error information to file
  nFile = FreeFile 'get next available file handle
  Open App.Path & "\errorlog.txt" For Append As #nFile
  Print #nFile, "********************************************************"
  Print #nFile, "DATE: " & CStr(Date) & "   TIME: " & CStr(Time)
  Print #nFile, sMsg
  Print #nFile, "*********************************************************"
  Close #nFile
    
  'clear error variables for reuse
  glErrorNumber = 0
  gsErrorDescription = "Not provided!"
  gsForm = "Not provided!"
  gsProcedure = "Not provided!"
  gsLocation = "Not provided!"
    
End Sub
