Attribute VB_Name = "Errors"
Option Explicit

'all errors provide this info that can be saved for debugging and/or BETA testing
Public glErrorNumber As Integer 'VB error number...generally helpful but often not specific enough to debug quickly
Public gsErrorDescription As String 'VB error description...often lacks sufficient detail for rapid debugging
Public gsForm As String 'form or module generating error
Public gsProcedure As String 'states procedure with occuring error
Public gsLocation As String 'used to insert locations in code for quick reference
Public gsInfo As String 'use this to store any info that may be relevent such as variables

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
  sMsg = sMsg & "Location in Code: " & gsLocation & vbCrLf
  sMsg = sMsg & "Info: " & gsInfo
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
  gsInfo = "Not provided!"
    
End Sub
