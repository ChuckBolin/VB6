Attribute VB_Name = "Errors"
'********************************************************************************
' ERRORS.BAS
' Provides for error handling and datalogging of such errors.
' Modifed July 13, 2002. Replaced gsLocation with gsLoc (shorter
'********************************************************************************

Option Explicit

'all errors provide this info that can be saved for debugging and/or BETA testing
Public glErrorNumber As Integer 'VB error number...generally helpful but often not specific enough to debug quickly
Public gsErrorDescription As String 'VB error description...often lacks sufficient detail for rapid debugging
Public gsForm As String 'form or module generating error
Public gsProcedure As String 'states procedure with occuring error
Public gsLoc As String 'used to insert locations in code for quick reference
Public gsInfo As String 'use this to store any info that may be relevent such as variables

Public Sub ErrorHandler()
  Dim sMsg As String
  Dim nFile As Integer
  Dim sVar As String 'stores all global variables
  
  'collect error specific info
  glErrorNumber = Err.Number
  gsErrorDescription = Err.Description
  
  'construct error message
  sMsg = " E R R O R ! ! !" & vbCrLf
  sMsg = sMsg & "No.: " & CStr(glErrorNumber) & vbCrLf
  sMsg = sMsg & "Description: " & gsErrorDescription & vbCrLf
  sMsg = sMsg & "Form: " & gsForm & vbCrLf
  sMsg = sMsg & "Procedure: " & gsProcedure & vbCrLf
  sMsg = sMsg & "Location in Code: " & gsLoc & vbCrLf
  sMsg = sMsg & "Info: " & gsInfo
  MsgBox sMsg
   
  FetchVariables sVar
  
  'save error information to file
  nFile = FreeFile 'get next available file handle
  Open App.Path & "\errorlog.txt" For Append As #nFile
    Print #nFile, " "
    Print #nFile, "********************************************************"
    Print #nFile, sMsg
    Print #nFile, " DATE: " & CStr(Date) & "  TIME: " & CStr(Time)
    Print #nFile, "********************************************************"
    Print #nFile, "        G L O B A L   V A R I A B L E S"
    Print #nFile, "********************************************************"
    Print #nFile, sVar
    Print #nFile, "********************************************************"
    Print #nFile, " "
  Close #nFile
    
  'clear error variables for reuse
  glErrorNumber = 0
  gsErrorDescription = "Not provided!"
  gsForm = "Not provided!"
  gsProcedure = "Not provided!"
  gsLoc = "Not provided!"
  gsInfo = "Not provided!"
    
End Sub

'collects all global variables and stores them into single
'string sVar for output to file
Public Sub FetchVariables(sVar As String)
  Dim x As Integer
  
  sVar = ""
  
End Sub
