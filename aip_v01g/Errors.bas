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
  sMsg = sMsg & "Location in Code: " & gsLocation & vbCrLf
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
  gsLocation = "Not provided!"
  gsInfo = "Not provided!"
    
End Sub

'collects all global variables and stores them into single
'string sVar for output to file
Public Sub FetchVariables(sVar As String)
  Dim x As Integer
  
  sVar = ""
  sVar = sVar & "gbFilenameExists: " & CStr(gbFilenameExists) & vbCrLf
  sVar = sVar & "gbWin: " & CStr(gbWin) & vbCrLf
  sVar = sVar & "gbProgramTurn: " & CStr(gbProgramTurn) & vbCrLf
  sVar = sVar & "gbGameOver: " & CStr(gbGameOver) & vbCrLf
  sVar = sVar & "gbLoss: " & CStr(gbLoss) & vbCrLf
  sVar = sVar & "gbTie: " & CStr(gbTie) & vbCrLf
  sVar = sVar & "gbTeacherDone: " & CStr(gbTeacherDone) & vbCrLf
  sVar = sVar & "gnGoFirst: " & CStr(gnGoFirst) & vbCrLf
  sVar = sVar & "gnRows: " & CStr(gnRows) & vbCrLf
  sVar = sVar & "gnCols: " & CStr(gnCols) & vbCrLf
  sVar = sVar & "gnProgramValue: " & CStr(gnProgramValue) & vbCrLf
  sVar = sVar & "gnTeacherValue: " & CStr(gnTeacherValue) & vbCrLf
  sVar = sVar & "gnProgramWins: " & CStr(gnProgramWins) & vbCrLf
  sVar = sVar & "gnProgramLosses: " & CStr(gnProgramLosses) & vbCrLf
  sVar = sVar & "gnProgramTies: " & CStr(gnProgramTies) & vbCrLf
  sVar = sVar & "gnTotalCells: " & CStr(gnTotalCells) & vbCrLf
  sVar = sVar & "gnGameType: " & CStr(gnGameType) & vbCrLf
  sVar = sVar & "gnMoveCount: " & CStr(gnMoveCount) & vbCrLf
  sVar = sVar & "gnGameCount: " & CStr(gnGameCount) & vbCrLf
  sVar = sVar & "gnPlayCount: " & CStr(gnPlayCount) & vbCrLf
  sVar = sVar & "glCellColor: " & CStr(glCellColor) & vbCrLf
  sVar = sVar & "glCellSelectedColor: " & CStr(glCellSelectedColor) & vbCrLf
  sVar = sVar & "gsRules: " & CStr(gsRules) & vbCrLf
  sVar = sVar & "gsVersion: " & CStr(gsVersion) & vbCrLf
  sVar = sVar & "gsProgramSymbol: " & CStr(gsProgramSymbol) & vbCrLf
  sVar = sVar & "gsTeacherSymbol: " & CStr(gsTeacherSymbol) & vbCrLf
  sVar = sVar & "gsFilename: " & CStr(gsFilename) & vbCrLf
  sVar = sVar & "gsGameName: " & CStr(gsGameName) & vbCrLf
  sVar = sVar & "gnABSTotal: " & CStr(gnABSTotal) & vbCrLf
  sVar = sVar & "glTeacher: " & CStr(glTeacher) & vbCrLf
  sVar = sVar & "glProgram: " & CStr(glProgram) & vbCrLf
  sVar = sVar & "glAllCells: " & CStr(glAllCells) & vbCrLf
  sVar = sVar & "glFreeCells: " & CStr(glFreeCells) & vbCrLf
  sVar = sVar & "gbWinNotSaved: " & CStr(gbWinNotSaved) & vbCrLf
  
  'extracts winning patterns (long integer form) and matching wins
  For x = 0 To gnABSTotal
    sVar = sVar & "uABS( " & CStr(x) & ") Pattern : " & CStr(uABS(x).word) & vbCrLf
    sVar = sVar & "uABS( " & CStr(x) & ") Wins : " & CStr(uABS(x).wins) & vbCrLf
  Next x

End Sub
