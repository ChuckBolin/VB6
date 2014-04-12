Attribute VB_Name = "Main"
Option Explicit

'*****************************************************************************
'MAIN
'Starts program
'*****************************************************************************
Public Sub Main()
  Initialize
  frmMain.Show
End Sub

'*****************************************************************************
'INITIALIZE
'Sets all variables required to start program
'*****************************************************************************
Public Sub Initialize()
  On Error GoTo MyError
  Dim bFileFound As Boolean
  
  'fixed variable data
  gsVersion = "v0.1d"
  gsVersionDate = "June 27, 2002"
  ReDim uABS(0)
  
  gnABSTotal = 0
  
  'load parameters from file
  LoadFile "learning.txt", bFileFound
  
  'load default values if file not found
  If bFileFound = False Then
    gsRules = "" 'stores all rules collected during program
    gnRows = 3  'defines 4 x 4 board
    gnCols = 3
    gsProgramSymbol = "X"
    gnProgramValue = 1
    gsTeacherSymbol = "O"
    gnTeacherValue = 2
    gnGoFirst = 1
    gnGameType = 1
    glCellSelectedColor = 65280
  End If
  
  'variable dependent upon above
  gnTotalCells = gnRows * gnCols
  
  'who goes first must go now
  If gnGoFirst = 1 Then
    gbProgramTurn = False
  Else
    gbProgramTurn = True
  End If
'  MsgBox uABS(1).word
'  MsgBox uABS(2).word
'  MsgBox uABS(3).word
  
  Exit Sub
MyError:
  gsForm = "Subs"
  gsProcedure = "Initialize"
  ErrorHandler
End Sub

