Attribute VB_Name = "MainProg"
Option Explicit

'program begins here
Public Sub Main()
  Initialize
  frmMain.Show
End Sub

'***********************************************
'INITIALIZE
'Preloads program with default attributes
'***********************************************
Public Sub Initialize()
  On Error GoTo MyError
  
  'seed randomizer
  Randomize Timer
  
  'let's grab default global information
  LoadDefaultData
  
  Exit Sub
MyError:
  gsForm = "MainProg"
  gsProcedure = "Initialize"
  ErrorHandler

End Sub


'***********************************************
'LOAD DEFAULT DATA
'Preloads program with default data
'***********************************************
Public Sub LoadDefaultData()
  On Error GoTo MyError
            
  'system variables
  gsVersion = "0.4a"
  gsProgramName = "Artificial Intelligence Program (AIP)"
  
  'this defines default grid size and look
  gnRows = 15
  gnCols = 15
  gnTotalCells = gnRows * gnCols
  glCellColor = RGB(255, 255, 255)
  glCellColorInverse = RGB(255, 0, 0)
  glCellSelectedColor = RGB(0, 255, 0)
  glGridLeft = 1000
  glGridTop = 1000
  glCellHeight = 500
  glCellWidth = 500
  gbGridReferenceOn = True
  gbGridCheckerBoardOn = False
  gnGridCheckerBoardType = 1  'this is standard checkers/chess board
  gbGridRandomPatternOn = True
  gnGridRandomPatternNum = 50
  gbGridVisible = True

  Exit Sub
MyError:
  gsForm = "MainProgram"
  gsProcedure = "LoadDefaultData"
  ErrorHandler
End Sub
