Attribute VB_Name = "MainSubs"
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
  Dim bFileFound As Boolean
  Dim sFileName As String
  
  On Error GoTo MyError
  
  'load variable data
  gsVersion = "v0.1g"
  gsVersionDate = "July 7, 2002"
  ReDim uABS(0) As Pattern
  gnABSTotal = 0
  gbWinNotSaved = False
  
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
  
  
  Exit Sub
MyError:
  gsForm = "MainSubs"
  gsProcedure = "Initialize"
  ErrorHandler
End Sub

'*************************************************************************
'PLAYCOORDINATOR( )
'Decides whose turn it is and direct the play from there.
'*************************************************************************
Public Sub PlayCoordinator()
  Dim nProgramMove As Integer
  Dim sMsg As String
  Dim nPlay As Integer
  
  'main game play
  Do
    
    'program's turn
    If gbProgramTurn = True And gbGameOver = False Then
      sMsg = "glTeacher: " & glTeacher & vbCrLf
      sMsg = sMsg & "glProgram: " & glProgram & vbCrLf
      sMsg = sMsg & "glAllCells: " & glAllCells & vbCrLf
      sMsg = sMsg & "glFreeCells: " & glFreeCells
      'MsgBox sMsg
      frmMain.SetForm 5
      nProgramMove = AIEngine
      frmMain.UpdateGrid nProgramMove
      
      'this updates display...does not play game
      nPlay = GetFactBasedMove
      gbProgramTurn = False
    End If
    DoEvents
    
    'teacher's turn
    If gbProgramTurn = False And gbGameOver = False Then
      frmMain.SetForm 4
    End If
    
    DoEvents
      
  Loop
End Sub

Public Sub CreateNewGame(nSize As Integer)
  
  On Error GoTo MyError
  
  'intialize global variables
  gnRows = nSize
  gnCols = nSize
  gnTotalCells = gnRows * gnCols
  gnProgramWins = 0
  gnProgramTies = 0
  gnProgramLosses = 0
  gsFilename = ""
  gsGameName = InputBox("Enter name of game.", "Game Name")
  gnGameType = 1
  ReDim uABS(0) As Pattern
  gnABSTotal = 0
  glTeacher = 0
  glProgram = 0
  glAllCells = 0
  glFreeCells = 0
  gbWinNotSaved = False
  
  'setup form appearance
  frmMain.SetForm 2
  frmMain.DrawNewGrid
  frmMain.staInfo.Panels(1).Text = "Game Name: " & gsGameName

  
  Exit Sub
MyError:
  gsForm = "MainSubs"
  gsProcedure = "CreateNewGame"
  ErrorHandler
End Sub
