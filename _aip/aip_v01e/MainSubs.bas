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
  gsVersion = "v0.1e"
  gsVersionDate = "June 30, 2002"
  ReDim uABS(0)
  gnABSTotal = 0
  
  'user selects file to load
  'frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  'frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  'frmMain.dlgFile.ShowOpen
  'sFileName = frmMain.dlgFile.FileName
  
  'reads file
  'bFileFound = ReadFile(sFileName)
  
  'save filename as global
  'If bFileFound = True Then gsFilename = sFileName
  
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
  
  'main game play
  Do
    
    'program's turn
    If gbProgramTurn = True And gbGameOver = False Then
      sMsg = "glTeacher: " & glTeacher & vbCrLf
      sMsg = sMsg & "glProgram: " & glProgram & vbCrLf
      sMsg = sMsg & "glAllCells: " & glAllCells & vbCrLf
      sMsg = sMsg & "glAllCellsInverted: " & glAllCellsInverted
      'MsgBox sMsg
      frmMain.SetForm 5
      nProgramMove = AIEngine
      frmMain.UpdateGrid nProgramMove
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

