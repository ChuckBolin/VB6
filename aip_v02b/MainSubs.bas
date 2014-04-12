Attribute VB_Name = "MainSubs"
Option Explicit

'*****************************************************************************
'MAIN
'Starts program
'*****************************************************************************
Public Sub Main()
  Dim sBuffer As String
  Dim lReturn As Long
  Dim sMsg As String
  Dim nFilepath As Integer
  
  On Error GoTo MyError
  nFilepath = 0
  
  'need to verify two files exist in system folder 2) MSFLXGRD.OCX  2) MSCOMCTL.OCX. These are needed
  'for the Grid, Status bar and Tab strip
  sBuffer = Space(MAX_PATH)
  lReturn = GetSystemDirectory(sBuffer, MAX_PATH)
  sBuffer = Left(sBuffer, lReturn)
  
  'read file lengths
  If FileLen(sBuffer & "\MSFLXGRD.OCX") > 0 Then nFilepath = 1
  If FileLen(sBuffer & "\MSCOMCTL.OCX") > 0 Then nFilepath = 2
  
  'program initialization
  Initialize
  frmMain.Show
  Exit Sub
  
MyError:
  
  'most likely one of two OCX files are missing
  If Err.Number = 53 Then
    'missing flex grid control
    If nFilepath = 0 Then
      sMsg = "AIP requires MSFLXGRD.OCX to be loaded into the system folder and to" & vbCrLf
      sMsg = sMsg & "be registered.  Visit http://www.clg-net.com/ai/aip.htm to download the" & vbCrLf
      sMsg = sMsg & "controls and instructions for proper installation."
      MsgBox sMsg
      End
    End If
    If nFilepath = 1 Then
      sMsg = "AIP requires MSCcOMCTL.OCX to be loaded into the system folder and to" & vbCrLf
      sMsg = sMsg & "be registered.  Visit http://www.clg-net.com/ai/aip.htm to download the" & vbCrLf
      sMsg = sMsg & "controls and instructions for proper installation."
      MsgBox sMsg
     End
    End If
  Else
    sMsg = "Error no.: " & Err.Number & "  " & Err.Description & " during  startup." & vbCrLf
    sMsg = sMsg & "Program cannot load.  Verify all files have been downloaded to the same" & vbCrLf
    sMsg = sMsg & "folder.  Visit http://www.clg-net.com/ai/aip.htm for the latest program files. "
    MsgBox sMsg
    End
  End If
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
  gsVersion = "v0.2b"
  gsVersionDate = "July 13, 2002"
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
      If gbPlayCheck = False Then gbPlayCheck = True
    End If

    
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
  
  'setup form appearance and clear previous values
  frmMain.txtProgramWins.Text = 0
  frmMain.txtProgramLosses.Text = 0
  frmMain.txtProgramTies.Text = 0
  If frmMain.tabInfo.SelectedItem.Index = 3 Then frmMain.txtKB.Text = "No data loaded!"
  frmMain.SetForm 2
  frmMain.DrawNewGrid
  frmMain.staInfo.Panels(1).Text = "Game Name: " & gsGameName

  
  Exit Sub
MyError:
  gsForm = "MainSubs"
  gsProcedure = "CreateNewGame"
  ErrorHandler
End Sub
