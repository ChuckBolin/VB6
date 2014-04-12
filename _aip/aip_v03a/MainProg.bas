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
  
  Dim bFound As Boolean
  
  'seed randomizer
  Randomize Timer
  
  'load AIP variables
  AI.FullName = "Artificial Intelligence Porgram (AIP)"
  AI.Version = "0.3a"
  AI.Date = "July 19, 2002"

  'let's grab default global information
  LoadDefaultData
  Game.FilePath = App.Path & "\"
  Game.FileName = "game1.aip"
  gsFileName = Game.FilePath & Game.FileName
  bFound = LoadGameData(Game.FileName)
  gbGridVisible = True  'show grid at startup
  
  'loaded for convenience of program calculations.
  gnRows = Game.Rows
  gnCols = Game.Cols
  gnTotalCells = gnRows * gnCols

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
            
            
  'main game variable
  Game.Name = "Standard"
  Game.Type = 1
  Game.Rows = 8
  Game.Cols = 8
  Game.PatternColor = RGB(255, 255, 255)
  Game.PatternColorInverse = RGB(255, 0, 0)
  Game.PatternColorSelected = RGB(0, 255, 0)
  Game.GridReferenceOn = True
  Game.PatternCheckerboardOn = False
  Game.PatternCheckerboardType = 2
  Game.PatternRandomOn = False
  Game.PatternRandomValue = 5
  
  'grid presets
  glGridLeft = GRID_LEFT
  glGridTop = GRID_TOP
  glCellHeight = CELL_HEIGHT
  glCellWidth = CELL_WIDTH

  Exit Sub
MyError:
  gsForm = "MainProg"
  gsProcedure = "LoadDefaultData"
  ErrorHandler
End Sub


'****************************************************************
'LOAD FILE DATA
'Read file and loads global data.  Returns true if file
'is found
'****************************************************************
Public Function LoadGameData(sFile As String) As Boolean
  On Error GoTo MyError
  
  Dim nFile As Integer ' file handle
  Dim sInput As String 'stores line read in from file
  Dim sProp As String 'extracted portion of line input that has property
  Dim sVar As String 'value of variables after AIPSL keywords
  Dim nEqual As Integer 'position of equal sign in string
  gsLoc = "10"
  LoadGameData = False 'default value...assume file does not exist unless proven wrong
 
  'verify file exists
  If Len(sFile) < 1 Then Exit Function
  nFile = FreeFile
  Open sFile For Append As nFile
    If LOF(nFile) < 1 Then Close nFile: Exit Function
  Close nFile
  'read file and loads global variables
  Open sFile For Input As nFile
    gsFile = "" 'clears global file string
    Do
      Line Input #nFile, sInput
      RemoveWhitespaces sInput 'get rid of spaces and tabs
      sInput = UCase(sInput) 'make everything uppercase
      gsFile = gsFile & sInput & vbCrLf 'adds line to global file string
      If InStr(1, sInput, "'") Then  'extract only info before apostrophe
        If Left(sInput, 1) = "'" Then
          sInput = ""
        Else
          sInput = Left(sInput, (InStr(1, sInput, "'")) - 1)
        End If
      End If
                
      'obtains for GAME object
      gsLoc = "60"
      If Left(sInput, 5) = "GAME." Then
        nEqual = InStr(1, sInput, "=")
        If nEqual > 0 Then
          sProp = Mid(sInput, 6, nEqual - 6) 'grab rest of info beyond
          sVar = Mid(sInput, nEqual + 1) 'actual variable data
          Select Case sProp
            Case "NAME"
              Game.Name = sVar
            Case "TYPE"
              If IsNumeric(sVar) Then Game.Type = CInt(sVar)
            Case "ROWS"
              If IsNumeric(sVar) Then Game.Rows = CInt(sVar)
            Case "COLS"
              If IsNumeric(sVar) Then Game.Cols = CInt(sVar)
            Case "PATTERN_COLOR"
              If IsNumeric(sVar) Then Game.PatternColor = CLng(sVar)
            Case "PATTERN_COLOR_INVERSE"
             If IsNumeric(sVar) Then Game.PatternColorInverse = CLng(sVar)
            Case "PATTERN_COLOR_SELECTED"
              If IsNumeric(sVar) Then Game.PatternColorSelected = CLng(sVar)
            Case "GRID_REFERENCE_ON"
              If IsNumeric(sVar) Then
                If CInt(sVar) = 1 Then
                  Game.GridReferenceOn = True
                Else
                  Game.GridReferenceOn = False
                End If
              End If
            Case "PATTERN_CHECKERBOARD_ON"
              If IsNumeric(sVar) Then
                If CInt(sVar) = 1 Then
                  Game.PatternCheckerboardOn = True
                Else
                  Game.PatternCheckerboardOn = False
                End If
              End If
            Case "PATTERN_CHECKERBOARD_TYPE"
              If IsNumeric(sVar) Then Game.PatternCheckerboardType = CInt(sVar)
            Case "PATTERN_RANDOM_ON"
              If IsNumeric(sVar) Then
                If CInt(sVar) = 1 Then
                  Game.PatternRandomOn = True
                Else
                  Game.PatternRandomOn = False
                End If
              End If
            Case "PATTERN_RANDOM_VALUE"
              If IsNumeric(sVar) Then Game.PatternRandomValue = CInt(sVar)
          End Select
        End If
      End If
 
    Loop Until EOF(nFile)
    
    'process global variables
    gsLoc = "70"
    LoadGameData = True
    gnTotalCells = gnRows * gnCols
  Close nFile
  
    gsLoc = "1000"

  Exit Function
MyError:
  gsForm = "MainProg"
  gsProcedure = "LoadGameData"
  ErrorHandler
End Function

'***************************************************************************
'   SAVE FILE DATA
'
'***************************************************************************
Public Function SaveGameData(sFile As String) As Boolean
  On Error GoTo MyError
  
  Dim x As Integer 'for general counting
  Dim nFile As Integer 'file handle
  Dim sFileName As String 'extracted filename from filespec
    
  'get filehandle
  nFile = FreeFile
  
  'extract filename from filespec
  For x = Len(sFile) To 1 Step -1
    If Mid(sFile, x, 1) = "\" Then
      sFileName = Mid(sFile, x + 1)
      Exit For
    End If
  Next x
  
  'write to file
  Open sFile For Output As nFile
    Print #nFile, "'**********************************"
    Print #nFile, "'" & LCase(sFileName)
    Print #nFile, "'Date: " & CStr(Date)
    Print #nFile, "'**********************************" & vbCrLf
    Print #nFile, "'system variables"
    Print #nFile, "GAME.NAME=" & Game.Name
    Print #nFile, "GAME.TYPE=" & CStr(Game.Type)
    Print #nFile, "GAME.ROWS=" & CStr(Game.Rows)
    Print #nFile, "GAME.COLS=" & CStr(Game.Cols)
    Print #nFile, "GAME.PATTERN_COLOR=" & CStr(Game.PatternColor)
    Print #nFile, "GAME.PATTERN_COLOR_INVERSE=" & CStr(Game.PatternColorInverse)
    Print #nFile, "GAME.PATTERN_COLOR_SELECTED=" & CStr(Game.PatternColorSelected)
    If Game.GridReferenceOn = True Then
      Print #nFile, "GAME.GRID_REFRENCE_ON=1"
    Else
      Print #nFile, "GAME.GRID_REFRENCE_ON=0"
    End If
    If Game.PatternCheckerboardOn = True Then
      Print #nFile, "GAME.PATTERN_CHECKERBOARD_ON=1"
    Else
      Print #nFile, "GAME.PATTERN_CHECKERBOARD_ON=0"
    End If
    Print #nFile, "GAME.PATTERN_CHECKERBOARD_TYPE=" & CStr(Game.PatternCheckerboardType)
    If Game.PatternRandomOn = True Then
      Print #nFile, "GAME.PATTERN_RANDOM_ON=1"
    Else
      Print #nFile, "GAME.PATTERN_RANDOM_ON=0"
    End If
    Print #nFile, "GAME.PATTERN_RANDOM_VALUE=" & CStr(Game.PatternRandomValue)
    
  Close nFile
  
  
  Exit Function
MyError:
  gsForm = "MainProg"
  gsProcedure = "SaveGameData"
  ErrorHandler
End Function

