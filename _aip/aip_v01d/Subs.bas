Attribute VB_Name = "Subs"
'*****************************************************************************
'T a b l e  O f  C o n t e n t s
'====================
' Main - program begins here
' Initialize - sets variables
' LoadFile - loads game file into memory
' PlayCoordinator - Direct play traffic
' AIEngine - determines program move based upon its mysterious
'                 algorithms... :-)
' ProcessABSWin - Takes ABS( );( );( );=x and feeds to array
' GetCoordinatePair - pulls row/col values from list of cooridinates
'*****************************************************************************
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

'*****************************************************************************
'GETABSSTRING
' Given a long variable with bits set, get the equivalent ABS string
'*****************************************************************************
Public Sub GetABSString(lNum As Long, sABS As String, nWin As Integer)
  Dim x As Integer
  Dim nRow As Integer
  Dim nCol As Integer
 
  On Error GoTo MyError
  
  'constructs ABS string from long integer
  sABS = "ABS"
  For x = 0 To 30
    If (lNum And 2 ^ x) Then
      nRow = GetCellRow(x + 1)
      nCol = GetCellCol(x + 1)
      sABS = sABS & "(" & CStr(nRow) & ", " & CStr(nCol) & ");"
    End If
   Next x
    
   'removes last
   sABS = Left(sABS, Len(sABS) - 1) & "=" & CStr(nWin)
    
  Exit Sub
MyError:
  gsForm = "Subs"
  gsProcedure = "GetABSString"
  ErrorHandler
End Sub


'*****************************************************************************
'LOADFILE
'Loads game information and knowledge into database arrays
'*****************************************************************************
Public Sub LoadFile(sFilename As String, bFileFound As Boolean)
  Dim sIn As String 'holds line read from file
  Dim nFile As Integer
  Dim x As Integer
  
  On Error GoTo MyError:
  bFileFound = False
  gsFilename = sFilename
  
  'constructs complete path and filename
  sFilename = App.Path & "\" & sFilename
  
  'if no file exists
  If FileLen(sFilename) < 1 Then
    sIn = "Filepath and name: " & sFilename & " does not exist!" & vbCrLf
    sIn = sIn & "Loading default values for 3x3 game."
    
    Exit Sub
  End If
  'file exists, open and read one line at a time
  nFile = FreeFile 'gets next file handle
  Open sFilename For Input As nFile
    Do
    Line Input #1, sIn
    sIn = UCase(LTrim(sIn)) 'forces string to uppercase for parsing next step
                                        'trims left whitespaces
    'reads only lines that are not empty
    If Len(sIn) > 0 Then
      
      'reads lines
      If Left(sIn, 1) = "'" Or Left(sIn, 2) = "//" Then
        'do nothing...comment found (VB or C format symbol)
      Else  'important variable here
        
        'system variable located here
        If InStr(1, sIn, "GAME.") Then
    
          x = 0: x = InStr(1, sIn, "=") 'gets positions of '=' sign
          If InStr(1, sIn, "GAME.NAME") Then gsGameName = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.TYPE") Then gnGameType = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.ROWS") Then gnRows = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.COLS") Then gnCols = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.TEACHER.SYMBOL") Then gsTeacherSymbol = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.TEACHER.VALUE") Then gnTeacherValue = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.SYMBOL") Then gsProgramSymbol = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.PROGRAM.VALUE") Then gnProgramValue = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.GOFIRST") Then gnGoFirst = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.WIN") Then gnProgramWins = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.LOSS") Then gnProgramLosses = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.TIE") Then gnProgramTies = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.CELL.COLOR") Then glCellColor = CLng(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.CELL.SELECTEDCOLOR") Then glCellSelectedColor = CLng(Mid(sIn, x + 1))
        End If
        
        'reads absolute win patterns
        If InStr(1, sIn, "ABS(") Then
          ProcessABSWin sIn
        End If
    End If
  End If
  Loop Until EOF(nFile)
  
  Close nFile
  bFileFound = True
  
  Exit Sub
MyError:
  gsForm = "Module Subs"
  gsProcedure = "LoadFile"
  ErrorHandler
End Sub

'*************************************************************************
'PLAYCOORDINATOR( )
'Decides whose turn it is and direct the play from there.
'*************************************************************************
Public Sub PlayCoordinator()
  Dim nRow As Integer
  Dim nCol As Integer
  Dim nGameType As Integer
  

  Do
    If gbProgramTurn = True Then AIEngine nRow, nCol
    If gbProgramTurn = False Then frmMain.cmdTeacher.Enabled = True
    DoEvents
      
  Loop
End Sub

'*************************************************************************
'AI ENGINE( )
'This sub returns the nRow and nCol of the determined play
'*************************************************************************
Public Sub AIEngine(nRow As Integer, nCol As Integer)
  
  'MsgBox "AI"
  gbProgramTurn = False
End Sub
'****************************************************************
' PROCESSABSWIN( )
'Takes an ABS string and adds it to array if it does not
'exist or updates the WIN value
'****************************************************************
Public Sub ProcessABSWin(sIn As String)
  Dim x, y As Integer
  Dim nIndex As Integer
  Dim nRow As Integer
  Dim nCol As Integer
  Dim lWord As Long 'stores ABS number in binary format matching pattern
  Dim sDat As String 'stores all loaded variable data for debug
  Dim bFound As Boolean
  Dim nWin As Integer
  
  On Error GoTo MyError:
  
  x = 0: x = InStr(1, sIn, "=") 'gets positions of '=' sign
      
  'extracts each coordinate pair from file --> lWord
  lWord = 0: nRow = 0: nCol = 0: nWin = 0
  For y = 1 To CountChar(sIn, "(")
    GetCoordinatePair sIn, y, nRow, nCol
    If nRow > 0 And nCol > 0 Then
      lWord = (2 ^ (GetCellPos(nRow, nCol) - 1)) Or lWord
    End If
  Next y
  nWin = CInt(Mid(sIn, x + 1)) 'extracts win from string
  
  'scan uABS array to see if win pattern already exists
  bFound = False
  
  For y = 0 To gnABSTotal 'LBound(uABS()) To UBound(uABS())
    
    'if it already exists then update win count
    If uABS(y).word = lWord Then
      bFound = True
      uABS(y).wins = uABS(y).wins + 1
    End If
  Next y
  
  'win does not already exist
  If bFound = False Then
    gnABSTotal = gnABSTotal + 1
    ReDim Preserve uABS(gnABSTotal) As Pattern 'increase array by one
    uABS(gnABSTotal).word = lWord
    uABS(gnABSTotal).wins = nWin
  End If
 
  Exit Sub
MyError:
  gsForm = "Module Subs"
  gsProcedure = "ProcessABSWin"
  ErrorHandler
End Sub
Private Sub ShowPattern()
  Dim sInput As String
  Dim nRow As Integer
  Dim nCol As Integer
  Dim x As Integer
  On Error GoTo MyError
  
  msgGrid.Clear
  For x = 1 To gnCols
    msgGrid.TextMatrix(0, x) = CStr(x) 'adds references points to rows/cols
    msgGrid.TextMatrix(x, 0) = CStr(x)
  Next x
  'sInput = cboPatterns.Text
  If Len(sInput) < 1 Then Exit Sub
  For x = 1 To CountChar(sInput, "(")
    GetCoordinatePair sInput, x, nRow, nCol
    If nRow > 0 And nCol > 0 Then
      msgGrid.TextMatrix(nRow, nCol) = "X"
    End If
  Next x
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowPattern"
End Sub
'*****************************************************************************
'GETCOORDINATEPAIR( )
'Given a string that has coordinate pairs between ( ) this sub
'returns the row and column value for the selected pair
'To get the row col for the second set of parentheses in ABS(2,2);(2,3);(2,4)=1
'use nIndex=2.  nRow=2 and nCol=3 will be returned with SUB.
' sIn=string to parse
' nIndex=the coordinate pair to evaluate 1 to max number
' nRow=row value returned
' nCol=column value returned
'*****************************************************************************
Public Sub GetCoordinatePair(sIn As String, nIndex As Integer, nRow As Integer, nCol As Integer)
  Dim nLP As Integer 'position of left paren
  Dim nRP As Integer 'position of right paren
  Dim nMaxLP As Integer 'maximum number of left parens
  Dim nMaxRP As Integer 'max number for right parens
  Dim nComma As Integer 'position of comman seperating coordinates
  Dim a, b, x As Integer
  On Error GoTo MyError:
  
  'set default values
  nRow = 0: nCol = 0
  
  'there must be equal number of parenthesis greater than 0
  If Len(sIn) < 1 Then Exit Sub
  If Len(nIndex) < 1 Then Exit Sub
  nMaxLP = CountChar(sIn, "(")
  nMaxRP = CountChar(sIn, ")")
  If nMaxLP < 1 Or nMaxRP < 1 Or nMaxLP <> nMaxRP Then Exit Sub
  If nIndex > nMaxLP Then Exit Sub
    
  'find the position of left and right parenthesis that holds required index
  'gets position of left parenthesis
  b = 1
  Do
    a = InStr(b, sIn, "(")
    If a > 0 Then
      x = x + 1 'counts occurence of left paren
      b = a + 1
    End If
  Loop Until x = nIndex
  nLP = a
  
  'gets position of right parenthesis
  b = 1: a = 0: x = 0
  Do
    a = InStr(b, sIn, ")")
    If a > 0 Then
      x = x + 1
      b = a + 1
    End If
  Loop Until x = nIndex
  nRP = a
   
  'coordinates are between nLP and nRP position separated by a comma
  a = 0
  nComma = InStr(nLP, sIn, ",")
  If nComma <= 0 Then Exit Sub
  If nComma >= nRP Then Exit Sub
  
  'extract row
  nRow = CInt(Mid(sIn, nLP + 1, nComma - nLP - 1))
  nCol = CInt(Mid(sIn, nComma + 1, nRP - nComma - 1))
  Exit Sub
  
MyError:
  gsForm = "Module Subs"
  gsProcedure = "GetCoordinatePair"
  ErrorHandler
  
End Sub



