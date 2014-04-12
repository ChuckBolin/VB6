Attribute VB_Name = "Subs"
'*****************************************************************************
'T a b l e  O f  C o n t e n t s
'====================
' Main - program begins here
' Initialize - sets variables
' NewGameType - allows change of grid during run-time
' GetCoordinatePair - pulls row/col values from list of cooridinates
' ParseInput - reads input string and adds to array
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

  'sets required variables
  gsRules = "" 'stores all rules collected during program
  gnRows = 4  'defines 4 x 4 board
  gnCols = 4
  'Public uABS(10) As Pattern 'stores absolute win patterns
  
  
  
  'performs required functions/subs
  
  
End Sub

'*****************************************************************************
'NEWGAMETYPE
'Sets new row/col for game board
'*****************************************************************************
Public Sub NewGameType(nRow As Integer, nCol As Integer)
  gsRules = "" 'stores all rules collected during program
  gnRows = nRow  'defines R x C board
  gnCols = nCol
End Sub

'*****************************************************************************
'GETCOORDINATEPAIR( )
'Given a string that has coordinate pairs between ( ) this sub
'returns the row and column value for the selected pair
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

'*************************************************************
'PARSEINPUT
'adds this input to array if it is a rule
' sIn=string to be parsed
' bError=Returns true if error occurs
'*************************************************************
Public Sub ParseInput(sIn As String, bError As Boolean)
  Dim nRow As Integer
  Dim nCol As Integer
  bError = True
  On Error GoTo MyError
  
  If Len(sIn) < 1 Then Exit Sub
  sIn = UCase(sIn)
  
  'process absolute patterns that constitute a win
  If Left(sIn, 3) = "ABS" Then
    GetCoordinatePair sIn, 2, nRow, nCol
    If nRow < 1 Or nCol < 1 Then
      bError = True
    Else
      bError = False
    End If
  End If
  Exit Sub
  
MyError:
  gsForm = "Module Subs"
  gsProcedure = "ParseInput'"
  ErrorHandler
End Sub

