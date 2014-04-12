Attribute VB_Name = "Variable"
Option Explicit

'global variables
Public gsRules As String 'stores all rules
Public gnRows As Integer 'total number of rows in grid
Public gnCols As Integer 'total number of columns in grid
Public gsVersion As String 'program version
Public gsVersionDate As String 'date of version change
Public gsProgramSymbol As String  'symbol played by pieces
Public gsTeacherSymbol As String
Public gnProgramValue As Integer  'numerical values corresponding to symbols, needed for arrays
Public gnTeacherValue As Integer
Public gnProgramWins As Integer
Public gnProgramLosses As Integer
Public gnProgramTies As Integer
Public gsFilename As String 'stores path to learning.txt..holds rules and accumulated knowledge
Public gsGameName As String
Public gnTotalCells As Integer 'total of cells

Public gnGameType As Integer 'determines type of game
                                              '1 - empty board, fill in grid 1 at time

'create type to hold patterns and other info
Public Type Pattern
  word  As Long '32 bit can hold a 5x5 array
  rel As Long 'stores equivalent relative pattern
  wins As Integer 'number of wins attributed to this pattern
  proven As Boolean 'TRUE if pattern is legal , FALSE if pattern is not legal
End Type

'uGame( ) array stays fixed in size during game play
'uGame(1).word stores grid cell status '0=empty, 1=full
'uGame(2).word stores teacher plays  '0=empty, 1=full
'uGame(3).word stores program plays '0=empty, 1=full
Public uGame(4) As Pattern
Public gnMoveCount As Integer 'stores move number

'uABS( ) array is updated with each new win pattern
'if the win is already in the array then the
'uABS( ).win value is updated
'All win patterns are automatically converted into
'relative positions.
Public uABS() As Pattern
Public gnABSTotal As Integer 'stores total number of ABS values in file/array

'uREL( ) array stores all relative patterns derived
'from real ABS winning patterns. Each time a win
'occurs this array is updated. Either a new REL
'pattern is recorded in uREL( ).word or the
'uREL ( ).win is updated
Public uREL() As Pattern
Public gnRELTotal As Integer 'stores total number of REL values in file/array

Public gnGameCount As Integer 'number of games played
Public gnPlayCount As Integer 'number of plays during a game
Public gbProgramTurn As Boolean 'whose turn is it?
Public gnGoFirst As Integer 'who goes first?
Public glCellColor As Long 'color of regular cell
Public glCellSelectedColor As Long 'color of selected cell
Public gbWinExists As Boolean 'true if a winning pattern exists


