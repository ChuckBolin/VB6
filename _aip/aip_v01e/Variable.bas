Attribute VB_Name = "Variable"
Option Explicit

'global variables
'boolean
Public gbFilenameExists As Boolean 'true if a global file exists
Public gbWin As Boolean 'true if a winning pattern exists
Public gbProgramTurn As Boolean 'whose turn is it?
Public gbGameOver As Boolean
Public gbLoss As Boolean
Public gbTie As Boolean
Public gbTeacherDone As Boolean

'integers
Public gnGoFirst As Integer 'who goes first?
Public gnRows As Integer 'total number of rows in grid
Public gnCols As Integer 'total number of columns in grid
Public gnProgramValue As Integer  'numerical values corresponding to symbols, needed for arrays
Public gnTeacherValue As Integer
Public gnProgramWins As Integer
Public gnProgramLosses As Integer
Public gnProgramTies As Integer
Public gnTotalCells As Integer 'total of cells
Public gnGameType As Integer 'determines type of game
                                              '1 - empty board, fill in grid 1 at time
Public gnMoveCount As Integer 'stores move number
Public gnGameCount As Integer 'number of games played
Public gnPlayCount As Integer 'number of plays during a game

'long integers
Public glCellColor As Long 'color of regular cell
Public glCellSelectedColor As Long 'color of selected cell

'strings
Public gsRules As String 'stores all rules
Public gsVersion As String 'program version
Public gsVersionDate As String 'date of version change
Public gsProgramSymbol As String  'symbol played by pieces
Public gsTeacherSymbol As String
Public gsFilename As String 'stores path to learning.txt..holds rules and accumulated knowledge
Public gsGameName As String




