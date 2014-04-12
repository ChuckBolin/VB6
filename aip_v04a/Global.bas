Attribute VB_Name = "Global"
Option Explicit

'global variables associated with the program
Public gsVersion As String
Public gsProgramName As String

'global variables associated with the game
Public gnGameType As Integer
Public gsGameName As String

'constants associated with game grid
Public Const MAX_ROWS = 15
Public Const MAX_COLS = 15
Public Const MIN_ROWS = 3
Public Const MIN_COLS = 3

'global variables associated game grid
Public gnRows As Integer 'number of rows Max=15 Min=3
Public gnCols As Integer  'number of cols  Max=15 Min=3
Public gnTotalCells As Integer 'produce of gnRows * gnCols
Public glCellColor As Long 'color of grid, usually white
Public glCellColorInverse As Long 'color of inverse cell color, such as checkerboard
Public glCellSelectedColor As Long 'color of cell when it is selected for purpose of teaching AIP
Public glCellWidth As Long 'width of cell in twips
Public glCellHeight As Long 'height of cell in twips
Public glGridLeft As Long 'top-left corner of grid placement
Public glGridTop As Long
Public gbGridVisible As Boolean 'TRUE it can be seen
Public gbGridReferenceOn As Boolean 'TRUE = user can see row/column ref. FALSE=user cannot see references
Public gbGridCheckerBoardOn As Boolean 'TRUE = use checkerboard
Public gnGridCheckerBoardType As Integer '1 or 2, affects the way the pattern begins. 1=top-left corner is inverse color, 2=std checker board
Public gbGridRandomPatternOn As Boolean 'TRUE = use random pattern
Public gnGridRandomPatternNum As Integer 'the number of random cells to fill with inverse color

