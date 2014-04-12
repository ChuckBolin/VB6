Attribute VB_Name = "Global"
Option Explicit

Public Type aiprogram
  FullName As String
  Version As String
  Date As String
End Type

'global game variable
Public Type basicgame
  Name As String
  Type As Integer
  Rows As Integer
  Cols As Integer
  PatternColor As Long
  PatternColorInverse As Long
  PatternColorSelected As Long
  GridReferenceOn As Boolean
  PatternCheckerboardOn As Boolean
  PatternCheckerboardType As Integer
  PatternRandomOn As Boolean
  PatternRandomValue As Integer
  FilePath As String
  FileName As String
End Type

'global variables
Public Game As basicgame 'main game properties are ghere
Public AI As aiprogram 'main AIP properties
Public gsFileName As String 'stores filename to open and close
Public gsFile As String 'stores file contents read into program from file

'constants associated with game grid
Public Const MAX_ROWS = 15
Public Const MAX_COLS = 15
Public Const MIN_ROWS = 1
Public Const MIN_COLS = 1
Public Const GRID_LEFT = 1000
Public Const GRID_TOP = 1000
Public Const CELL_HEIGHT = 500
Public Const CELL_WIDTH = 500

'global variables associated with game grid
Public gnRows As Integer 'number of rows Max=15 Min=3
Public gnCols As Integer  'number of cols  Max=15 Min=3
Public gnTotalCells As Integer 'produce of gnRows * gnCols
Public glCellWidth As Long 'width of cell in twips
Public glCellHeight As Long 'height of cell in twips
Public glGridLeft As Long 'top-left corner of grid placement
Public glGridTop As Long
Public gbGridVisible As Boolean 'true if grid is supposed to be visible



