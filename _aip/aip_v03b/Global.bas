Attribute VB_Name = "Global"
Option Explicit

Public Type aiprogram
  FullName As String
  Version As String
  Date As String
  FileExists As Boolean
  Filename As String
  Filepath As String
  FileContents As String
  GameChanged As Boolean
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
End Type

'game pieces
Public Type gamepiece
  ID As Integer        'unique identifer for this game symbol
  Symbol As String  'this is what user sees
  PromotionSymbol As String 'when promoted, this is new symbol
  Value As Integer   'this is a value of symbols..not useful with type 1 games
  Color As Long           'font color on cell color background
  ColorInverse As Long 'font color on inverse color background
  PlayColor As Integer '1=color only, 2=inverse only, 3=both colors allowable play
  CellsAllowed As String '1 = piece allowed, 0=disallowed.
  CellsPlayed As String   '1= piece played here, 0=piece not played
End Type

'global variables
Public Game As basicgame 'main game properties are ghere
Public AI As aiprogram 'main AIP properties
Public Program As gamepiece
Public Teacher As gamepiece

'Public gsFileName As String 'stores filename to open and close
'Public gsFile As String 'stores file contents read into program from file

'constants associated with game
Public Const MAX_ROWS = 15
Public Const MAX_COLS = 15
Public Const MIN_ROWS = 1
Public Const MIN_COLS = 1
Public Const GRID_LEFT = 1000
Public Const GRID_TOP = 1000
Public Const GRID_CENTER_X = 3500
Public Const GRID_CENTER_Y = 4000
Public Const CELL_HEIGHT = 500
Public Const CELL_WIDTH = 500
Public Const MIN_CHAR = 65 'lowest ASCII value for symbol
Public Const MAX_CHAR = 90 'highest ASCII value for symbol
Public Const PLAY_ONLY_COLOR = 1  'what is symbol allowed to play on
Public Const PLAY_ONLY_INVERSECOLOR = 2
Public Const PLAY_BOTH_COLORS = 3


'global variables associated with game grid
Public gnRows As Integer 'number of rows Max=15 Min=3
Public gnCols As Integer  'number of cols  Max=15 Min=3
Public gnTotalCells As Integer 'produce of gnRows * gnCols
Public glCellWidth As Long 'width of cell in twips
Public glCellHeight As Long 'height of cell in twips
Public glGridLeft As Long 'top-left corner of grid placement
Public glGridTop As Long
Public gbGridVisible As Boolean 'true if grid is supposed to be visible



