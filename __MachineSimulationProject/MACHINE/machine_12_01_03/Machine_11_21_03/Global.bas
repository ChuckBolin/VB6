Attribute VB_Name = "Global"
'***************************************************************
' Global Types, variables and arrays
' Written by Chuck Bolin, November 2003
'***************************************************************
Option Explicit

'***************************************************************
' T Y P E S     T Y P E S       T Y P E S       T Y P E S
'***************************************************************
'used for coordinates of a point
Public Type COORDINATE_PAIR
  x As Single
  y As Single
End Type

'used for coordinates of a quad
'names are direction of a compass..ie. NE = NorthEast
Public Type QUAD_COORDINATES
  NE As COORDINATE_PAIR
  SE As COORDINATE_PAIR
  SW As COORDINATE_PAIR
  NW As COORDINATE_PAIR
End Type

'used to identify boundaries of all objects on drawing board
Public Type OBJECT_DATA
  type As Integer
  quad As QUAD_COORDINATES
End Type

'defines properties of cylinder
Public Type OBJECT_CYLINDER
  value As Integer '0 to length of cylinder extension
  quad As OBJECT_DATA 'coordinates of tray corners
  max_length As Integer 'maximum length of extension
  speed As Integer 'number of twips/update when extending/retracting
  function_name As String 'name of cylinder...describes function
  designation As String 'technical designation
End Type

'defines properties of tray
Public Type OBJECT_TRAY
  quad As OBJECT_DATA 'coordinates of tray corners
  backcolor As Long  'color of tray
End Type

'****************************************************************
' A R R A Y S       A R R A Y S     A R R A Y S     A R R A Y S
'****************************************************************
Public gObj() As OBJECT_DATA
Public gCyl() As OBJECT_CYLINDER
Public gTray() As OBJECT_TRAY

'****************************************************************
' V A R I A B L E S     V A R I A B L E S   V A R I A B L E S
'****************************************************************

'cosmetic stuff
Public gstrProgramName As String
Public gstrProgramDate As String
Public gstrProgramVersion As String

'array variables
Public gintTotalObjects As Integer 'all objects
Public gintTotalCylinders As Integer
Public gintTotalTrays As Integer


