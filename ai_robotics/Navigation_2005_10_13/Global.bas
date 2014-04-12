Attribute VB_Name = "Global"
'sets up global variables, and virtual world initialization
'Things to add
' 1) GPS position with offset related to number of satelites - 10.12.05 CHB
' 2) LORAN position with radio stations
' 3) Monitor boundaries..stop when out of bounds
' 4) Start timer...count time
' 5) Display average speed/second

Option Explicit

'moving items
Public Type MOBILE_OBJECT
  X As Single
  Y As Single
  VX As Single
  VY As Single
  Velocity As Single
  Direction As Single
  Turn As Single 'amount of turning to affect direction
  MaxVel As Single
  MinVel As Single
  Energy As Single
End Type

'beacons - fixed navigation resources
Public Type NAV_BEACON
  X As Single
  Y As Single
  ID As Integer
  Offset As Single 'error source
End Type

'stores leg information
Public Type AUTO_LEG
  X1 As Single
  Y1 As Single
  X2 As Single
  Y2 As Single
  Width As Single '1/2 of lane width...mpy by 2
  Orientation As Integer '1=N,2=E,3=S,4=W
End Type

Public Type RECT_COORD
  X As Single
  Y As Single
End Type

Public Type BOX
  A As RECT_COORD
  B As RECT_COORD
  Num As Integer
End Type

Public Type OBSTACLE
  X As Single
  Y As Single
  Radius As Single
  Color As Long
End Type

Public bot As MOBILE_OBJECT
Public nav(3) As NAV_BEACON
Public dr As MOBILE_OBJECT 'this is dead reckoning info..not real..best guess
Public leg(8) As AUTO_LEG
Public g_nLegNum As Integer 'number of leg
Public g_nLastLegNum As Integer 'last leg number..1 is always first leg
Public g_nOdometer As Single
Public g_uOb(100) As OBSTACLE

'GPS stuff
Public GPS(4) As BOX
Public u_GPS As MOBILE_OBJECT
Public g_bGPSStatus As Boolean
Public g_nNumGPSSat As Integer 'number of sats...0 through 5
Public g_nGPSOffsetX As Single
Public g_nGPSOffsetY As Single
Public g_nGPSOffsetVel As Single
Public g_nGPSOffsetDir As Single

'constants used to make data realistic...error prone
Public Const NAV_DR_VEL_FACTOR = 1.1
Public Const NAV_DR_DIR_FACTOR = 1.1
Public Const NAV_TRIANGULATION_FACTOR = 1

Public Sub LoadVariables()
 Dim i As Integer
 
 'this is the bot
 bot.X = 10000
 bot.Y = 10000
 bot.Direction = 1.57
 bot.Velocity = 0
 bot.Turn = 0
 bot.MaxVel = 30
 bot.MinVel = -15
 bot.Energy = 100000
 dr.X = bot.X
 dr.Y = bot.Y
  
 g_nLegNum = 1
 g_nLastLegNum = 8
  
 'beacon data
 nav(1).ID = 1: nav(1).X = 9000: nav(1).Y = 12000
 nav(2).ID = 2: nav(2).X = 15000: nav(2).Y = 17000
 nav(3).ID = 3: nav(3).X = 18000: nav(3).Y = 11000
 
 'GPS boxes...indicates number of satelites
 GPS(1).A.X = 6500: GPS(1).A.Y = 15000: GPS(1).B.X = 11000: GPS(1).B.Y = 7000: GPS(1).Num = 4
 GPS(2).A.X = 11000: GPS(2).A.Y = 18000: GPS(2).B.X = 21000: GPS(2).B.Y = 13000: GPS(2).Num = 0
 GPS(3).A.X = 16000: GPS(3).A.Y = 13000: GPS(3).B.X = 28000: GPS(3).B.Y = 10000: GPS(3).Num = 3
 GPS(4).A.X = 11000: GPS(4).A.Y = 13000: GPS(4).B.X = 16000: GPS(4).B.Y = 9000: GPS(4).Num = 0

 'route 2
 leg(1).X1 = 10000: leg(1).Y1 = 10000: leg(1).X2 = 10000: leg(1).Y2 = 14000: leg(1).Width = 800: leg(1).Orientation = 1
 leg(2).X1 = 10000: leg(2).Y1 = 14000: leg(2).X2 = 12500: leg(2).Y2 = 14000: leg(2).Width = 600: leg(2).Orientation = 2
 leg(3).X1 = 12500: leg(3).Y1 = 14000: leg(3).X2 = 12500: leg(3).Y2 = 16000: leg(3).Width = 400: leg(3).Orientation = 1
 leg(4).X1 = 12500: leg(4).Y1 = 16000: leg(4).X2 = 27500: leg(4).Y2 = 16000: leg(4).Width = 400: leg(4).Orientation = 2
 leg(5).X1 = 27500: leg(5).Y1 = 16000: leg(5).X2 = 27500: leg(5).Y2 = 12000: leg(5).Width = 500: leg(5).Orientation = 3
 leg(6).X1 = 27500: leg(6).Y1 = 12000: leg(6).X2 = 12000: leg(6).Y2 = 12000: leg(6).Width = 300: leg(6).Orientation = 4
 leg(7).X1 = 12000: leg(7).Y1 = 12000: leg(7).X2 = 12000: leg(7).Y2 = 10000: leg(7).Width = 400: leg(7).Orientation = 3
 leg(8).X1 = 12000: leg(8).Y1 = 10000: leg(8).X2 = 10000: leg(8).Y2 = 10000: leg(8).Width = 500: leg(8).Orientation = 4

 'load obstacles
 For i = 0 To 100
   g_uOb(i).X = 8000 + GetRandomSingle(0, 20000)
   g_uOb(i).Y = 8000 + GetRandomSingle(0, 10000)
   g_uOb(i).Radius = 50 + GetRandomSingle(0, 350)
   g_uOb(i).Color = RGB(GetRandomInteger(0, 255), GetRandomInteger(0, 255), GetRandomInteger(0, 255))
 Next i
 
End Sub

'this returns the best estimation of robot position based upon
'available triangulation data
Public Function GetTriangulationPosition() As RECT_COORD
  Dim i As Integer
  
End Function

'returns the angular difference between two angles in radians
'each between 0 and 6.28 radians
Public Function GetAngularDifference(A As Single, B As Single) As Single
  Dim diff As Single
  
  diff = A - B
  
  If diff > PI Then
    diff = diff - (PI * 2)
  ElseIf diff < -PI Then
    diff = diff + (PI * 2)
  Else
    'do nothing
  End If
  GetAngularDifference = diff

End Function
