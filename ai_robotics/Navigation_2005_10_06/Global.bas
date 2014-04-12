Attribute VB_Name = "Global"
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

Public bot As MOBILE_OBJECT
Public nav(3) As NAV_BEACON
Public dr As MOBILE_OBJECT 'this is dead reckoning info..not real..best guess
Public leg(5) As AUTO_LEG
Public g_nLegNum As Integer 'number of leg
Public g_nLastLegNum As Integer 'last leg number

Public Sub LoadVariables()
 
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
 g_nLastLegNum = 5
  
 'beacon data
 nav(1).ID = 1: nav(1).X = 8000: nav(1).Y = 13200
 nav(2).ID = 2: nav(2).X = 7000: nav(2).Y = 8500
 nav(3).ID = 3: nav(3).X = 13000: nav(3).Y = 11500
 
 'route data
 leg(1).X1 = 10000: leg(1).Y1 = 10000: leg(1).X2 = 10000: leg(1).Y2 = 13000: leg(1).Width = 300: leg(1).Orientation = 1
 leg(2).X1 = 10000: leg(2).Y1 = 13000: leg(2).X2 = 17000: leg(2).Y2 = 13000: leg(2).Width = 300: leg(2).Orientation = 2
 leg(3).X1 = 17000: leg(3).Y1 = 13000: leg(3).X2 = 17000: leg(3).Y2 = 7000: leg(3).Width = 300: leg(3).Orientation = 3
 leg(4).X1 = 17000: leg(4).Y1 = 7000: leg(4).X2 = 10000: leg(4).Y2 = 7000: leg(4).Width = 300: leg(4).Orientation = 4
 leg(5).X1 = 10000: leg(5).Y1 = 7000: leg(5).X2 = 10000: leg(5).Y2 = 10000: leg(5).Width = 300: leg(5).Orientation = 1

End Sub
