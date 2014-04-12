Attribute VB_Name = "Global"
Option Explicit

'create global object f - field stuff
Public f As New CField              'ball, lines, field, etc
Public pbh() As New CPlayBook  'playbook for home team
Public pbv() As New CPlayBook   'playbook for home team
Public tp As New CTextParser
Public g As New CGame
Public b As New CBots          'stores all player dynamic/static data

'generates positions for all players
Public Sub GeneratePositions()
  Dim i As Integer
  Dim bRet As Boolean
  
  Randomize Timer
  
  'home
  For i = 1 To 12
    bRet = b.SetX(i, -20 + Rnd() * 40)
    bRet = b.SetY(i, 20 + Rnd() * 10)
    bRet = b.SetTeam(i, F_HOME_TEAM)
    bRet = b.SetMaxVelocity(i, 0.3 + Rnd() * 0.4)
    bRet = b.SetVelocity(i, b.GetMaxVelocity(i))
  Next i
  
  'visitors
  For i = 13 To MAX_BOTS
    bRet = b.SetX(i, -20 + Rnd() * 40)
    bRet = b.SetY(i, -20 + Rnd() * -10)
    bRet = b.SetTeam(i, F_VISITOR_TEAM)
    bRet = b.SetMaxVelocity(i, 0.3 + Rnd() * 0.4)
    bRet = b.SetVelocity(i, 0.3 + Rnd() * 0.4)
  Next i
  
End Sub
