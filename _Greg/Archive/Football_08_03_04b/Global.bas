Attribute VB_Name = "Global"
Option Explicit

'create global object f - field stuff
Public f As New CField              'ball, lines, field, etc
Public p(23) As New CPerson        'coaches, players, referees...first half is for home
Public pbh() As New CPlayBook  'playbook for home team
Public pbv() As New CPlayBook   'playbook for home team
Public tp As New CTextParser
Public g As New CGame


'generates positions for all players
Public Sub GeneratePositions()
  Dim i As Integer
  
  Randomize Timer
  
  'home
  For i = 0 To 11
    p(i).x = -20 + Rnd() * 40
    p(i).Y = 20 + Rnd() * 10
  Next i
  
  'visitors
  For i = 12 To 23
    p(i).x = -20 + Rnd() * 40
    p(i).Y = -20 + Rnd() * -10
  Next i
  
  'assigned different TX, TY and vel
  For i = 0 To 23
    'p(i).TX = -60 + Rnd() * 120
   ' p(i).TY = -30 + Rnd() * 60
    p(i).Vel = 0.3 + Rnd() * 0.4
  Next i
  
End Sub
