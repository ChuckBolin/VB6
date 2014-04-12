Attribute VB_Name = "GameFunctions"
'*********************************************************************************
'  GAME FUNCTIONS.BAS
'
'
'*********************************************************************************
Option Explicit

'returns 1 or 2 to indicate winner of gametoss
Public Function CoinToss() As Integer
  
  Randomize Timer
  If Rnd < 0.5 Then
    CoinToss = 1
  Else
    CoinToss = 2
  End If
  
End Function

'returns 1 or 2 to indicate end zone defended by defense
Public Function GetEndZone() As Integer

  Randomize Timer
  If Rnd < 0.5 Then
    GetEndZone = 1
  Else
    GetEndZone = 2
  End If

End Function
