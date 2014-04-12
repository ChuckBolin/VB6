Attribute VB_Name = "Global"
Option Explicit

'public constants
Public Const HOME = 1

'public types and enums
Public Type PAIR
  Row As Single
  Column As Single
End Type

'public objects and variables
Public P As New CBots
Public g_nNumHome As Integer    'number of players for home
Public g_nNumVisitor As Integer 'number of players for visitors



'initialize all variables to begin
Public Sub LoadVariables()
  Dim i As Integer
  Dim j As Integer
  
  Dim bRet As Boolean
  Dim uPair As PAIR
  
  Randomize Timer
  g_nNumHome = 6
  
  'calc zones for improved player spacing
  uPair = GetTwoFactors(24)
  
  'draw boxes
  Dim h, w As Single
  h = Abs(frmMain.pic.ScaleHeight / uPair.Row)
  w = frmMain.pic.ScaleWidth / uPair.Column
  'MsgBox h
  
  Dim nCt As Integer
  
  For i = 1 To uPair.Row
    For j = 1 To uPair.Column
      nCt = nCt + 1
      
      frmMain.pic.ForeColor = RGB(GetRandomInteger(0, 255), GetRandomInteger(0, 255), GetRandomInteger(0, 255))
      frmMain.pic.Line ((j - 1) * w, i * h)-(j * w, (i - 1) * h), , BF

    Next j
  Next i
  
  'set initial positions, directions and speeds of all players
  For i = 1 To 6
    bRet = P.SetX(i, GetRandomSingle(2, 98))
    bRet = P.SetY(i, GetRandomSingle(2, 98))
    bRet = P.SetTargetX(i, GetRandomSingle(2, 98))
    bRet = P.SetTargetY(i, GetRandomSingle(2, 98))
    bRet = P.SetVelocity(i, 1)
    bRet = P.SetDiameter(i, 5)
    bRet = P.SetTeam(i, HOME)
    bRet = P.SetColor(i, vbGreen)
    
  Next i
  
  
  
  'MsgBox GetTwoFactors(96).Row & "  " & GetTwoFactors(96).Column
  
End Sub


'determines two factors for number of players
'to calculate zones
Public Function GetTwoFactors(num As Integer) As PAIR
  Dim i As Integer
  Dim bPrimeNum As Boolean
  Dim nPair As PAIR
  Dim nBest As PAIR
  
  'test to see if it is a prime number
  bPrimeNum = True
  nBest.Row = num  'prime factors
  nBest.Column = 1
  
  For i = 2 To num - 1
    If num Mod i = 0 Then
      bPrimeNum = False
      nPair.Column = i
      nPair.Row = num / nPair.Column
      If Abs(nPair.Row - nPair.Column) < Abs(nBest.Row - nBest.Column) Then
        nBest = nPair
      End If
    End If
  Next i
  GetTwoFactors = nBest
End Function
