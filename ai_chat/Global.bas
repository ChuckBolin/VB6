Attribute VB_Name = "Global"
Option Explicit

'global objects
Public clepProfile As New Profile

'global variables
Public clepThinking As Boolean

'global sub MAIN
Public Sub Main()
  loadClepInfo
  loadHumanInputVariables
  
  clepThinking = False

  frmMain.Show
End Sub

Private Sub loadClepInfo()
  clepProfile.setFirstName "Clep"
  clepProfile.setLastName "AI"
End Sub
