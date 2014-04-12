VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
' CGame.cls - July 25, 2005
' Manages game timing. In order to use this class
'1) Create a new object. I.e. public g as new CGame
'2) Add and enable 1 second timer to form
'3) Each 1 second, perform g.UpdateGameClock
'4) Start clock counting down with g.StartGameClock
'5) Stop clock counting down with g.StopGameClock
'6) Time can be increased with g.TimeInterval nint where nint = integer between
'    0 and 60 seconds
'******************************************************************************************
Option Explicit

'constants required for this game object
Private Const G_FIRST = 1
Private Const G_SECOND = 2
Private Const G_THIRD = 3
Private Const G_FOURTH = 4
Private Const G_QUARTER_INTERVAL = 900 '900 seconds or 15 minutes
Private Const G_PREGAME_INTERVAL = 180 'coin toss 3 minutes before kickoff
Private Const G_INTERMISSION = 120 '2 minutes between quarters
Private Const G_HALFTIME_INTERVAL = 720 '12 minutes at halftime
Private Const G_INTERMISSION_SD = 180 '3 minutes between reg game and sudden death
Private Const G_PLAY_CLOCK_SHORT = 25
Private Const G_PLAY_CLOCK_LONG = 40
Private Const G_TIMEOUT = 40  'regular timeout interval
Private Const G_TIMEOUT_TV = 110  'TV timeout

'phases of football game
Private Const G_PHASE_PREGAME = 0     'waiting for coin toss
Private Const G_PHASE_COINTOSS = 1     'coin toss 3 minute before kickoff
Private Const G_PHASE_FIRSTQ = 2          '1st quarter
Private Const G_PHASE_INT1 = 3               '2 min break
Private Const G_PHASE_SECONDQ = 4     '2nd quarter
Private Const G_PHASE_HALFTIME = 5     '12 minute halftime
Private Const G_PHASE_THIRDQ = 6         '3rd quarter
Private Const G_PHASE_INT2 = 7               '2 min break
Private Const G_PHASE_FOURTHQ = 8      '4th quarter
Private Const G_PHASE_INT3 = 9               '3 min break before sudden death overtime
Private Const G_PHASE_OT = 10               'sudden death overtime
Private Const G_PHASE_GAMEOVER = 11 'post game phase

'private variables
Private bGameClockRunning As Boolean 'true if clock is supposed to be running
Private bRegularGame As Boolean 'true if in regular game periods
Private bOTGame As Boolean 'true if in OT game
Private nPhase As Integer 'phase of game
Private sPhaseText As String 'name of phase
Private nTimeFactor As Integer '

Private tStartTime As Date
Private tStopTime As Date
Private nGameClockSeconds As Integer
Private nPlayClockSeconds As Integer
Private nGameQuarter As Integer
Private nGameHalf As Integer

Private Sub Class_Initialize()
  bGameClockRunning = False
  bRegularGame = True
  bOTGame = False
  nPhase = G_PHASE_COINTOSS
  nGameClockSeconds = G_PREGAME_INTERVAL
  nTimeFactor = 1
End Sub

'Changes phase and resets the game clock
Private Sub ChangePhase()
  nPhase = nPhase + 1
  If bOTGame = False And nPhase > G_PHASE_FOURTHQ Then
    nPhase = G_PHASE_GAMEOVER
  ElseIf bOTGame = True And nPhase > G_PHASE_OT Then
    nPhase = G_PHASE_GAMEOVER
  End If
  
  Select Case nPhase
    Case G_PHASE_PREGAME:
      nGameClockSeconds = 0
    Case G_PHASE_COINTOSS:
      nGameClockSeconds = G_PREGAME_INTERVAL
    Case G_PHASE_FIRSTQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT1:
      nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_SECONDQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_HALFTIME:
      nGameClockSeconds = G_HALFTIME_INTERVAL
    Case G_PHASE_THIRDQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT2:
      nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_FOURTHQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT3:
      nGameClockSeconds = G_INTERMISSION_SD
    Case G_PHASE_OT:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_GAMEOVER:
      nGameClockSeconds = 0
      
  End Select
  
End Sub

'this private sub ensures that the value of time does not exceed the number of seconds allowed
'for a particular phase
Private Sub ValidateClock()
  Select Case nPhase
    Case G_PHASE_PREGAME:
      'if ngameclockseconds
    Case G_PHASE_COINTOSS:
      If nGameClockSeconds > G_PREGAME_INTERVAL Then nGameClockSeconds = G_PREGAME_INTERVAL
    Case G_PHASE_FIRSTQ:
      If nGameClockSeconds > G_QUARTER_INTERVAL Then nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT1:
      If nGameClockSeconds > G_INTERMISSION Then nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_SECONDQ:
      If nGameClockSeconds > G_QUARTER_INTERVAL Then nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_HALFTIME:
      If nGameClockSeconds > G_HALFTIME_INTERVAL Then nGameClockSeconds = G_HALFTIME_INTERVAL
    Case G_PHASE_THIRDQ:
      If nGameClockSeconds > G_QUARTER_INTERVAL Then nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT2:
      If nGameClockSeconds > G_INTERMISSION Then nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_FOURTHQ:
      If nGameClockSeconds > G_QUARTER_INTERVAL Then nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT3:
      If nGameClockSeconds > G_INTERMISSION_SD Then nGameClockSeconds = G_INTERMISSION_SD
    Case G_PHASE_OT:
      If nGameClockSeconds > G_QUARTER_INTERVAL Then nGameClockSeconds = G_QUARTER_INTERVAL
  End Select
End Sub

'this must be performed every 1 second by main program
Public Sub UpdateGameClock()
  
  'update game clock only if true
  If bGameClockRunning = True Then
    nGameClockSeconds = nGameClockSeconds - nTimeFactor
    
    'if time expires, change phase
    If nGameClockSeconds < 1 Then
      bGameClockRunning = False
      nGameClockSeconds = 0
      ChangePhase
    End If
  
  End If
End Sub

'starts game clock running
Public Sub StartGameClock()
  bGameClockRunning = True
End Sub

'stops game clock running
Public Sub StopGameClock()
  bGameClockRunning = False
End Sub

'controls time factor to make time advance more quickly
Public Sub SetTimeFactor(nFactor As Integer)
  If nFactor < 0 Then Exit Sub
  If nFactor > 60 Then Exit Sub
  nTimeFactor = nFactor
End Sub

'adds time to game clock
Public Sub IncreaseGameClock(nSec As Integer)
  If nSec < 0 Then Exit Sub
  If nSec > 900 Then Exit Sub
  nGameClockSeconds = nGameClockSeconds + nSec
  ValidateClock
End Sub

'adds time to game clock
Public Sub DecreaseGameClock(nSec As Integer)
  If nSec < 0 Then Exit Sub
  If nSec > 900 Then Exit Sub
  nGameClockSeconds = nGameClockSeconds - nSec
  ValidateClock
End Sub

'returns game clock time
Public Function GameClock() As String
  Dim sMin As String
  Dim sSec As String
  
  sMin = CStr(nGameClockSeconds \ 60)
  sSec = CStr(nGameClockSeconds Mod 60)
  If Len(sMin) < 2 Then sMin = "0" & sMin
  If Len(sMin) < 1 Then sMin = "00"
  If Len(sSec) < 2 Then sSec = "0" & sSec
  If Len(sSec) < 1 Then sSec = "00"
  GameClock = sMin & ":" & sSec
End Function


'resets clock to its initial value for the particular phase
Public Sub ResetClock()
  Select Case nPhase
    Case G_PHASE_PREGAME:
      nGameClockSeconds = 0
    Case G_PHASE_COINTOSS:
      nGameClockSeconds = G_PREGAME_INTERVAL
    Case G_PHASE_FIRSTQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT1:
      nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_SECONDQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_HALFTIME:
      nGameClockSeconds = G_HALFTIME_INTERVAL
    Case G_PHASE_THIRDQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT2:
      nGameClockSeconds = G_INTERMISSION
    Case G_PHASE_FOURTHQ:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_INT3:
      nGameClockSeconds = G_INTERMISSION_SD
    Case G_PHASE_OT:
      nGameClockSeconds = G_QUARTER_INTERVAL
    Case G_PHASE_GAMEOVER:
      nGameClockSeconds = 0
      
  End Select

End Sub

'numerical value of game - Read Only
Public Property Get Phase() As Integer
  Phase = nPhase
End Property

'sets phase and automatically resets clock
Public Sub SetPhase(nNewPhase As Integer)
  If nNewPhase < 0 Then Exit Sub
  If nNewPhase > G_PHASE_GAMEOVER Then Exit Sub
  bGameClockRunning = False
  nPhase = nNewPhase
  ResetClock
End Sub


'Public Property Let Phase(ByVal vNewValue As Integer)
'  nPhase = vNewValue
'End Property

'text value of game phase - Read Only
Public Property Get PhaseText() As Variant
  Select Case nPhase
    Case G_PHASE_PREGAME:
      sPhaseText = "Pre-Game"
    Case G_PHASE_COINTOSS:
      sPhaseText = "Coin Toss"
    Case G_PHASE_FIRSTQ:
      sPhaseText = "First Quarter"
    Case G_PHASE_INT1:
      sPhaseText = "Intermission between 1st and 2nd Quarter"
    Case G_PHASE_SECONDQ:
      sPhaseText = "Second Quarter"
    Case G_PHASE_HALFTIME:
      sPhaseText = "Half Time"
    Case G_PHASE_THIRDQ:
      sPhaseText = "Third Quarter"
    Case G_PHASE_INT2:
      sPhaseText = "Intermission between 3rd and 4th Quarter"
    Case G_PHASE_FOURTHQ:
      sPhaseText = "Fourth Quarter"
    Case G_PHASE_INT3:
      sPhaseText = "Intermission Before Sudden Death"
    Case G_PHASE_OT:
      sPhaseText = "Sudden Death Overtime"
    Case G_PHASE_GAMEOVER:
      sPhaseText = "Post-Game"
      
  End Select
  PhaseText = sPhaseText
End Property

