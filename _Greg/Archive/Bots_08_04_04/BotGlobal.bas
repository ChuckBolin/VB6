Attribute VB_Name = "BotGlobal"
Option Explicit

Public Const MAX_BOTS = 0
Public Const MAX_BOT_TYPES = 5

Public bot(MAX_BOTS) As New CBot

'sets initial parameters
Public Sub InitializeBots()
  Dim i As Integer
  
  For i = 0 To MAX_BOTS
    bot(i).X = GetRandomSingle(1, 99)
    bot(i).Y = GetRandomSingle(1, 99)
    bot(i).TX = GetRandomSingle(1, 99)
    bot(i).TY = GetRandomSingle(1, 99)
    bot(i).BotType = GetRandomSingle(0, MAX_BOT_TYPES)
    'frmMain.Text1.Text = frmMain.Text1.Text & bot(i).TX & ", " & bot(i).TY & vbCrLf
    'frmMain.Text1.Text = frmMain.Text1.Text & bot(i).X & ", " & bot(i).Y & vbCrLf
    'direction is in compass radians
    bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
    frmMain.Text1.Text = frmMain.Text1.Text & bot(i).X & vbTab & bot(i).Y & vbTab & bot(i).TX & vbTab & bot(i).TY & vbTab & bot(i).Direction & vbCrLf
    bot(i).Speed = 0.3
  Next i
End Sub


Public Sub UpdateBots()
  Dim i As Integer
  Dim nDir As Single 'stores direction
    
  For i = 0 To MAX_BOTS
    If bot(i).Found = False Then
      bot(i).Direction = GetTargetDirection2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY)
      nDir = CRtoR(bot(i).Direction)
      bot(i).DX = bot(i).Speed * Cos(nDir)
      bot(i).DY = bot(i).Speed * Sin(nDir)
      bot(i).X = bot(i).X + bot(i).DX
      bot(i).Y = bot(i).Y + bot(i).DY
      If GetTargetDistance2D(bot(i).X, bot(i).Y, bot(i).TX, bot(i).TY) < 5 Then bot(i).Found = True
    End If
  Next i
End Sub
