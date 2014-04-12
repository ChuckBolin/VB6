Attribute VB_Name = "Auto"
Option Explicit
'integrating all the data from all the sensors is a major undertaking,
'especially when you figure that the they are using two laser scanners,
'a video camera, a compass, and a differential GPS that is accurate to 1 inch.

'*****************************************************************
' This is the Autonomous routine. Its purpose is to modify two
' variables:
' Inputs:   bot.X   'used only for testing...not competition
'           bot.Y
'           bot.Direction
'           dr.X
'           dr.Y
'           dr.Direction
'           u_u_GPS.X
'           u_u_GPS.Y
'           u_u_GPS.Direction
'           u_u_GPS.Velocity
'           lr(0)..lr(180)  laser ranging
'           leg(1).X1
'           leg(1).Y1
'           leg(1).X2
'           leg(1).Y2
'           leg(1).Orientation
'           leg(1).Width
'           g_uLR( ).Range
'           g_uLR( ).Bearing   Points 1 to 36, 5 deg each slice
' Outputs: bot.Turn
'          bot.velocity
'
' Try using these functions with bot position and waypoint info (leg())
'  GetTargetDistance2D ()
'  GetTargetDirection2D()
'  GetAngularDifference
'*****************************************************************
'Greg 10.20.05
Public Sub Autonomous4()
  Dim nAngDiff As Single
  Dim nWPDir As Single
  Dim nWPDirOld As Single
  Dim nWPDist As Single
  Dim nWPDistOld As Single
  Dim nV1, nV2, nV3, nV4, nV5 As Integer
  
  Static nWPIndex As Integer
  
  nV1 = 15: nV2 = nV1 * 2: nV3 = nV1 * 3: nV4 = nV1 * 4: nV5 = nV1 * 5
  
  bot.Velocity = nV5
  
  If nWPIndex = 0 Then nWPIndex = 1
  
  nWPDir = GetTargetDirection2D(u_GPS.X, u_GPS.Y, leg(nWPIndex).X2, leg(nWPIndex).Y2)
  nWPDirOld = GetTargetDirection2D(u_GPS.X, u_GPS.Y, leg(nWPIndex - 1).X2, leg(nWPIndex - 1).Y2)
  
  nAngDiff = GetAngularDifference(bot.Direction, nWPDir)
  
  nWPDist = GetTargetDistance2D(u_GPS.X, u_GPS.Y, leg(nWPIndex).X2, leg(nWPIndex).Y2)
  nWPDistOld = GetTargetDistance2D(u_GPS.X, u_GPS.Y, leg(nWPIndex - 1).X2, leg(nWPIndex - 1).Y2)
  
  frmMain.Caption = nWPIndex
  
  If nWPDist < 30 Then
    nWPIndex = nWPIndex + 1
    If nWPIndex > g_nLastLegNum Then nWPIndex = 1
  End If
  
  If nAngDiff < 0.005 Then
    bot.Turn = 0.1
  ElseIf nAngDiff > 0.005 Then
    bot.Turn = -0.1
  Else
    bot.Turn = 0
  End If
  
  If nWPDist < 250 Then
    bot.Velocity = nV1
  ElseIf nWPDistOld < 250 Then
    bot.Velocity = nV1
  ElseIf nWPDist < 500 Then
    bot.Velocity = nV2
  ElseIf nWPDistOld < 500 Then
    bot.Velocity = nV2
  ElseIf nWPDist < 750 Then
    bot.Velocity = nV3
  ElseIf nWPDistOld < 750 Then
    bot.Velocity = nV3
  ElseIf nWPDist < 1000 Then
    bot.Velocity = nV4
  ElseIf nWPDistOld < 1000 Then
    bot.Velocity = nV4
  ElseIf nWPDist < 1250 Then
    bot.Velocity = nV5
  ElseIf nWPDistOld < 1250 Then
    bot.Velocity = nV5
  End If
   
End Sub

Public Sub Autonomous3()
  Dim nAngDiff As Single
  Dim nWPDir As Single
  Dim nWPDist As Single
  
  Static nWPIndex As Integer
  
  If nWPIndex = 0 Then nWPIndex = 1
  bot.Velocity = 35
  
  nWPDir = GetTargetDirection2D(bot.X, bot.Y, leg(nWPIndex).X2, leg(nWPIndex).Y2)
  nAngDiff = GetAngularDifference(bot.Direction, nWPDir)
  nWPDist = GetTargetDistance2D(bot.X, bot.Y, leg(nWPIndex).X2, leg(nWPIndex).Y2)
  
  If nWPDist < 1000 Then
    nWPIndex = nWPIndex + 1
    If nWPIndex > g_nLastLegNum Then nWPIndex = 1
    'bot.Velocity = 0
  End If
  
  frmMain.Caption = nWPIndex
  
  If nAngDiff < 1 Then
    bot.Turn = 0.01 '''bot.Turn + 0.01
     bot.Velocity = 20
  ElseIf nAngDiff > 1 Then
    bot.Turn = -0.01 '''bot.Turn - 0.01
     bot.Velocity = 20
  Else
    bot.Turn = 0
  End If
  
  
  'bot.Turn = -0.01
End Sub

'October 12, 2005 by Chuck Bolin
Public Sub Autonomous()
  Dim nWPDir As Single
  Dim nWPDist As Single
  Dim nDirDiff As Single 'difference between nWPDir and bot.dir
  Dim nLeft, nRight, nFront As Single
  
  'determines what feeds the auto algorithms
  If frmMain.optAutoActual.Value = True Then
    nWPDir = GetTargetDirection2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
    nWPDist = GetTargetDistance2D(bot.X, bot.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  ElseIf frmMain.optAutoDR.Value = True Then
    nWPDir = GetTargetDirection2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
    nWPDist = GetTargetDistance2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  ElseIf frmMain.optAutoGPS.Value = True Then
    nWPDir = GetTargetDirection2D(u_GPS.X, u_GPS.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
    nWPDist = GetTargetDistance2D(u_GPS.X, u_GPS.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
  
  'this is a programmed response
  ElseIf frmMain.optAutoProg.Value = True Then
    If g_bGPSStatus = True Then  'gps is default
      nWPDir = GetTargetDirection2D(u_GPS.X, u_GPS.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
      nWPDist = GetTargetDistance2D(u_GPS.X, u_GPS.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
      
      'this resets DR
      dr.X = u_GPS.X
      dr.Y = u_GPS.Y
      dr.Velocity = u_GPS.Velocity
      dr.Direction = u_GPS.Direction

    Else 'no gps use dr
      nWPDir = GetTargetDirection2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
      nWPDist = GetTargetDistance2D(dr.X, dr.Y, leg(g_nLegNum).X2, leg(g_nLegNum).Y2)
    End If
  End If
  
  'calc angular difference
  nDirDiff = GetAngularDifference(nWPDir, bot.Direction)
    
  'determine turning amount..watch for collision
  nLeft = g_uLR(19).Range + g_uLR(20).Range + g_uLR(21).Range + g_uLR(22).Range
  nRight = g_uLR(15).Range + g_uLR(16).Range + g_uLR(17).Range + g_uLR(18).Range
  nFront = g_uLR(17).Range + g_uLR(18).Range + g_uLR(19).Range + g_uLR(20).Range
  
  If nLeft > 6000 And nRight > 6000 Then
    If nDirDiff > 0 Then
      If bot.Velocity <> 0 Then bot.Turn = (nDirDiff / bot.Velocity)
    ElseIf nDirDiff < 0 Then
      If bot.Velocity <> 0 Then bot.Turn = (nDirDiff / bot.Velocity)
    Else
      bot.Turn = 0
    End If
  ElseIf nLeft <= 4000 And nRight > 6000 Then
    If bot.Velocity <> 0 Then
      bot.Turn = (nDirDiff / bot.Velocity) + 0.1
    Else
      bot.Turn = bot.Turn + 0.1
    End If
  ElseIf nLeft > 6000 And nRight <= 4000 Then
    If bot.Velocity <> 0 Then
      bot.Turn = (nDirDiff / bot.Velocity) - 0.1
    Else
      bot.Turn = bot.Turn - 0.1
    End If
  End If
    
    
  'get distance to next waypoint
  If nWPDist > leg(g_nLegNum).Width * 2 And Abs(nDirDiff) < 0.2 Then
    bot.Velocity = leg(g_nLegNum).Width / 10
    If bot.Velocity > bot.MaxVel Then bot.Velocity = bot.MaxVel
  ElseIf nWPDist > leg(g_nLegNum).Width * 2 Then
'    bot.Velocity = bot.MaxVel / 2
    bot.Velocity = leg(g_nLegNum).Width / 15
    If bot.Velocity > bot.MaxVel Then bot.Velocity = bot.MaxVel

  ElseIf nWPDist > leg(g_nLegNum).Width Then
    'bot.Velocity = bot.MaxVel / 3
    bot.Velocity = leg(g_nLegNum).Width / 20
    If bot.Velocity > bot.MaxVel Then bot.Velocity = bot.MaxVel
    
  Else  'arrived at waypoint
    g_nLegNum = g_nLegNum + 1
    If g_nLegNum > g_nLastLegNum Then g_nLegNum = 1
  End If
  
  If nFront < 2000 Then bot.Velocity = bot.MaxVel / 5
  
End Sub

