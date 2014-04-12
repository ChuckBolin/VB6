Attribute VB_Name = "Auto"
Option Explicit

'*****************************************************************
' This is the Autonomous routine. Its purpose is to modify two
' variables:
' Inputs:   bot.X
'           bot.Y
'           bot.Direction
'           dr.X
'           dr.Y
'           dr.Direction
'
'           leg(1).X1
'           leg(1).Y1
'           leg(1).X2
'           leg(1).Y2
'           leg(1).Orientation
'           leg(1).Width
'
' Outputs: bot.Turn
'          bot.velocity
'
' Try using these functions with bot position and waypoint info (leg())
'  GetTargetDistance2D ()
'  GetTargetDirection2D()
'  GetAngularDifference
'*****************************************************************

'October 12, 2005 by Chuck Bolin
Public Sub Autonomous()
  Dim nWPDir As Single
  Dim nWPDist As Single
  Dim nDirDiff As Single 'difference between nWPDir and bot.dir
  
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
    
  'determine turning amount
  If nDirDiff > 0 Then
    If bot.Velocity <> 0 Then bot.Turn = (nDirDiff / bot.Velocity)
  ElseIf nDirDiff < 0 Then
    If bot.Velocity <> 0 Then bot.Turn = (nDirDiff / bot.Velocity)
  Else
    bot.Turn = 0
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
End Sub

