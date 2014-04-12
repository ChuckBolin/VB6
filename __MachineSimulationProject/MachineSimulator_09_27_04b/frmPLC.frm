VERSION 5.00
Begin VB.Form frmPLC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLC I/O"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4665
   Begin VB.Frame frmCPU 
      Caption         =   "CPU"
      Height          =   4695
      Left            =   1440
      TabIndex        =   39
      Top             =   120
      Width           =   1095
      Begin VB.Label Label6 
         Caption         =   "Clear Outputs"
         Height          =   405
         Left            =   480
         TabIndex        =   41
         Top             =   1200
         Width           =   555
      End
      Begin VB.Shape shpClearSw 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   165
         Left            =   180
         Top             =   1380
         Width           =   195
      End
      Begin VB.Shape shpClearOutline 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Height          =   345
         Left            =   150
         Top             =   1230
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Stop Run"
         Height          =   465
         Left            =   480
         TabIndex        =   40
         Top             =   630
         Width           =   375
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   180
         Shape           =   3  'Circle
         Top             =   360
         Width           =   195
      End
      Begin VB.Shape shpRunSw 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   165
         Left            =   180
         Top             =   810
         Width           =   195
      End
      Begin VB.Shape shpRunOutline 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Height          =   345
         Left            =   150
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   6360
      Top             =   3600
   End
   Begin VB.Frame frmPS 
      Caption         =   "Power Supply"
      Height          =   4695
      Left            =   90
      TabIndex        =   34
      Top             =   120
      Width           =   1305
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   330
         X2              =   450
         Y1              =   4080
         Y2              =   3930
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "OFF"
         Height          =   255
         Left            =   210
         TabIndex        =   38
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ON"
         Height          =   255
         Left            =   210
         TabIndex        =   37
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape shpPwrSw 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   300
         Shape           =   2  'Oval
         Top             =   2760
         Width           =   155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "5V"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "12 V"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape shpPLC5V 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   240
         Shape           =   3  'Circle
         Top             =   600
         Width           =   195
      End
      Begin VB.Shape shpPLC12V 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   240
         Shape           =   3  'Circle
         Top             =   360
         Width           =   195
      End
      Begin VB.Shape shpFuse 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00404040&
         Height          =   375
         Left            =   150
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   495
      End
   End
   Begin VB.Frame frmOutput 
      Caption         =   "Outputs"
      Height          =   4695
      Left            =   3600
      TabIndex        =   17
      Top             =   120
      Width           =   975
      Begin VB.Label lblOut 
         Caption         =   "16"
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   33
         Top             =   4320
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   15
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "15"
         Height          =   255
         Index           =   14
         Left            =   480
         TabIndex        =   32
         Top             =   4080
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   14
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "14"
         Height          =   255
         Index           =   13
         Left            =   480
         TabIndex        =   31
         Top             =   3840
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   13
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "13"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   30
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   12
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "12"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   29
         Top             =   3240
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   11
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "11"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   28
         Top             =   3000
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   10
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "10"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   27
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   9
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   26
         Top             =   2520
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   8
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   25
         Top             =   2160
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   7
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   24
         Top             =   1920
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   6
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   23
         Top             =   1680
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   5
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   4
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   3
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   2
         Left            =   240
         Shape           =   3  'Circle
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   1
         Left            =   240
         Shape           =   3  'Circle
         Top             =   600
         Width           =   195
      End
      Begin VB.Label lblOut 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape shpOut 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame frmInput 
      Caption         =   "Inputs"
      Height          =   4695
      Left            =   2580
      TabIndex        =   0
      Top             =   120
      Width           =   975
      Begin VB.Label lblIn 
         Caption         =   "16"
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   16
         Top             =   4320
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   15
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "15"
         Height          =   255
         Index           =   14
         Left            =   480
         TabIndex        =   15
         Top             =   4080
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   14
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "15"
         Height          =   255
         Index           =   13
         Left            =   480
         TabIndex        =   14
         Top             =   3840
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   13
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "13"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   13
         Top             =   3600
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   12
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "12"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   12
         Top             =   3240
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   11
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "11"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   11
         Top             =   3000
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   10
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "10"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   10
         Top             =   2760
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   9
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   9
         Top             =   2520
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   8
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   7
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   6
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   5
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   4
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   3
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   2
         Left            =   240
         Shape           =   3  'Circle
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   1
         Left            =   240
         Shape           =   3  'Circle
         Top             =   600
         Width           =   195
      End
      Begin VB.Label lblIn 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape shpIn 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   360
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub frmCPU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  'toggling run/stop switch
  If X > shpRunOutline.Left And X < shpRunOutline.Left + shpRunOutline.Width Then
    If Y > shpRunOutline.Top And Y < shpRunOutline.Top + shpRunOutline.Height Then
      If g_uPLC.RunSwitch = True Then
        g_uPLC.RunSwitch = False
      Else
        g_uPLC.RunSwitch = True
      End If
    End If
  End If
  
  'toggling clear output switch
  If X > shpClearOutline.Left And X < shpClearOutline.Left + shpClearOutline.Width Then
    If Y > shpClearOutline.Top And Y < shpClearOutline.Top + shpClearOutline.Height Then
      If g_uPLC.ClearOutputSwitch = True Then
        g_uPLC.ClearOutputSwitch = False
      Else
        g_uPLC.ClearOutputSwitch = True
      End If
    End If
  End If

End Sub

Private Sub frmPS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X > shpPwrSw.Left And X < shpPwrSw.Left + shpPwrSw.Width Then
    If Y > shpPwrSw.Top And Y < shpPwrSw.Top + shpPwrSw.Height Then
      If g_uPLC.PowerSwitch = True Then
        g_uPLC.PowerSwitch = False
      Else
        g_uPLC.PowerSwitch = True
      End If
    End If
  End If
End Sub

'updates all visiualization of PLC
Private Sub tmrUpdate_Timer()
  Dim i As Integer

  'update PLC power supply indicators
  If g_uPLC.Power12VIndicator = True Then
    shpPLC12V.BackColor = RGB(0, 255, 0)
  Else
    shpPLC12V.BackColor = RGB(0, 155, 0)
  End If
  
  If g_uPLC.Power5VIndicator = True Then
    shpPLC5V.BackColor = RGB(0, 255, 0)
  Else
    shpPLC5V.BackColor = RGB(0, 155, 0)
  End If
  
  'update PLC power switch
  If g_uPLC.PowerSwitch = True Then
    shpPwrSw.Top = 2760
  Else
    shpPwrSw.Top = 3100
  End If
  
  'run switch on CPU
  If g_uPLC.RunSwitch = True Then
    shpRunSw.Top = 810
  Else
    shpRunSw.Top = 665
  End If
  
  'run/stop indicator
  If g_uPLC.RunIndicator = True Then
    shpRun.BackColor = RGB(155, 0, 0)
  Else
    shpRun.BackColor = RGB(255, 0, 0)
  End If
    
  'clear output switch
  If g_uPLC.ClearOutputSwitch = True Then
    shpClearSw.Top = 1260
  Else
    shpClearSw.Top = 1380
  End If
  
  'set inputs and outputs
  For i = 1 To IO.GetMaxInputBits
    If IO.GetInput(i) = True Then
      shpIn(i - 1).BackColor = RGB(0, 255, 0)
    Else
      shpIn(i - 1).BackColor = RGB(0, 155, 0)
    End If
  Next i
  
  'show outputs only if clear switch is not pressed and plc is running
  For i = 1 To IO.GetMaxoutputBits
    If IO.GetOutput(i) = True Then
      shpOut(i - 1).BackColor = RGB(0, 255, 0)
    Else
      shpOut(i - 1).BackColor = RGB(0, 155, 0)
    End If
  Next i
  

  
End Sub
