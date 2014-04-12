VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSecond 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2610
      Top             =   9210
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   26
      Left            =   2160
      Top             =   9210
   End
   Begin VB.Frame fraOI 
      BackColor       =   &H00000000&
      Caption         =   "Operator Interface"
      ForeColor       =   &H00FFFFFF&
      Height          =   4245
      Left            =   60
      TabIndex        =   6
      Top             =   5520
      Width           =   1995
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Digital Display"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   975
         Left            =   150
         TabIndex        =   14
         Top             =   2850
         Width           =   1785
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Switch 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   13
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Switch 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   12
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Switch 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   11
         Top             =   1860
         Width           =   1065
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Relay 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   10
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Relay 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   9
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "PWM 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         TabIndex        =   8
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         Caption         =   "PWM 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   300
         Width           =   1065
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   10
         Left            =   510
         Shape           =   3  'Circle
         Top             =   2550
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   9
         Left            =   510
         Shape           =   3  'Circle
         Top             =   2190
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   8
         Left            =   510
         Shape           =   3  'Circle
         Top             =   1830
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   7
         Left            =   510
         Shape           =   3  'Circle
         Top             =   1470
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   180
         Shape           =   3  'Circle
         Top             =   1470
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   510
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   180
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   3
         Left            =   510
         Shape           =   3  'Circle
         Top             =   690
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   180
         Shape           =   3  'Circle
         Top             =   690
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   510
         Shape           =   3  'Circle
         Top             =   300
         Width           =   255
      End
      Begin VB.Shape shpLED 
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   180
         Shape           =   3  'Circle
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.Frame fraCode 
      Caption         =   "Autonomous Code"
      Height          =   5445
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   5355
      Begin VB.CommandButton cmdCompile 
         Caption         =   "&Compile/Load"
         Height          =   375
         Left            =   4080
         TabIndex        =   32
         Top             =   4950
         Width           =   1185
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4665
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame fraField 
      Caption         =   "Field"
      Height          =   5445
      Left            =   5460
      TabIndex        =   2
      Top             =   30
      Width           =   5445
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   5000
         Left            =   390
         ScaleHeight     =   -54
         ScaleMode       =   0  'User
         ScaleTop        =   54
         ScaleWidth      =   54
         TabIndex        =   3
         Top             =   390
         Width           =   5000
         Begin VB.Shape shpRobot 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   228
            Left            =   840
            Shape           =   3  'Circle
            Top             =   3810
            Width           =   228
         End
      End
      Begin VB.Label lblCoord 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1860
         TabIndex        =   31
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   9660
      TabIndex        =   1
      Top             =   9240
      Width           =   1245
   End
   Begin VB.Frame fraAuto 
      Caption         =   "Autonomous Controls"
      Height          =   3615
      Left            =   2130
      TabIndex        =   0
      Top             =   5550
      Width           =   8775
      Begin VB.TextBox txtInter 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   33
         Top             =   150
         Width           =   5415
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   7530
         TabIndex        =   29
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 14"
         Height          =   285
         Index           =   13
         Left            =   150
         TabIndex        =   28
         Top             =   3000
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 13"
         Height          =   285
         Index           =   12
         Left            =   150
         TabIndex        =   27
         Top             =   2790
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 12"
         Height          =   285
         Index           =   11
         Left            =   150
         TabIndex        =   26
         Top             =   2580
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 11"
         Height          =   285
         Index           =   10
         Left            =   150
         TabIndex        =   25
         Top             =   2370
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 10"
         Height          =   285
         Index           =   9
         Left            =   150
         TabIndex        =   24
         Top             =   2160
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 9"
         Height          =   285
         Index           =   8
         Left            =   150
         TabIndex        =   23
         Top             =   1950
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 8"
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   22
         Top             =   1740
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 7"
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   21
         Top             =   1530
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 6"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   20
         Top             =   1320
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 5"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   19
         Top             =   1110
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 4"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   900
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 3"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   690
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 2"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   480
         Width           =   1635
      End
      Begin VB.CheckBox chkDigital 
         Caption         =   "RC Digital Input 1"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Digital Display"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1005
         Left            =   7530
         TabIndex        =   30
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileProg1 
         Caption         =   "Load Sample Program 1"
      End
      Begin VB.Menu mnuFileProg2 
         Caption         =   "Load Sample Program 2"
      End
      Begin VB.Menu mnuFileProg3 
         Caption         =   "Load Sample Program 3"
      End
      Begin VB.Menu mnuFileProg4 
         Caption         =   "Load Sample Program 4"
      End
      Begin VB.Menu mnuFileProg5 
         Caption         =   "Load Sample Program 5"
      End
      Begin VB.Menu mnuFileProg6 
         Caption         =   "Load Sample Program 6"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
' FIRST VIRTUAL MACHINE - Written by Chuck Bolin, Team 342
' April 2005
' This program allows students to write C code for a virtual robot
' and then runs the program to control the robot in autonomous
' mode. The program should be accompanied by a tutorial for teaching
' C programming. Students are not required to compile, link or load
' the program.  This robot never needs a battery change.
'******************************************************************
Option Explicit

Private Sub cmdCompile_Click()
  Dim sVM As String
  
  g_nMaxLines = 0
  ReDim g_sCode(1)
  ClearVariables
  txtInter = ""
  sVM = Translate(AlignBraces(RemoveCommentsWhitespaces(txtCode)))
  g_sVM = Split(sVM, vbCrLf)  'loads VM code into array
  txtInter = sVM
    
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdStart_Click()
  tmrUpdate.Enabled = True
  tmrSecond.Enabled = True
  cmdStart.Enabled = False
  g_nTimeLeft = 15
End Sub

Private Sub Form_Load()
  LoadVariables
  
  frmMain.Caption = g_sProgram & " - " & g_sVersion & " by " & g_sTeam
  mnuFileProg6_Click
  pic.ForeColor = vbWhite
  
  'hardcoded field dimensions
  pic.Line (0, 0)-(27, 54), , B
  
End Sub

Private Function GetFile(sFile As String) As String
  Dim sInput As String
  Dim sTest As String
  Dim nFile As Integer
  
  nFile = FreeFile()
  GetFile = ""
  sTest = Dir(sFile)
  If Len(sTest) < 1 Then
    MsgBox "File name " & sFile & " does not exist!", vbOKOnly, "Bad Filename"
    Exit Function
  End If
  
  Open sFile For Input As #nFile
    Do
      Input #nFile, sInput
      GetFile = GetFile & sInput & vbCrLf
    Loop Until EOF(nFile)
  Close #nFile

End Function

Private Sub mnuFileProg1_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample1.txt")
  txtCode = sCode
End Sub

Private Sub mnuFileProg2_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample2.txt")
  txtCode = sCode
End Sub

Private Sub mnuFileProg3_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample3.txt")
  txtCode = sCode
End Sub

Private Sub mnuFileProg4_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample4.txt")
  txtCode = sCode
End Sub

Private Sub mnuFileProg5_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample5.txt")
  txtCode = sCode
End Sub

Private Sub mnuFileProg6_Click()
  Dim sCode As String
  sCode = GetFile(App.Path & "\sample6.txt")
  txtCode = sCode
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblCoord.Caption = "X: " & Format(X, "##.#") & "   Y: " & Format(Y, "##.#")
End Sub

Private Sub tmrSecond_Timer()
  g_nTimeLeft = g_nTimeLeft - 1
  If g_nTimeLeft < 1 Then
    tmrUpdate.Enabled = False
    tmrSecond.Enabled = False
    lblTimer.Caption = CStr(g_nTimeLeft)
    cmdStart.Enabled = True
  Else
    lblTimer.Caption = CStr(g_nTimeLeft)
  End If

End Sub

'This updates OI display and robot on field each 26.2 mSec
Private Sub tmrUpdate_Timer()
  Dim i As Integer
  
  'turnoff all outputs
  InitializeRC
  
  'read inputs
  rc.rc_dig_in01 = chkDigital(0).Value
  rc.rc_dig_in02 = chkDigital(1).Value
  rc.rc_dig_in03 = chkDigital(2).Value
  rc.rc_dig_in04 = chkDigital(3).Value
  rc.rc_dig_in05 = chkDigital(4).Value
  rc.rc_dig_in06 = chkDigital(5).Value
  rc.rc_dig_in07 = chkDigital(6).Value
  rc.rc_dig_in08 = chkDigital(7).Value
  rc.rc_dig_in09 = chkDigital(8).Value
  rc.rc_dig_in10 = chkDigital(9).Value
  rc.rc_dig_in11 = chkDigital(10).Value
  rc.rc_dig_in12 = chkDigital(11).Value
  rc.rc_dig_in13 = chkDigital(12).Value
  rc.rc_dig_in14 = chkDigital(13).Value
 
  'initialize OI LEDs
  For i = 0 To 10
    shpLED(i).BackColor = vbWhite
  Next i
  
  'test code
  If rc.rc_dig_in01 = True Then shpLED(8).BackColor = vbGreen
  If rc.rc_dig_in02 = True Then shpLED(9).BackColor = vbGreen
  If rc.rc_dig_in03 = True Then shpLED(10).BackColor = vbGreen
  If rc.rc_dig_in04 = True Then shpLED(0).BackColor = vbGreen
  If rc.rc_dig_in05 = True Then shpLED(1).BackColor = vbRed
  If rc.rc_dig_in06 = True Then shpLED(2).BackColor = vbGreen
  If rc.rc_dig_in07 = True Then shpLED(3).BackColor = vbRed
  If rc.rc_dig_in08 = True Then shpLED(4).BackColor = vbGreen
  If rc.rc_dig_in09 = True Then shpLED(5).BackColor = vbRed
  If rc.rc_dig_in10 = True Then shpLED(6).BackColor = vbGreen
  If rc.rc_dig_in11 = True Then shpLED(7).BackColor = vbRed
  
  
  'update robot physical location and direction
  shpRobot.Left = robot.X - 1.25
  shpRobot.Top = robot.Y + 1.25
  
End Sub

'******************************************
' SAMPLE PROGRAM 1
'******************************************
Private Sub LoadSampleProgram1()
  txtCode = ""
  txtCode = txtCode & "/*"
  txtCode = txtCode & "Pwm1_green = 1;" & vbCrLf
  txtCode = txtCode & "Pwm2_green = 1;*/" & vbCrLf
  txtCode = txtCode & "Pwm1_red = 1;" & vbCrLf
  txtCode = txtCode & "Pwm2_red = 1;" & vbCrLf
  txtCode = txtCode & "Relay1_green = 1;" & vbCrLf
  txtCode = txtCode & "Relay2_green = 1;" & vbCrLf
  txtCode = txtCode & "/*Relay1_red = 1;" & vbCrLf
  txtCode = txtCode & "Relay2_red = 1;" & vbCrLf
  txtCode = txtCode & "Switch1_LED = 1;" & vbCrLf
  txtCode = txtCode & "*/"
  txtCode = txtCode & "//Switch2_LED = 1;//This is a test" & vbCrLf
  txtCode = txtCode & "Switch3_LED = 1;" & vbCrLf
End Sub

'******************************************
' SAMPLE PROGRAM 2
'******************************************
Private Sub LoadSampleProgram2()
  txtCode = ""
  txtCode = txtCode & "Switch1_LED = 1;" & vbCrLf
  txtCode = txtCode & "Switch2_LED = 1;" & vbCrLf
  txtCode = txtCode & "Switch3_LED = 1;" & vbCrLf
End Sub

'******************************************
' SAMPLE PROGRAM 3
'******************************************
Private Sub LoadSampleProgram3()
  txtCode = ""
  txtCode = txtCode & "char i;" & vbCrLf
  txtCode = txtCode & "unsigned char j;" & vbCrLf
  txtCode = txtCode & "int k = 3;" & vbCrLf
  txtCode = txtCode & "unsigned int l = 4;" & vbCrLf
  txtCode = txtCode & "long m = 5;" & vbCrLf
  txtCode = txtCode & "unsigned long n = 6;" & vbCrLf
  txtCode = txtCode & "for(i = 0; i<10;i++){" & vbCrLf
  txtCode = txtCode & "for(j = 0; j<5;j++){" & vbCrLf
  txtCode = txtCode & "   k++;}" & vbCrLf

  txtCode = txtCode & "   m++;}" & vbCrLf
  txtCode = txtCode & "i=7;" & vbCrLf
  txtCode = txtCode & "j=8;" & vbCrLf
  

End Sub

'******************************************
' SAMPLE PROGRAM 4
'******************************************
Private Sub LoadSampleProgram4()
  txtCode = ""
  txtCode = txtCode & "char i;" & vbCrLf
  txtCode = txtCode & "unsigned char j;" & vbCrLf
  txtCode = txtCode & "int k = 3;" & vbCrLf
  txtCode = txtCode & "unsigned int l = 4;" & vbCrLf
  txtCode = txtCode & "long m = 5;" & vbCrLf
  txtCode = txtCode & "unsigned long n = 6;" & vbCrLf
  txtCode = txtCode & "for(i = 0; i<10;i++){" & vbCrLf
  txtCode = txtCode & "   k++;}" & vbCrLf
  txtCode = txtCode & "i=7;" & vbCrLf
  txtCode = txtCode & "j=8;" & vbCrLf
  

End Sub

