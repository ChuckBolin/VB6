VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard Client v0.1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   2407
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   -312
      ScaleMode       =   0  'User
      ScaleTop        =   312
      ScaleWidth      =   648
      TabIndex        =   6
      Top             =   2400
      Width           =   5000
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   5
         Left            =   2848
         Top             =   360
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   4
         Left            =   2239
         Top             =   0
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   3
         Left            =   1782
         Top             =   360
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   2
         Left            =   3000
         Top             =   960
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   1
         Left            =   2239
         Top             =   1560
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   450
         Index           =   0
         Left            =   1782
         Top             =   1680
         Width           =   457
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   324.066
         X2              =   324.066
         Y1              =   312
         Y2              =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   120
         Left            =   4905
         Top             =   1080
         Width           =   61
      End
   End
   Begin VB.Timer tmrRead 
      Interval        =   1000
      Left            =   2040
      Top             =   1800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdRC 
      Caption         =   "View RC Packets"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   240
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   960
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblY 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblX 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblTilt 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblPan 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' frmClient.frm - Written by Chuck Bolin, Thanks to Keral C. Patel for his
' tutorial.
'***************************************************************************
Option Explicit

Private frameNumber As Integer
Private frameCount As Long

Private Sub cmdConnect_Click()
  On Error Resume Next
  
  Winsock1.Connect txtIP.Text, "1412"
  Timer1.Enabled = True
  frmClient.Caption = "Monitoring "
End Sub

Private Sub cmdExit_Click()
  'Winsock1.SendData "close"
  'winsock1.Close
  End
End Sub

Private Sub cmdRC_Click()
  frmViewRC.Show
End Sub

Private Sub Form_Load()
  frameNumber = 1
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblX = FormatNumber(X, 0)
  lblY = FormatNumber(Y, 0)
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  
  If frameNumber = 1 Then
    Winsock1.SendData "getFrame1"
  ElseIf frameNumber = 2 Then
   Winsock1.SendData "getFrame2"
  ElseIf frameNumber = 3 Then
    Winsock1.SendData "getFrame3"
  End If
  
End Sub

Private Sub tmrRead_Timer()
  Dim radius As Single
  
  lblPan = frame1.Byte3
  lblTilt = frame1.Byte5
  
  'pic.Circle (Val(lblPan), Val(lblTilt)), 1
  'radius = 3 * (190 - Val(lblTilt))
  'pic.Cls
  'If radius > 0 Then pic.Circle (150, 190), radius
  'linPan.X1 = Val(lblPan) + 15
  'linPan.X2 = linPan.X1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 ' lblBytes.Caption = bytesTotal
  On Error Resume Next
  Static frameMore As Long
  Dim str As String
  If bytesTotal > 0 Then
    Winsock1.GetData str
    
    If frameMore = 0 Then
      frmClient.Caption = frameCount
    End If
    frameCount = frameCount + 1
    If frameCount > 39 Then
      frameCount = 0
      frameMore = frameMore + 1
      frmClient.Caption = frameMore
    End If
    
    If frameNumber = 1 Then
      'txtFrame1 = str
      loadFrame1 str
      'txtFrame2 = frame1.Byte1 & vbCrLf & frame1.Byte2
    ElseIf frameNumber = 2 Then
      'txtFrame2 = str
      loadFrame2 str
    ElseIf frameNumber = 3 Then
      'txtFrame3 = str
      loadFrame3 str
    End If
    frameNumber = frameNumber + 1
    If frameNumber > 3 Then frameNumber = 1
    
  End If
End Sub

