VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard Client v0.1"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   2610
   StartUpPosition =   3  'Windows Default
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
  End
End Sub

Private Sub cmdRC_Click()
  frmViewRC.Show
End Sub

Private Sub Form_Load()
  frameNumber = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
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

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  On Error Resume Next
  Static frameMore As Long
  Dim str As String
  If bytesTotal > 0 Then
    Winsock1.GetData str
    
    
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

