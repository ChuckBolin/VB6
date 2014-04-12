VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard Client v0.1"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFrame3 
      Height          =   5010
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtFrame2 
      Height          =   5010
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   5760
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   960
   End
   Begin VB.TextBox txtFrame1 
      Height          =   5010
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
End Sub

Private Sub Form_Load()
  frameNumber = 1
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  
  'If frameNumber = 1 Then
    Winsock1.SendData "getFrame1"
  'ElseIf frameNumber = 2 Then
   'Winsock1.SendData "getFrame2"
  'ElseIf frameNumber = 3 Then
   ' Winsock1.SendData "getFrame3"
  'End If
  
  'frameNumber = frameNumber + 1
  'If frameNumber > 3 Then frameNumber = 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  On Error Resume Next
  Dim str As String
  If bytesTotal > 0 Then
  Winsock1.GetData str
  frameCount = frameCount + 1
  frmClient.Caption = frameCount
  'If frameNumber = 1 Then
    txtFrame1 = str
  'ElseIf frameNumber = 2 Then
    'txtFrame2 = str
  'ElseIf frameNumber = 3 Then
    'txtFrame3 = str
  'End If
  End If
End Sub

