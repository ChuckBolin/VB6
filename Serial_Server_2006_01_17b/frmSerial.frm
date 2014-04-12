VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSerial 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard Server v0.1"
   ClientHeight    =   975
   ClientLeft      =   10020
   ClientTop       =   1515
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2835
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   26
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   600
      Top             =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serial Data Status"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shpRec 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Basic Dashboard Viewer - Written by Chuck Bolin, July 2004
' Updated: January 2006     This uses COMM1 and monitors RC data
' only. Source code must be changed to get OI data
' Needs to controls:
'   MSWINSCK.OCX
'   MSCOMM32.OCX
' Remember to set this property.
' MSComm1.RThreshold = 26
'*************************************************************
Option Explicit

'variable declarations
Private receiveData As Boolean  'receive data if true
Private modeOI As Boolean 'true = OI data, false = RC data
Private countFrames As Long
Private stringFrame1 As String
Private stringFrame2 As String
Private stringFrame3 As String

'comm port 1 is used to connect to the dashboard
Private Sub connectToDashboard()
  Dim sRec As String
  If MSComm1.PortOpen = False Then
    MSComm1.CommPort = 1
    MSComm1.Settings = "19200,n,8,1"
    MSComm1.PortOpen = True
  End If
  receiveData = True
  sRec = MSComm1.Input  'forces initial input to clear
End Sub

'start listening as soon as program loads
Private Sub Form_Load()
  On Error Resume Next
  If Not App.PrevInstance = True Then
    Winsock1.LocalPort = 1412
    Winsock1.Listen
  End If
  receiveData = False
  modeOI = False
End Sub

'gets data request from client
Private Sub MSComm1_OnComm()
  On Error Resume Next
  Select Case MSComm1.CommEvent
    Case comEvReceive
        If receiveData = True Then
          ConvertS MSComm1.Input
        End If
   End Select
End Sub

'process frame data and generates string of 26 values
Private Sub ConvertS(sRec As String)
  Dim nLen As Integer
  Dim X As Integer
  Dim sNum As String
  Dim sData As String
  Dim bC7 As Boolean
  Dim bA4 As Boolean
  Dim nA4 As Integer
  Dim nC7 As Integer
  Dim nFrameNum As Integer 'frame number 1,2,3
    
  Static nPacket As Integer
  Static countMore As Long
  
  'length of packet must be 26 bytes to be valid
  nLen = Len(sRec)
  If nLen <> 26 Then
    shpRec.FillColor = RGB(0, 130, 0)
    Exit Sub
  Else
    shpRec.FillColor = RGB(0, 255, 0)
  End If
  countFrames = countFrames + 1
  If countFrames > 39 Then
    countFrames = 0
    countMore = countMore + 1
    frmSerial.Caption = countMore
  End If
  'frmSerial.Caption = countFrames
  'frmSerial.Caption = Len(sRec)
  
  
  'this is OI packet data...all packets are assumed to
  'be from the Operator Interface
  
    '1st two bytes must be 255 to be valid
    If Asc(Mid(sRec, 1, 1)) = 255 And Asc(Mid(sRec, 2, 1)) = 255 Then
         
      'clear boolean states of two bits and redefine
      'in order to determine packet number 1,2 or 3
      bC7 = False
      bA4 = False
      nC7 = Asc(Mid(sRec, 12, 1)) And &H80
      nA4 = Asc(Mid(sRec, 8, 1)) And &H10
      If nC7 > 0 Then bC7 = True
      If nA4 > 0 Then bA4 = True
            
      'construct string of 26 characters to fit inside text boxes
      For X = 1 To nLen
        sData = sData & CStr((Asc(Mid(sRec, X, 1)))) & vbCrLf
      Next X
      
      'display data in respective textbox
      'FRAME 1
      If bA4 = False And bC7 = False Then
        stringFrame1 = sData
        'txtRec(0).Text = sData
        nFrameNum = 1
      
      'FRAME 2
      ElseIf bA4 = True And bC7 = False Then
        stringFrame2 = sData
        'txtRec(1).Text = sData
        nFrameNum = 2
      'FRAME 3
      ElseIf bA4 = True And bC7 = True Then
        stringFrame3 = sData
        'txtRec(2).Text = sData
        nFrameNum = 3
      End If
    End If
  
End Sub

'every 500 mSec, the comm port is checked. Attempts to
'reconnect will happen if there is no connection
Private Sub Timer1_Timer()
  If MSComm1.PortOpen = False Then connectToDashboard
End Sub

'send frame1 to client
Private Sub sendFrame1()
  Winsock1.SendData stringFrame1
End Sub

'send frame2 to client
Private Sub sendFrame2()
  Winsock1.SendData stringFrame2
End Sub

'send frame3 to client
Private Sub sendFrame3()
  Winsock1.SendData stringFrame3
End Sub

'initial connection to client
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
  On Error Resume Next
  If Winsock1.State <> sckClosed Then Winsock1.Close
  Winsock1.Accept requestID
End Sub

'process incoming requests
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  On Error Resume Next
  Dim str As String
  Winsock1.GetData str
  
  'sendFrame1
  
  'frmSerial = str
  If str = "getFrame1" Then
    sendFrame1
  ElseIf str = "getFrame2" Then
    sendFrame2
  ElseIf str = "getFrame3" Then
    sendFrame3
  End If
End Sub


