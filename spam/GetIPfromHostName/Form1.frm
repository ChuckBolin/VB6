VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   300
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   300
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1620
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   1020
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   4395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Const WS_VERSION_REQD As Long = &H101

Public Function GetIPFromHostName(ByVal sHostName As String) As String
Dim nbytes As Long
Dim ptrHosent As Long
Dim ptrName As Long
Dim ptrAddress As Long
Dim ptrIPAddress As Long
Dim sAddress As String
Dim WSAD As WSADATA

If WSAStartup(WS_VERSION_REQD, WSAD) = 0 Then
    sAddress = Space$(4)
    ptrHosent = gethostbyname(sHostName & vbNullChar)
    If ptrHosent <> 0 Then
        ptrAddress = ptrHosent + 12
        CopyMemory ptrAddress, ByVal ptrAddress, 4
        CopyMemory ptrIPAddress, ByVal ptrAddress, 4
        CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
    
        GetIPFromHostName = Str2IP(sAddress)
        
        WSACleanup
    End If
End If
End Function

Private Function Str2IP(addr As String) As String
Str2IP = CStr(Asc(addr)) & "." & _
            CStr(Asc(Mid$(addr, 2, 1))) & "." & _
            CStr(Asc(Mid$(addr, 3, 1))) & "." & _
            CStr(Asc(Mid$(addr, 4, 1)))
End Function

Private Sub Command1_Click()
  Text1.Text = GetIPFromHostName(Text1.Text)
End Sub
