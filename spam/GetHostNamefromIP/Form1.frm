VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   300
      TabIndex        =   2
      Top             =   1380
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   180
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2

Private Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  imaxsockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSADATA) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
  
  
  
Private Sub Command1_Click()

   Text2.Text = GetHostNameFromIP(Text1.Text)
   
End Sub

  
Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Private Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub


Private Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         
         Else
            MsgBox "Call to gethostbyaddr failed."
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else
         MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   Else
      MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function


