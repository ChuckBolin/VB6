VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetHostByAddress"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIpAddress 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "txtIpAddress"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtHostName 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "txtHostName"
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP address:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   810
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGet_Click()
    '
    '-----------------------------------------
    'address as a Long value returned by
    'the inet_addr function
    Dim lngInetAdr      As Long
    '
    'pointer to the HOSTENT structure
    Dim lngPtrHostEnt   As Long
    '
    'host name we are looking for
    Dim strHostName     As String
    '
    'HOSTENT structure
    Dim udtHostEnt      As HOSTENT
    '
    'address in dotted notation
    Dim strIpAddress    As String
    '-----------------------------------------
    '
    txtHostName.Text = ""
    '
    strIpAddress = Trim$(txtIpAddress.Text)
    '
    'Convert the IP address string to Long
    lngInetAdr = inet_addr(strIpAddress)
    '
    'if the IP address is in wrong format
    'the inet_addr function returns INADDR_NONE value
    If lngInetAdr = INADDR_NONE Then
        '
        ShowErrorMsg (Err.LastDllError)
        '
    Else
        '
        '## Retrieve host name
        '
        'Get the pointer to the HostEnt structure
        lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, PF_INET)
        '
        'if the gethostbyaddr function can't find teh host,
        'it returns a NULL pointer
        If lngPtrHostEnt = 0 Then
            '
            ShowErrorMsg (Err.LastDllError)
            '
        Else
            '
            'Copy data into the HostEnt structure
            RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
            '
            'Prepare the buffer to receive a string
            strHostName = String(256, 0)
            '
            'Copy the host name into the strHostName variable
            RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256
            '
            'Cut received string by first chr(0) character
            strHostName = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
            '
            'Return the found value
            txtHostName.Text = strHostName
            '
        End If
        '
    End If
    '
End Sub

Private Sub Form_Load()
    '
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
    Dim lngType        As Long
    Dim lngProtocol    As Long
    '
    'start up winsock service
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    '
    If lngRetVal <> 0 Then
        '
        '
        Select Case lngRetVal
            Case WSASYSNOTREADY
                strErrorMsg = "The underlying network subsystem is not " & _
                    "ready for network communication."
            Case WSAVERNOTSUPPORTED
                strErrorMsg = "The version of Windows Sockets API support " & _
                    "requested is not provided by this particular " & _
                    "Windows Sockets implementation."
            Case WSAEINVAL
                strErrorMsg = "The Windows Sockets version specified by the " & _
                    "application is not supported by this DLL."
        End Select
        '
        MsgBox strErrorMsg, vbCritical
        '
    End If
    '
    txtHostName.Text = ""
    txtIpAddress.Text = ""
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WSACleanup
End Sub

Private Sub ShowErrorMsg(lngError As Long)
    '
    Dim strMessage As String
    '
    Select Case lngError
        Case WSANOTINITIALISED
            strMessage = "A successful WSAStartup call must occur " & _
                         "before using this function."
        Case WSAENETDOWN
            strMessage = "The network subsystem has failed."
        Case WSAHOST_NOT_FOUND
            strMessage = "Authoritative answer host not found."
        Case WSATRY_AGAIN
            strMessage = "Nonauthoritative host not found, or server failure."
        Case WSANO_RECOVERY
            strMessage = "A nonrecoverable error occurred."
        Case WSANO_DATA
            strMessage = "Valid name, no data record of requested type."
        Case WSAEINPROGRESS
            strMessage = "A blocking Windows Sockets 1.1 call is in " & _
                         "progress, or the service provider is still " & _
                         "processing a callback function."
        Case WSAEFAULT
            strMessage = "The name parameter is not a valid part of " & _
                         "the user address space."
        Case WSAEINTR
            strMessage = "A blocking Windows Socket 1.1 call was " & _
                         "canceled through WSACancelBlockingCall."
    End Select
    '
    MsgBox strMessage, vbExclamation
    '
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

