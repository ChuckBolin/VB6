VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Reader"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   7800
      Width           =   1515
   End
   Begin VB.TextBox txtDead 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "frmMain.frx":0000
      Top             =   5580
      Width           =   9675
   End
   Begin VB.TextBox txtTotal 
      Height          =   345
      Left            =   7560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   360
      Width           =   795
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5940
      TabIndex        =   10
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "&Check mailbox"
      Height          =   375
      Left            =   4500
      TabIndex        =   9
      Top             =   7800
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2760
      Width           =   9675
   End
   Begin VB.Frame Frame4 
      Caption         =   "Messages"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   9675
      Begin ComctlLib.ListView lvMessages 
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Password:"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Name:"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "cbolin"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remote Host:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "mail.dycon.com"
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mblnDownload As Boolean

Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

Private m_State         As POP3States

Private m_oMessage      As CMessage
Private m_colMessages   As New CMessages
'

Private Sub cmdCheckMail_Click()
    
    'Check the emptiness of all the text fields except for the txtBody
    For Each c In Controls
        If TypeOf c Is TextBox And c.Name <> "txtBody" Then
            If Len(c.Text) = 0 Then
                MsgBox c.Name & " can't be empty", vbCritical
                Exit Sub
            End If
        End If
    Next
    '
    m_State = POP3_Connect
    Winsock1.Close
    Winsock1.LocalPort = 0
    Winsock1.Connect txtHost, 110

End Sub

Private Sub cmdDel_Click()
    Unload Me
End Sub

Private Sub cmdStop_Click()
  If mblnDownload = True Then
    mblnDownload = False
    m_State = POP3_STAT
  Else
    mblnDownload = True
    m_State = POP3_RETR
  End If
End Sub

Private Sub lvMessages_ItemClick(ByVal Item As ComctlLib.ListItem)
    txtBody = m_colMessages(Item.Key).MessageBody
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    '
    'Save the received data into strData variable
    Winsock1.GetData strData
    'Debug.Print strData

    
    If Left$(strData, 1) = "+" Or m_State = POP3_RETR Then
        Select Case m_State
            Case POP3_Connect
                intMessages = 0
                m_State = POP3_USER
                Winsock1.SendData "USER " & txtUserName & vbCrLf
                'Debug.Print "USER " & txtUserName
            Case POP3_USER
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & txtPassword & vbCrLf
                'Debug.Print "PASS " & txtPassword
            Case POP3_PASS
                m_State = POP3_STAT
                Winsock1.SendData "STAT" & vbCrLf
                'Debug.Print "STAT"
                mblnDownload = True
            Case POP3_STAT
                intMessages = CInt(Mid$(strData, 5, InStr(5, strData, " ") - 5))
                txtTotal.Text = intMessages                  '<<<<<<<<<<<<< total emails
                If intMessages > 0 Then
                    m_State = POP3_RETR
                    intCurrentMessage = intCurrentMessage + 1
                    Winsock1.SendData "RETR 1" & vbCrLf
                    'Debug.Print "RETR 1"
                Else
                    m_State = POP3_QUIT
                    Winsock1.SendData "QUIT" & vbCrLf
                    'Debug.Print "QUIT"
                    MsgBox "You have no mail.", vbInformation
                End If
            Case POP3_RETR
                strBuffer = strBuffer & strData
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    Set m_oMessage = New CMessage
                    m_oMessage.CreateFromText strBuffer  '<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    Open App.Path & "\subject.txt" For Append As #1
                      Print #1, CStr(intCurrentMessage) & ": " & m_oMessage.Subject
                    Close #1
                    Open App.Path & "\from.txt" For Append As #1
                      Print #1, CStr(intCurrentMessage) & ": " & m_oMessage.From
                    Close #1
                    
                    
                    txtBody.Text = txtBody.Text & CStr(intCurrentMessage) & ": " & m_oMessage.Subject & ", " & vbCrLf
                                                                'm_oMessage.From & ", " & _
                                                               '  m_oMessage.MessageID & vbCrLf
                    
                   ' m_colMessages.Add m_oMessage, m_oMessage.MessageID
                    Dim strSub As String
                    strSub = LCase(m_oMessage.Subject)  'convert subject to lower case
                    Set m_oMessage = Nothing
                    strBuffer = ""
                    If intCurrentMessage = intMessages Then
                        m_State = POP3_QUIT
                        Winsock1.SendData "QUIT" & vbCrLf
                        'Debug.Print "QUIT"
                    Else
                        intCurrentMessage = intCurrentMessage + 1
                        frmMain.Caption = intCurrentMessage
                        '*************************************************
                        'my filter and delete go here
                       
                        'If InStr(1, strSub, "business cards") Then             'sample filter word
                        '  Winsock1.SendData "DELE " & CStr(intCurrentMessage - 1) & vbCrLf
                        '   txtDead.Text = txtDead.Text & strSub & vbCrLf
                        'End If
                        'Change current state of session
                        m_State = POP3_RETR
                        Winsock1.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        'Debug.Print "RETR " & intCurrentMessage
                    End If
                End If
                
            Case POP3_QUIT
                Winsock1.Close
        End Select
    Else
            Winsock1.Close
            MsgBox "POP3 Error: " & strData, _
            vbExclamation, "POP3 Error"
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    MsgBox "Winsock Error: #" & Number & vbCrLf & _
            Description
            
End Sub

Private Sub ListMessages()

    Dim oMes As CMessage
    Dim lvItem As ListItem
    
    For Each oMes In m_colMessages
        Set lvItem = lvMessages.ListItems.Add
        lvItem.Key = oMes.MessageID
        lvItem.Text = oMes.From
        lvItem.SubItems(1) = oMes.Subject
        lvItem.SubItems(2) = oMes.SendDate
        lvItem.SubItems(3) = oMes.Size
    Next
    
End Sub
