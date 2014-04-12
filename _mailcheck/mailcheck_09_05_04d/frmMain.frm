VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   1815
   ClientTop       =   615
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11295
   Begin VB.TextBox txtViewAll 
      Height          =   3555
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   4680
      Width           =   9495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Spam Count:"
      Height          =   615
      Left            =   7200
      TabIndex        =   17
      Top             =   120
      Width           =   2235
      Begin VB.TextBox txtTotal2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   19
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSpamTotal 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   195
         Left            =   1620
         TabIndex        =   21
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "of"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Winsock Status"
      Height          =   1455
      Left            =   9660
      TabIndex        =   12
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Closed"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Open"
         Height          =   195
         Left            =   540
         TabIndex        =   13
         Top             =   360
         Width           =   555
      End
      Begin VB.Shape shpClose 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   180
         Shape           =   3  'Circle
         Top             =   720
         Width           =   195
      End
      Begin VB.Shape shpOpen 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   180
         Shape           =   3  'Circle
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame fraSubject 
      Caption         =   "Subject Lines"
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   9495
      Begin VB.CheckBox chkViewList 
         Caption         =   "View Subjects"
         Height          =   255
         Left            =   7800
         TabIndex        =   16
         Top             =   180
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ListBox lstSubject 
         Height          =   2400
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   9255
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11040
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Password:"
      Height          =   615
      Left            =   4140
      TabIndex        =   4
      Top             =   120
      Width           =   1395
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "agency"
         ToolTipText     =   "Ex: ********"
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Name:"
      Height          =   615
      Left            =   2460
      TabIndex        =   2
      Top             =   120
      Width           =   1635
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "cbolin"
         ToolTipText     =   "Ex: johndoe"
         Top             =   240
         Width           =   1395
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
         ToolTipText     =   "Ex: mail.server.com"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Total Emails:"
      Height          =   615
      Left            =   5580
      TabIndex        =   6
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtCurrent 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtTotal 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "of"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   300
         Width           =   135
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnectCheck 
         Caption         =   "&Check Mailbox"
      End
      Begin VB.Menu mnuConnectProcess 
         Caption         =   "&Process Email"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnectExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Con&figuration"
      Begin VB.Menu mnuConfigFilter 
         Caption         =   "Spam &Filter"
      End
      Begin VB.Menu mnuConfigDump 
         Caption         =   "&Dump to File"
      End
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupView 
         Caption         =   "&View Email"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "&Delete Email"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
' M A I L C H E C K E R - December 2003
' Modified by Chuck B
' Original source downloaded from vbforums.com
' Original author is unknown. Let me know if you
' know who it is so they can get proper credit.
'**************************************************************************
Option Explicit

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
Private mblnDownload As Boolean
Private mblnSending As Boolean
Private mintDelete As Integer 'email to delete
Private m_Status As Integer
Private mblnDownloadComplete As Boolean 'true means it is complete


Private Sub cmdCheckMail_Click()

End Sub

Private Sub chkViewList_Click()
  If chkViewList.Value = vbChecked Then
    'frmMain.Height = 5160
    lstSubject.Visible = True
   ' lstStatus.Visible = True
    fraSubject.Height = 735 + lstSubject.Height + 250
    'cmdExit.Top = 4000
  Else
    'frmMain.Height = 2750
    lstSubject.Visible = False
   ' lstStatus.Visible = False
    fraSubject.Height = 735
    'cmdExit.Top = 1680
  End If
End Sub

'exits program
Private Sub cmdExit_Click()
 
End Sub


Private Sub cmdStop_Click()
   ' cmdStop.Enabled = False
   ' cmdCheckMail.Enabled = True
    'm_State = POP3_QUIT
   ' intCurrentMessage = 0
End Sub



Private Sub Command3_Click()

End Sub

'Load form
Private Sub Form_Load()
  
  'loads module variables
  LoadGlobalVariables
    
  'loads controls on interface
  frmMain.Caption = gstrProgram & " " & gstrVersion & " - " & gstrDate
  WinsockStatus False
  fraSubject.Height = 735
  chkViewList_Click
  mnuConfigDump_Click
  mnuConnectCheck_Click
  
  'loads spam word array from file
  'Open App.Path & "\spamwords.txt" For Append As #1
  '  If LOF(1) < 1 Then
  '    Close 1
  '    Exit Sub
  '  End If
  'Close
  'Dim X As Integer
  'Open App.Path & "\spamwords.txt" For Input As #1
  '  Do
  '  X = X + 1
  '  ReDim Preserve word(X)
  '  Line Input #1, word(X)
  '  Loop Until EOF(1)
  'Close #1
  
End Sub

Private Sub lstSubject_Click()
  Dim X, s As Long
  
 'determine list item selected
  For X = 0 To lstSubject.ListCount - 1
    If lstSubject.Selected(X) = True Then
      s = lstSubject.ListIndex + 1
      Exit For
    End If
  Next X
  
  'update text box
 ' txtDelete.Text = s
'  cmdDelete.Enabled = True
'  txtDeleteStatus.Text = ""
 ' txtDeleteStatus.Text = FilterSubject(CInt(s))
'  Text1.Text = em(s).subject
  

End Sub


Private Sub lstSubject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    If gblnDownloadComplete = True Then
      mnuPopupDelete.Enabled = True
    Else
      mnuPopupDelete.Enabled = False
    End If
    If lstSubject.ListCount > 0 Then PopupMenu Popup
  End If
End Sub

'this options allows all emails to be dumped to a text file "emaildump.txt"
Private Sub mnuConfigDump_Click()
  If mnuConfigDump.Checked = True Then
    mnuConfigDump.Checked = False
    gblnDumpToFile = False
  Else
    mnuConfigDump.Checked = True
    gblnDumpToFile = True
  End If
End Sub

Private Sub mnuConfigFilter_Click()
  frmFilter.Show
End Sub

'**************************************************************************
'  C H E C K   M A I L    M E N U
'  Initiates log on sequence with mail server
'**************************************************************************
Private Sub mnuConnectCheck_Click()
 
  gblnDownloadComplete = False
  mnuConnectProcess.Enabled = False
  mnuPopupDelete.Enabled = False
  mnuConfig.Enabled = False
  
  txtCurrent.Text = 0
  txtTotal.Text = 0
  gintTotalEmails = 0
  lstSubject.Clear
  ReDim em(0)
  
  'verify critical text fields are completed
    If Len(txtHost.Text) < 1 Then
      MsgBox "Host field is empty!"
      Exit Sub
    End If
    If Len(txtUserName.Text) < 1 Then
      MsgBox "Username  field is empty!"
      Exit Sub
    End If
    If Len(txtPassword.Text) < 1 Then
      MsgBox "Password field is empty!"
      Exit Sub
    End If
    
    'load variables for filtering use
    g_sUserDomain = "dycon.com"  '<<<<<<<<< note change this
    g_sUserName = LCase(txtUserName.Text) & "@" & LCase(g_sUserDomain)
    
  
    
    'configure interface controls
    'cmdStop.Enabled = True
  '  cmdCheckMail.Enabled = False
    lstSubject.Clear
    
    'sets initial program state, closes winsock in case it is open and sets up
    'winsock for POP3 protocol
    m_State = POP3_Connect
    Winsock1.Close
    Winsock1.LocalPort = 0
    Winsock1.Connect txtHost, 110
    

End Sub

Private Sub mnuConnectExit_Click()
  Winsock1.Close
  Unload frmReview
  Unload frmFilter
  Unload frmAddWord
  Unload Me
End Sub

Private Sub mnuConnectProcess_Click()
  Dim X As Integer
  For X = 1 To UBound(em)
    If em(X).delete = True And em(X).dead = False Then
      Winsock1.SendData "DELE " & CStr(X) & vbCrLf
      lstSubject.List(X - 1) = "[ DELETED FROM MAIL SERVER ]"
      em(X).dead = True
    End If
  Next X
  Winsock1.SendData "QUIT" & vbCrLf
End Sub

Private Sub mnuPopupDelete_Click()
  mintDelete = lstSubject.ListIndex + 1
'  If mintDelete > -1 Then Winsock1.SendData "DELE " & CStr(mintDelete) & vbCrLf
  If em(mintDelete).delete = True And em(mintDelete).dead = False Then 'already marked for delete
    em(mintDelete).delete = False
    lstSubject.List(mintDelete - 1) = "[ SPAM ]: " & CStr(mintDelete) & ":  (" & CStr(em(mintDelete).bytes_total) & ")  " & CStr(em(mintDelete).subject)
  Else                                               'NOT marked for delete
    em(mintDelete).delete = True
    lstSubject.List(mintDelete - 1) = "[DELETE]  [ SPAM ]: " & CStr(mintDelete) & ":  (" & CStr(em(mintDelete).bytes_total) & ")  " & CStr(em(mintDelete).subject)
  End If
End Sub

Private Sub mnuPopupView_Click()
  gintEmailToReview = lstSubject.ListIndex + 1
  If gintEmailToReview > -1 Then frmReview.Show
End Sub

'**************************************************************************
'  D A T A   A R R I V A L
'  Manage all incoming data
'**************************************************************************
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   
    mblnSending = False
    Dim strData As String
    Dim intCode As String 'indicates spam code
    Dim strMsg As String 'message to indicate SPAM or OK
    Dim intAccCode As Integer 'accumulated codes for one email...several reasons why it is spam
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    
    'Save the received data into strData variable
    Winsock1.GetData strData
    DoEvents
    If Left$(strData, 1) = "+" Or m_State = POP3_RETR Then
        Select Case m_State
            Case POP3_Connect
                intMessages = 0
                m_State = POP3_USER
                Winsock1.SendData "USER " & txtUserName & vbCrLf
                txtStatus.Text = "User Login"
            Case POP3_USER
                m_State = POP3_PASS
                Winsock1.SendData "PASS " & txtPassword & vbCrLf
                txtStatus.Text = "Password"
            Case POP3_PASS
                m_State = POP3_STAT
                Winsock1.SendData "STAT" & vbCrLf
                mblnDownload = True
                txtStatus.Text = "Stat"
            Case POP3_STAT
                WinsockStatus True
                intMessages = CInt(Mid$(strData, 5, InStr(5, strData, " ") - 5))
                txtTotal.Text = intMessages                  '<<<<<<<<<<<<< total emails
                txtTotal2.Text = intMessages
                gintTotalEmails = intMessages
                If intMessages > 0 Then
                  If mintDelete > 0 Then
                    m_State = POP3_DELE
                  Else
                    m_State = POP3_RETR
                    intCurrentMessage = intCurrentMessage + 1
                    Winsock1.SendData "RETR 1" & vbCrLf
                    txtStatus.Text = "Retrieving " & CStr(intCurrentMessage + 1)
                  End If
                Else
                    m_State = POP3_QUIT
                    Winsock1.SendData "QUIT" & vbCrLf
                    MsgBox "You have no mail.", vbInformation
                    txtStatus.Text = "Quit"
                End If
            Case POP3_RETR
                strBuffer = strBuffer & strData
                If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                    strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                    strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                    Set m_oMessage = New CMessage
                    m_colMessages.Add m_oMessage
                    
                    m_oMessage.CreateFromText strBuffer  '<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    
                    
                    'displays email to text box and writes to file
                    txtViewAll.Text = txtViewAll.Text & "*********************" & intCurrentMessage & "*******************" & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "From: " & m_oMessage.from & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "To: " & m_oMessage.MessageTo & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "CC: " & m_oMessage.cc & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "BCC: " & m_oMessage.BCC & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "MessageID: " & m_oMessage.MessageID & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "SendDate: " & m_oMessage.SendDate & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Sender: " & m_oMessage.Sender & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "ReturnPath: " & m_oMessage.ReturnPath & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Size: " & m_oMessage.Size & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Comments: " & m_oMessage.Comments & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Encrypted: " & m_oMessage.Encrypted & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "InReplyTo: " & m_oMessage.InReplyTo & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Received: " & m_oMessage.Received & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "References: " & m_oMessage.References & vbCrLf
                    txtViewAll.Text = txtViewAll.Text & "Subject: " & m_oMessage.subject & vbCrLf
                    
                    If gblnDumpToFile = True Then
                      Open App.Path & "\raw.txt" For Append As #10
                        Print #10, "************************************************************"
                        Print #10, " Email Number: " & CStr(intCurrentMessage)
                        Print #10, "From: " & m_oMessage.from
                        Print #10, "To: " & m_oMessage.MessageTo
                        Print #10, "CC: " & m_oMessage.cc
                        Print #10, "BCC: " & m_oMessage.BCC
                        Print #10, "MessageID: " & m_oMessage.MessageID
                        Print #10, "SendDate: " & m_oMessage.SendDate
                        Print #10, "Sender: " & m_oMessage.Sender
                        Print #10, "ReturnPath: " & m_oMessage.ReturnPath
                        Print #10, "Size: " & m_oMessage.Size
                        Print #10, "Comments: " & m_oMessage.Comments
                        Print #10, "Encrypted: " & m_oMessage.Encrypted
                        Print #10, "InReplyTo: " & m_oMessage.InReplyTo
                        Print #10, "Received: " & m_oMessage.Received
                        Print #10, "References: " & m_oMessage.References
                        Print #10, "Subject: " & m_oMessage.subject
                        Print #10, "MessageBody: " & m_oMessage.messagebody
                        Print #10, "************************************************************"
                      Close #10
                    End If
                    
                    '*****************************************************************************************************
                    ' F I L T E R I N G   F O R    S P A M   B E G I N S     H E R E
                    ' Definitions in GLOBAL.BAS file.
                    '*****************************************************************************************************
                    
                    'specific filtering
                    'Subject:  field
                    m_oMessage.Score = 0
                    m_oMessage.Score = m_oMessage.Score + LookForEmptySubject(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForSubjectAddOn(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForSubjectXAscii(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForSubjectNumbers(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForHealth(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForFinance(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForPorn(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForMisc(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForHWSW(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForAttract(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForDegree(m_oMessage.subject)
                    m_oMessage.Score = m_oMessage.Score + LookForHoliday(m_oMessage.subject)
                    
                    'Received:  field
                    m_oMessage.Score = m_oMessage.Score + LookForReceivedUnknown(m_oMessage.Received)
                    
                    'TO: field
                    m_oMessage.Score = m_oMessage.Score + LookForDomainCount(m_oMessage.MessageTo)
                    m_oMessage.Score = m_oMessage.Score + LookForMissingDomain(m_oMessage.MessageTo)
                    m_oMessage.Score = m_oMessage.Score + LookForMissingUserName(m_oMessage.MessageTo)
                    
                    'CC: field
                    m_oMessage.Score = m_oMessage.Score + LookForDomainCount(m_oMessage.cc)
                    
                    'FROM: field
                    m_oMessage.Score = m_oMessage.Score + LookForMissingFromAddress(m_oMessage.from)
                    m_oMessage.Score = m_oMessage.Score + LookForCountryCode(m_oMessage.from)
                    
                    'SendDate: field
                    m_oMessage.Score = m_oMessage.Score + LookForMissingDate(m_oMessage.SendDate)
                    
                    'MessageBody: field
                    m_oMessage.Score = m_oMessage.Score + LookForBodyText(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForHealth(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForFinance(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForPorn(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForMisc(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForHWSW(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForAttract(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForDegree(m_oMessage.messagebody)
                    m_oMessage.Score = m_oMessage.Score + LookForHoliday(m_oMessage.messagebody)
                    
                    'update pop status
                    txtStatus.Text = "Retrieving " & CStr(intCurrentMessage + 1)
                    
                    'calculate score
                    strMsg = "[Score: " & m_oMessage.Score & "]"
                    If m_oMessage.Score > 0 Then
                    txtSpamTotal.Text = CInt(txtSpamTotal.Text) + 1
                    If Val(txtTotal2.Text) > 0 Then
                      lblPercent.Caption = CStr(CInt(CSng(txtSpamTotal.Text) / CSng(txtTotal2.Text) * 100)) & " %"
                    End If
                    End If
                                        
                    'display subject
                    lstSubject.AddItem strMsg & vbTab & CStr(intCurrentMessage) & ":" & vbTab & "(" & CStr(bytesTotal) & ")" & vbTab & m_oMessage.subject
                        
                    'add email content to em( ) array
                    ReDim Preserve em(intCurrentMessage)
                    em(intCurrentMessage).subject = m_oMessage.subject
                    em(intCurrentMessage).from = m_oMessage.from
                    em(intCurrentMessage).messagebody = m_oMessage.messagebody
                    em(intCurrentMessage).cc = m_oMessage.cc
                    em(intCurrentMessage).delete_code = intAccCode
                    em(intCurrentMessage).bytes_total = bytesTotal
                   
                    'cleanup
                    Set m_oMessage = Nothing
                    strBuffer = ""
                    
                    If intCurrentMessage = intMessages Then
                       gblnDownloadComplete = True
                       txtStatus.Text = "Standby"
                       mnuConnectProcess.Enabled = True
                       mnuPopupDelete.Enabled = True
                       mnuConfig.Enabled = True
                       mnuPopupView.Enabled = True
                    Else
                        intCurrentMessage = intCurrentMessage + 1
                        txtCurrent.Text = intCurrentMessage
                        
                        'Change current state of session
                        m_State = POP3_RETR
                        Winsock1.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                    End If
                End If
            Case POP3_DELE
               Winsock1.SendData "DELE " & CStr(mintDelete) & vbCrLf
               m_Status = POP3_QUIT
            Case POP3_QUIT
                Winsock1.SendData "QUIT" & vbCrLf
                Winsock1.Close
                WinsockStatus False
                txtStatus.Text = "Quit"
        End Select
        mblnSending = True
    Else
            Winsock1.Close
            WinsockStatus False
            'MsgBox "POP3 Error: " & strData,  vbExclamation , "POP3 Error"
    End If
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: #" & number '& vbCrLf & Description
End Sub

'used for visual indication only
Private Sub WinsockStatus(state As Boolean)
  If state = True Then  'winsock is open
    shpOpen.BackColor = RGB(0, 255, 0)
    shpClose.BackColor = RGB(0, 100, 0)
  Else                       'winsock is closed
    shpClose.BackColor = RGB(0, 255, 0)
    shpOpen.BackColor = RGB(0, 100, 0)
  End If
End Sub
