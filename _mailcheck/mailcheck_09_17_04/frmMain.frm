VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   1860
   ClientTop       =   1140
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   12030
   Begin VB.Frame Frame6 
      Caption         =   "User Domain:"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   60
      Width           =   2475
      Begin VB.TextBox txtDomain 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Text            =   "dycon.com"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Spam Count:"
      Height          =   615
      Left            =   9660
      TabIndex        =   17
      Top             =   1620
      Width           =   1995
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
         Left            =   1500
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
      Width           =   1995
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   1020
         Width           =   1815
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
         Height          =   2595
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   9255
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   10980
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Password:"
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   60
      Width           =   1635
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         ToolTipText     =   "Ex: ********"
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Name:"
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   60
      Width           =   1635
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Ex: johndoe"
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mail Server:"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   60
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
      Left            =   9660
      TabIndex        =   6
      Top             =   2280
      Width           =   1995
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
' Updated: September 2004 by Chuck B
' Incorporates spam filtering techniques that assign a weighted value
' to the filters.
'**************************************************************************
Option Explicit

'used for winsock to pop comms
Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum

'variable declarations
Private m_State               As POP3States
Private m_oMessage            As CMessage
Private mblnDownload          As Boolean
Private mblnSending           As Boolean
Private mintDelete            As Integer 'email number to delete
Private m_Status              As Integer
Private mblnDownloadComplete  As Boolean 'true download of emails are complete

'************************************************ chkViewList_Click
'enables/disables the display of subject headers
'added in case children walk up behind you and
'begin reading over your shoulder as they tend to
'do :-)
Private Sub chkViewList_Click()
  If chkViewList.Value = vbChecked Then
    frmMain.Height = 5160
    lstSubject.Visible = True
    fraSubject.Height = 735 + lstSubject.Height + 250
  Else
    frmMain.Height = 2750
    lstSubject.Visible = False
    fraSubject.Height = 735
  End If
End Sub

'*********************************************** Form_Load
Private Sub Form_Load()
  
  'loads global variables - Global.bas
  LoadGlobalVariables
    
  'loads controls on interface
  frmMain.Caption = gstrProgram & " " & gstrVersion & " - " & gstrDate
  WinsockStatus False
  fraSubject.Height = 735
  chkViewList_Click
End Sub

'*********************************************** lstSubject_MouseDown
Private Sub lstSubject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    If gblnDownloadComplete = True Then
      mnuPopupDelete.Enabled = True
    Else
      mnuPopupDelete.Enabled = False  'don't delete until download complete
    End If
    If lstSubject.ListCount > 0 Then PopupMenu Popup
  End If
End Sub

'*********************************************** mnuConfigDump_Click
'this options allows all emails to be dumped to a text file
Private Sub mnuConfigDump_Click()
  If mnuConfigDump.Checked = True Then
    mnuConfigDump.Checked = False
    gblnDumpToFile = False
  Else
    mnuConfigDump.Checked = True
    gblnDumpToFile = True
  End If
End Sub

'**************************************************************************
'  C H E C K   M A I L    M E N U
'  Initiates log on sequence with mail server
'**************************************************************************
Private Sub mnuConnectCheck_Click()
 
  'initializes variables and control properties
  gblnDownloadComplete = False
  mnuConnectProcess.Enabled = False
  mnuPopupDelete.Enabled = False
  mnuConfig.Enabled = False
  txtCurrent.Text = 0
  txtTotal.Text = 0
  gintTotalEmails = 0
  lstSubject.Clear
  ReDim em(0) 'this stores all relevent details pertaining to all emails
              'eventually I will replace this with an array of objects
              'CMessage
              
  'verify critical text fields are completed
  If Len(txtDomain.Text) < 1 Then
    MsgBox "Domain field is empty!"
    Exit Sub
  End If
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
  g_sUserDomain = txtDomain.Text
  g_sUserName = txtUserName.Text & "@" & g_sUserDomain
  lstSubject.Clear
  
  'sets initial program state, closes winsock in case it is open and sets up
  'winsock for POP3 protocol
  m_State = POP3_Connect
  ws.Close  'in case it has been left open
  ws.LocalPort = 0
  ws.Connect txtHost.Text, 110
End Sub

'********************************************** mnuConnectExit_Click
'close winsock connection
Private Sub mnuConnectExit_Click()
  ws.Close
  Unload frmReview
End Sub

'********************************************** mnuConnectProcess_Click
'After a brief few seconds, the lstSubject list should begin to fill up
'with email subject lines
Private Sub mnuConnectProcess_Click()
  Dim X As Integer
  For X = 1 To UBound(em)
    If em(X).Delete = True And em(X).Dead = False Then
      ws.SendData "DELE " & CStr(X) & vbCrLf
      lstSubject.List(X - 1) = "[ DELETED FROM MAIL SERVER ]"
      em(X).Dead = True
    End If
  Next X
  ws.SendData "QUIT" & vbCrLf
End Sub

'**************************************** mnuPopupDelete_Click
Private Sub mnuPopupDelete_Click()
  mintDelete = lstSubject.ListIndex + 1
'  If mintDelete > -1 Then ws.SendData "DELE " & CStr(mintDelete) & vbCrLf
  If em(mintDelete).Delete = True And em(mintDelete).Dead = False Then 'already marked for delete
    em(mintDelete).Delete = False
    lstSubject.List(mintDelete - 1) = CStr(mintDelete) & ":  (" & CStr(em(mintDelete).Bytes_total) & ")  " & CStr(em(mintDelete).Subject)
  Else                                               'NOT marked for delete
    em(mintDelete).Delete = True
    lstSubject.List(mintDelete - 1) = "[DELETE]: " & CStr(mintDelete) & ":  (" & CStr(em(mintDelete).Bytes_total) & ")  " & CStr(em(mintDelete).Subject)
  End If
End Sub

'**************************************** mnuPopupView_Click
' shows frmReview with selected email
Private Sub mnuPopupView_Click()
  gintEmailToReview = lstSubject.ListIndex + 1
  If gintEmailToReview > -1 Then frmReview.Show
End Sub

'**************************************************************************
'  D A T A   A R R I V A L
'  Manage all incoming data
'**************************************************************************
Private Sub ws_DataArrival(ByVal bytesTotal As Long)
   
  'variable declaration and initialization
  
  Dim strData                 As String
  Dim intCode                 As String 'indicates spam code
  Dim strMsg                  As String 'message to indicate SPAM or OK
  Dim intAccCode              As Integer 'accumulated codes for one email...several reasons why it is spam
  Static intMessages          As Integer 'the number of messages to be loaded
  Static intCurrentMessage    As Integer 'the counter of loaded messages
  Static strBuffer            As String  'the buffer of the loading message
  mblnSending = False
  
  'Save the received data into strData variable
  ws.GetData strData
  DoEvents
  If Left$(strData, 1) = "+" Or m_State = POP3_RETR Then
    Select Case m_State
    
      Case POP3_Connect  '************************* POP Connect
        intMessages = 0
        m_State = POP3_USER
        ws.SendData "USER " & txtUserName & vbCrLf
        txtStatus.Text = "User Login"
        intCurrentMessage = 0 'in case you are logging in again
        
      Case POP3_USER     '************************* POP User
        m_State = POP3_PASS
        ws.SendData "PASS " & txtPassword & vbCrLf
        txtStatus.Text = "Password"
        
      Case POP3_PASS      '************************* POP Pass
        m_State = POP3_STAT
        ws.SendData "STAT" & vbCrLf
        mblnDownload = True
        txtStatus.Text = "Stat"
        
      Case POP3_STAT      '************************* POP Stat
        WinsockStatus True
        intMessages = CInt(Mid$(strData, 5, InStr(5, strData, " ") - 5))
        txtTotal.Text = intMessages       ' total emails
        txtTotal2.Text = intMessages
        gintTotalEmails = intMessages
        If intMessages > 0 Then
          If mintDelete > 0 Then
            m_State = POP3_DELE
          Else
            m_State = POP3_RETR
            intCurrentMessage = intCurrentMessage + 1
            ws.SendData "RETR 1" & vbCrLf
            txtStatus.Text = "Retrieving " & CStr(intCurrentMessage + 1)
          End If
        Else
          m_State = POP3_QUIT
          ws.SendData "QUIT" & vbCrLf
          MsgBox "You have no mail.", vbInformation
          txtStatus.Text = "Quit"
        End If
        
      Case POP3_RETR      '************************* POP Retrieve
        strBuffer = strBuffer & strData
        If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
          strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
          strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
          Set m_oMessage = New CMessage
          Dim uScore As EMAIL_SCORE
          Dim nFile As Integer
          nFile = FreeFile
          
          'strBuffer is text stream of this email. It is parsed into various
          'm_oMessage propertiers by the method .CreateFromText strBuffer
          m_oMessage.CreateFromText strBuffer
             
          'print if selected
          If gblnDumpToFile = True Then
            Open App.Path & "\emaildump.txt" For Append As #nFile
              Print #nFile, "************************************************************"
              Print #nFile, "Email Number: " & CStr(intCurrentMessage)
              Print #nFile, "From: " & m_oMessage.From
              Print #nFile, "To: " & m_oMessage.MessageTo
              Print #nFile, "CC: " & m_oMessage.CC
              Print #nFile, "BCC: " & m_oMessage.BCC
              Print #nFile, "MessageID: " & m_oMessage.MessageID
              Print #nFile, "SendDate: " & m_oMessage.SendDate
              Print #nFile, "Sender: " & m_oMessage.Sender
              Print #nFile, "ReturnPath: " & m_oMessage.ReturnPath
              Print #nFile, "Size: " & m_oMessage.Size + Len(m_oMessage.MessageBody)
              Print #nFile, "Comments: " & m_oMessage.Comments
              Print #nFile, "Encrypted: " & m_oMessage.Encrypted
              Print #nFile, "InReplyTo: " & m_oMessage.InReplyTo
              Print #nFile, "Received: " & m_oMessage.Received
              Print #nFile, "References: " & m_oMessage.References
              Print #nFile, "Subject: " & m_oMessage.Subject
              Print #nFile, "MessageBody: " & m_oMessage.MessageBody
              Print #nFile, "************************************************************"
            Close #nFile
          End If
            
          '*****************************************************************************************************
          ' F I L T E R I N G   F O R    S P A M   B E G I N S     H E R E
          ' Definitions in GLOBAL.BAS file.
          ' Filter returns a score and the scores are ultimately totaled.
          ' Bad things increase the score, good things decrease the score.
          '*****************************************************************************************************
          
          'specific filtering
          'Subject:  field
          m_oMessage.Score = 0
          uScore.EmptySubject = LookForEmptySubject(m_oMessage.Subject)
          uScore.SubjectAddon = LookForSubjectAddOn(m_oMessage.Subject)
          uScore.SubjectXAscii = LookForSubjectXAscii(m_oMessage.Subject)
          uScore.SubjectNumbers = LookForSubjectNumbers(m_oMessage.Subject)
          uScore.Health = LookForHealth(m_oMessage.Subject)
          uScore.Finance = LookForFinance(m_oMessage.Subject)
          uScore.Porn = LookForPorn(m_oMessage.Subject)
          uScore.Misc = LookForMisc(m_oMessage.Subject)
          uScore.HWSW = m_oMessage.Score + LookForHWSW(m_oMessage.Subject)
          uScore.Attract = LookForAttract(m_oMessage.Subject)
          uScore.Degree = LookForDegree(m_oMessage.Subject)
          uScore.Holiday = LookForHoliday(m_oMessage.Subject)
          uScore.SubjectConsecCount = LookForSubjectConsecCon(m_oMessage.Subject)
          uScore.Friend = LookForFriendText(m_oMessage.Subject)
          
          'subtotal 1
          m_oMessage.Score = uScore.EmptySubject + uScore.SubjectAddon + uScore.SubjectXAscii
          m_oMessage.Score = m_oMessage.Score + uScore.SubjectNumbers + uScore.SubjectConsecCount
                              
          'Received:  field
          uScore.ReceivedUnknown = LookForReceivedUnknown(m_oMessage.Received)
                              
          'TO: field
          uScore.DomainCount = LookForDomainCount(m_oMessage.MessageTo)
          uScore.ToMissingDomain = LookForMissingDomain(m_oMessage.MessageTo)
          uScore.ToMissingUserName = LookForMissingUserName(m_oMessage.MessageTo)
          uScore.Friend = uScore.Friend + LookForFriendText(m_oMessage.MessageTo)
          
          'CC: field
          uScore.DomainCount = uScore.DomainCount + LookForDomainCount(m_oMessage.CC)
          uScore.Friend = uScore.Friend + LookForFriendText(m_oMessage.CC)
          
          'FROM: field
          uScore.FromMissing = LookForMissingFromAddress(m_oMessage.From)
          uScore.CountryCode = LookForCountryCode(m_oMessage.From)
          uScore.Friend = uScore.Friend + LookForFriendText(m_oMessage.From)
          
          'SendDate: field
          uScore.DateMissing = LookForMissingDate(m_oMessage.SendDate)
          
          'subtotal 2
          m_oMessage.Score = m_oMessage.Score + uScore.ReceivedUnknown + uScore.DomainCount + uScore.ToMissingDomain
          m_oMessage.Score = m_oMessage.Score + uScore.ToMissingUserName + uScore.FromMissing + uScore.CountryCode
          m_oMessage.Score = m_oMessage.Score + uScore.DateMissing
          
          'MessageBody: field
          uScore.BodyText = LookForBodyText(m_oMessage.MessageBody)
          uScore.Health = uScore.Health + LookForHealth(m_oMessage.MessageBody)
          uScore.Finance = uScore.Finance + LookForFinance(m_oMessage.MessageBody)
          uScore.Porn = uScore.Porn + LookForPorn(m_oMessage.MessageBody)
          uScore.Misc = uScore.Misc + LookForMisc(m_oMessage.MessageBody)
          uScore.HWSW = uScore.HWSW + LookForHWSW(m_oMessage.MessageBody)
          uScore.Attract = uScore.Attract + LookForAttract(m_oMessage.MessageBody)
          uScore.Degree = uScore.Degree + LookForDegree(m_oMessage.MessageBody)
          uScore.Holiday = uScore.Holiday + LookForHoliday(m_oMessage.MessageBody)
          uScore.Friend = uScore.Friend + LookForFriendText(m_oMessage.MessageBody)
          
          'subtotal 3
          Dim nScore As Integer
          nScore = 0
          nScore = nScore + uScore.BodyText + uScore.Health + uScore.Finance
          nScore = nScore + uScore.Porn + uScore.Misc + uScore.HWSW
          nScore = nScore + uScore.Attract + uScore.Degree + uScore.Holiday + uScore.Friend
          
          '****************************************** Extra filtering
          'just in case this email scored a good value then check again using more specialized
          'techniques
          'technique #1 - Look at messagebody and only keep numbers and letters
          '               Then check against all word list arrays
          If m_oMessage.Score + nScore < g_uScore.SpamMinimum Then
            Dim sTemp As String
            sTemp = RemoveNonAlphaNumeric(m_oMessage.MessageBody)
            
            uScore.BodyText = LookForBodyText(sTemp)
            uScore.Health = uScore.Health + LookForHealth(sTemp)
            uScore.Finance = uScore.Finance + LookForFinance(sTemp)
            uScore.Porn = uScore.Porn + LookForPorn(sTemp)
            uScore.Misc = uScore.Misc + LookForMisc(sTemp)
            uScore.HWSW = uScore.HWSW + LookForHWSW(sTemp)
            uScore.Attract = uScore.Attract + LookForAttract(sTemp)
            uScore.Degree = uScore.Degree + LookForDegree(sTemp)
            uScore.Holiday = uScore.Holiday + LookForHoliday(sTemp)
            uScore.Friend = uScore.Friend + LookForFriendText(sTemp)
            
            'nScore = 0
            'nScore = nScore + uScore.BodyText + uScore.Health + uScore.Finance
            'nScore = nScore + uScore.Porn + uScore.Misc + uScore.HWSW
            'nScore = nScore + uScore.Attract + uScore.Degree + uScore.Holiday + uScore.Friend
            'If nScore > g_uScore.SpamMinimum Then
            '  MsgBox "Got it!  " & nScore & " points found!"
            'End If
          End If
                    
          'TOtoal score
          m_oMessage.Score = m_oMessage.Score + uScore.BodyText + uScore.Health + uScore.Finance
          m_oMessage.Score = m_oMessage.Score + uScore.Porn + uScore.Misc + uScore.HWSW
          m_oMessage.Score = m_oMessage.Score + uScore.Attract + uScore.Degree + uScore.Holiday + uScore.Friend
          
          'update pop status
          txtStatus.Text = "Retrieving " & CStr(intCurrentMessage + 1)
          
          'calculate score and display overall
          ReDim Preserve em(intCurrentMessage)
          If m_oMessage.Score > uScore.SpamMinimum Then
              em(intCurrentMessage).Delete = True
              strMsg = "[DELETE]: [Score: " & m_oMessage.Score & "]" & vbTab & CStr(intCurrentMessage) & ":" & vbTab & "(" & CStr(m_oMessage.Size + Len(m_oMessage.MessageBody)) & ")" & vbTab & m_oMessage.Subject
          Else
              em(intCurrentMessage).Delete = False
              strMsg = "[Score: " & m_oMessage.Score & "]" & vbTab & CStr(intCurrentMessage) & ":" & vbTab & "(" & CStr(m_oMessage.Size + Len(m_oMessage.MessageBody)) & ")" & vbTab & m_oMessage.Subject
          End If
          
          If m_oMessage.Score > g_uScore.SpamMinimum Then
            txtSpamTotal.Text = CInt(txtSpamTotal.Text) + 1
            If Val(txtTotal2.Text) > 0 Then
              lblPercent.Caption = CStr(CInt(CSng(txtSpamTotal.Text) / CSng(txtTotal2.Text) * 100)) & " %"
            End If
          End If
                                
          'display subject
          lstSubject.AddItem strMsg '& vbTab & CStr(intCurrentMessage) & ":" & vbTab & "(" & CStr(m_oMessage.Size + Len(m_oMessage.MessageBody)) & ")" & vbTab & m_oMessage.Subject
              
          'add email content to em( ) array
          'this is used after email have been
          'downloaded
  
          em(intCurrentMessage).From = m_oMessage.From
          em(intCurrentMessage).MessageTo = m_oMessage.MessageTo
          em(intCurrentMessage).CC = m_oMessage.CC
          em(intCurrentMessage).BCC = m_oMessage.BCC
          em(intCurrentMessage).MessageID = m_oMessage.MessageID
          em(intCurrentMessage).SendDate = m_oMessage.SendDate
          em(intCurrentMessage).Sender = m_oMessage.Sender
          em(intCurrentMessage).ReturnPath = m_oMessage.ReturnPath
          em(intCurrentMessage).Size = m_oMessage.Size
          em(intCurrentMessage).Comments = m_oMessage.Comments
          em(intCurrentMessage).Encrypted = m_oMessage.Encrypted
          em(intCurrentMessage).InReplyTo = m_oMessage.InReplyTo
          em(intCurrentMessage).Received = m_oMessage.Received
          em(intCurrentMessage).References = m_oMessage.References
          em(intCurrentMessage).Subject = m_oMessage.Subject
          em(intCurrentMessage).MessageBody = m_oMessage.MessageBody
          em(intCurrentMessage).delete_code = intAccCode
          em(intCurrentMessage).Bytes_total = m_oMessage.Size + Len(m_oMessage.MessageBody)
          em(intCurrentMessage).Score = m_oMessage.Score

          'cleanup
          Set m_oMessage = Nothing
          strBuffer = ""
          
          'evaluate status of download
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
            ws.SendData "RETR " & _
            CStr(intCurrentMessage) & vbCrLf
          End If
        End If
        
      Case POP3_DELE      '************************* POP Delete
        ws.SendData "DELE " & CStr(mintDelete) & vbCrLf
        m_Status = POP3_QUIT
        
      Case POP3_QUIT      '************************* POP Quite
        ws.SendData "QUIT" & vbCrLf
        ws.Close
        WinsockStatus False
        txtStatus.Text = "Quit"
    End Select
    mblnSending = True
  Else
    ws.Close
    WinsockStatus False
  End If
End Sub

'******************************************** ws_Error
Private Sub ws_Error(ByVal number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: #" & number '& vbCrLf & Description
End Sub

'******************************************** WinsockStatus
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
