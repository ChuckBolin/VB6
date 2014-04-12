VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2115
   ClientLeft      =   1815
   ClientTop       =   615
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   11295
   Begin VB.Frame Frame7 
      Caption         =   "Winsock Status"
      Height          =   1455
      Left            =   9780
      TabIndex        =   13
      Top             =   120
      Width           =   1395
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Closed"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Open"
         Height          =   195
         Left            =   540
         TabIndex        =   14
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
      Height          =   3555
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   9495
      Begin VB.CheckBox chkViewList 
         Caption         =   "View Subjects"
         Height          =   255
         Left            =   7800
         TabIndex        =   17
         Top             =   180
         Width           =   1455
      End
      Begin VB.ListBox lstSubject 
         Height          =   2985
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   9255
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
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
   Begin VB.Frame Frame4 
      Caption         =   "Total Emails:"
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtCurrent 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtTotal 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "of"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnectCheck 
         Caption         =   "&Check Mailbox"
      End
      Begin VB.Menu mnuConnectProcess 
         Caption         =   "&Process Email"
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
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupView 
         Caption         =   "&View Email"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "&Delete Email"
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


Private Sub cmdDelete_Click()
  'Winsock1.Close
  'WinsockStatus False
 ' Winsock1.LocalPort = 0
 ' Winsock1.Connect txtHost, 110
 ' m_State = POP3_Connect
 ' cmdDelete.Enabled = False
  'mintDelete = CInt(txtDelete.Text)
  'Winsock1.SendData "DELE " & CStr(mintDelete) & vbCrLf
  'm_State = POP3_QUIT
End Sub

Private Sub cmdCheckMail_Click()

End Sub

Private Sub chkViewList_Click()
  If chkViewList.Value = vbChecked Then
    frmMain.Height = 5160
    lstSubject.Visible = True
    fraSubject.Height = 500 + lstSubject.Height + 50
    cmdExit.Top = 4000
  Else
    frmMain.Height = 2750
    lstSubject.Visible = False
    fraSubject.Height = 500
    cmdExit.Top = 1680
  End If
End Sub

'exits program
Private Sub cmdExit_Click()
  Winsock1.Close
  Unload Me
End Sub


Private Sub cmdStop_Click()
   ' cmdStop.Enabled = False
   ' cmdCheckMail.Enabled = True
    'm_State = POP3_QUIT
   ' intCurrentMessage = 0
End Sub



Private Sub Command1_Click()
          '      Winsock1.SendData "QUIT" & vbCrLf
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
  fraSubject.Height = 500
  
  'loads spam word array from file
  Open App.Path & "\spamwords.txt" For Append As #1
    If LOF(1) < 1 Then
      Close 1
      Exit Sub
    End If
  Close
  Dim X As Integer
  Open App.Path & "\spamwords.txt" For Input As #1
    Do
    X = X + 1
    ReDim Preserve word(X)
    Line Input #1, word(X)
    Loop Until EOF(1)
  Close #1
  
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
     If lstSubject.ListCount > 0 Then PopupMenu Popup
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
  Winsock1.SendData "QUIT" & vbCrLf
End Sub

Private Sub mnuPopupDelete_Click()
  mintDelete = CInt(txtDelete.Text)
  Winsock1.SendData "DELE " & CStr(mintDelete) & vbCrLf
End Sub

Private Sub mnuPopupView_Click()
  gintEmailToReview = lstSubject.ListIndex - 1
  If gintEmailToReview > -1 Then frmReview.Show
End Sub

'**************************************************************************
'  D A T A   A R R I V A L
'  Manage all incoming data
'**************************************************************************
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   
    mblnSending = False
    Dim strData As String
    
    Static intMessages          As Integer 'the number of messages to be loaded
    Static intCurrentMessage    As Integer 'the counter of loaded messages
    Static strBuffer            As String  'the buffer of the loading message
    
    'Save the received data into strData variable
    Winsock1.GetData strData
    
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
                gintTotalEmails = intMessages
                If intMessages > 0 Then
                  If mintDelete > 0 Then
                    m_State = POP3_DELE
                  Else
                    m_State = POP3_RETR
                    intCurrentMessage = intCurrentMessage + 1
                    Winsock1.SendData "RETR 1" & vbCrLf
                    txtStatus.Text = "Retrieve"
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
                    m_oMessage.CreateFromText strBuffer  '<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    'Open App.Path & "\subject.txt" For Append As #1
                    '  Print #1, CStr(intCurrentMessage) & ": " & m_oMessage.Subject
                    'Close #1
                    'Open App.Path & "\from.txt" For Append As #1
                    '  Print #1, CStr(intCurrentMessage) & ": " & m_oMessage.From
                    'Close #1
                    lstSubject.AddItem CStr(intCurrentMessage) & ":  (" & CStr(bytesTotal) & ")  " & m_oMessage.subject
                    ReDim Preserve em(intCurrentMessage)
                    em(intCurrentMessage).subject = m_oMessage.subject
                    em(intCurrentMessage).from = m_oMessage.from
                    em(intCurrentMessage).messagebody = m_oMessage.messagebody
                    em(intCurrentMessage).cc = m_oMessage.cc
                    
                   Set m_oMessage = Nothing
                    strBuffer = ""
                    If intCurrentMessage = intMessages Then
                        'm_State = POP3_QUIT
                       ' Winsock1.SendData "QUIT" & vbCrLf
                    Else
                        intCurrentMessage = intCurrentMessage + 1
                        txtCurrent.Text = intCurrentMessage
                        
                        'Change current state of session
                        m_State = POP3_RETR
                        Winsock1.SendData "RETR " & _
                        CStr(intCurrentMessage) & vbCrLf
                        'Debug.Print "RETR " & intCurrentMessage
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

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
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
