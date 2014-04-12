VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9780
   ClientLeft      =   1815
   ClientTop       =   330
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   11070
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   6900
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   6540
      Width           =   8775
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   360
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6180
      Width           =   8715
   End
   Begin VB.Frame Frame7 
      Caption         =   "Winsock Status"
      Height          =   1455
      Left            =   9480
      TabIndex        =   18
      Top             =   6240
      Width           =   1335
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Closed"
         Height          =   255
         Left            =   540
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Open"
         Height          =   195
         Left            =   540
         TabIndex        =   19
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
   Begin VB.Frame Frame6 
      Caption         =   "Email to Delete:"
      Height          =   1275
      Left            =   120
      TabIndex        =   15
      Top             =   4500
      Width           =   4155
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtDeleteStatus 
         Height          =   315
         Left            =   900
         TabIndex        =   23
         Top             =   660
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Email"
         Height          =   315
         Left            =   900
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtDelete 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   8460
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Subject Lines"
      Height          =   3555
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   9495
      Begin VB.ListBox lstSubject 
         Height          =   3180
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   9300
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "&Check Mailbox"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10440
      Top             =   180
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
      TabIndex        =   8
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtCurrent 
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtTotal 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "of"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   300
         Width           =   255
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
Private mstrProgram As String 'name of program
Private mstrVersion As String 'version of program
Private mstrDate As String 'date of last program change
Private mintDelete As Integer 'email to delete

'**************************************************************************
'  C H E C K   M A I L
'  Initiates log on sequence with mail server
'**************************************************************************
Private Sub cmdCheckMail_Click()
    
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
    cmdStop.Enabled = True
    cmdCheckMail.Enabled = False
    lstSubject.Clear
    
    'sets initial program state, closes winsock in case it is open and sets up
    'winsock for POP3 protocol
    m_State = POP3_Connect
    Winsock1.Close
    Winsock1.LocalPort = 0
    Winsock1.Connect txtHost, 110

End Sub

Private Sub cmdDelete_Click()
  'Winsock1.Close
  'WinsockStatus False
 ' Winsock1.LocalPort = 0
 ' Winsock1.Connect txtHost, 110
 ' m_State = POP3_Connect
 ' cmdDelete.Enabled = False
  mintDelete = CInt(txtDelete.Text)
  Winsock1.SendData "DELE " & CStr(mintDelete) & vbCrLf
  'm_State = POP3_QUIT
End Sub

'exits program
Private Sub cmdExit_Click()
  Winsock1.Close
  Unload Me
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    cmdCheckMail.Enabled = True
    m_State = POP3_QUIT
    intCurrentMessage = 0
End Sub



Private Sub Command1_Click()
                Winsock1.SendData "QUIT" & vbCrLf
End Sub

'Load form
Private Sub Form_Load()
  
  'loads module variables
  mstrProgram = "Mail Checker"
  mstrVersion = "v0.1"
  mstrDate = "December 12, 2003"
  
  'loads controls on interface
  frmMain.Caption = mstrProgram & " " & mstrVersion & " - " & mstrDate
  WinsockStatus False
  
  'loads spam word array from file
  Dim x As Integer
  Open App.Path & "\spamwords.txt" For Input As #1
    Do
    x = x + 1
    ReDim Preserve word(x)
    Line Input #1, word(x)
    Loop Until EOF(1)
  Close #1
  
End Sub

Private Sub lstSubject_Click()
  Dim x, s As Long
  
 'determine list item selected
  For x = 0 To lstSubject.ListCount - 1
    If lstSubject.Selected(x) = True Then
      s = lstSubject.ListIndex + 1
      Exit For
    End If
  Next x
  
  'update text box
  txtDelete.Text = s
  cmdDelete.Enabled = True
  txtDeleteStatus.Text = ""
  txtDeleteStatus.Text = FilterSubject(CInt(s))
  Text1.Text = em(s).subject

End Sub

'analyses subject for filtering state
Private Function FilterSubject(intNum As Integer) As String
  Dim intCode As Integer
  Dim strS, strTemp As String
  Dim x, z As Integer
   
  'load original subject string
  intCode = 0
  em(intNum).delete_code = intCode
  strS = em(intNum).subject
   
  'remove all spaces
  For x = 1 To Len(strS)
    If Mid(strS, x, 1) <> " " Then strTemp = strTemp & Mid(strS, x, 1)
  Next x
  strS = strTemp
   
  'convert to lower case
  strS = LCase(strS)
  
  'replace
  '| or ! or 1 with i
  '0       with o
  '@      with a
  '$       with s
  strTemp = ""
  For x = 1 To Len(strS)
    If Mid(strS, x, 1) = "!" Or Mid(strS, x, 1) = "|" Or Mid(strS, x, 1) = "1" Then
      strTemp = strTemp & "i"
    ElseIf Mid(strS, x, 1) = "0" Then
      strTemp = strTemp & "o"
    ElseIf Mid(strS, x, 1) = "@" Then
      strTemp = strTemp & "a"
    ElseIf Mid(strS, x, 1) = "$" Then
      strTemp = strTemp & "s"
    Else
      strTemp = strTemp & Mid(strS, x, 1)
    End If
  Next x
  strS = strTemp
  
  'delete all remaining pronunciation  marks
  strTemp = ""
  
  For x = 1 To Len(strS)
     If Asc(Mid(strS, x, 1)) >= 48 And Asc(Mid(strS, x, 1)) <= 57 Then
       strTemp = strTemp & Mid(strS, x, 1)
     ElseIf Asc(Mid(strS, x, 1)) >= 97 And Asc(Mid(strS, x, 1)) <= 122 Then
       strTemp = strTemp & Mid(strS, x, 1)
      Else
      End If
  Next x
  strS = strTemp

  'determines code
  'empty subjects
  If Len(strS) < 1 Then
    intCode = intCode + 1
  End If
  
  'longer subjects
  If Len(strS) > 30 Then
    intCode = intCode + 2
  End If
  
  'search for spamwords in subject
  For z = 1 To UBound(word)
    If InStr(1, strS, word(z)) > 0 Then
      intCode = intCode + 10
      Text3.Text = word(z)
      Exit For
    End If
  Next z
   
   
  em(intNum).delete_code = intCode 'save filter code...0 is okay
  Text2.Text = strS
  FilterSubject = CStr(intCode)

End Function


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
    'Debug.Print strData

    
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
                    em(intCurrentMessage).subject = LCase(m_oMessage.subject)
                    Dim strSub As String
                    strSub = LCase(m_oMessage.subject)  'convert subject to lower case
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
        End Select
        mblnSending = True
    Else
            Winsock1.Close
            WinsockStatus False
            'MsgBox "POP3 Error: " & strData,  vbExclamation , "POP3 Error"
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: #" & Number '& vbCrLf & Description
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
