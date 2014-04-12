VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   1860
   ClientTop       =   855
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   10680
   Begin VB.Frame Frame6 
      Caption         =   "To:"
      Height          =   555
      Left            =   60
      TabIndex        =   15
      Top             =   1440
      Width           =   5355
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   180
         Width           =   5175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Filter Results:"
      Height          =   3495
      Left            =   5460
      TabIndex        =   10
      Top             =   60
      Width           =   5175
      Begin VB.TextBox txtWhy 
         Height          =   2895
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   540
         Width           =   4995
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CC:"
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   5355
      Begin VB.TextBox txtCC 
         Height          =   285
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Update Filter"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   8520
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Message Body:"
      Height          =   4755
      Left            =   60
      TabIndex        =   4
      Top             =   3660
      Width           =   10575
      Begin VB.TextBox txtMessageBody 
         Height          =   4455
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   10395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From:"
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   2040
      Width           =   5355
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject:"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5355
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   900
      TabIndex        =   14
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Score:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2820
      Width           =   615
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
' f r m R E V I E W  - December 2003
' Allows user to view complete email
'**************************************************************************
Option Explicit
Private mstrKeep As String 'stores select string to add to filters
Public gblnUpdate As Boolean 'true if filter updated

Private Sub cmdClose_Click()

  Unload frmAddWord
  Unload Me
End Sub

Private Sub cmdExtract_Click()
  'If Len(mstrKeep) < 1 Then Exit Sub
  'gstrString = mstrKeep
  'frmAddWord.Show
End Sub


Private Sub Form_Load()
  On Error GoTo myerror
  
  If gintEmailToReview < 1 Then Exit Sub
  If gintEmailToReview > gintTotalEmails Then Exit Sub
  gstrString = ""
  
  'shows subject
  If Len(em(gintEmailToReview).Subject) > 40 Then
   txtSubject.Text = Left(em(gintEmailToReview).Subject, 40)
  Else
    txtSubject.Text = em(gintEmailToReview).Subject
  End If
  
  txtCC.Text = em(gintEmailToReview).CC
  txtFrom.Text = em(gintEmailToReview).From
  txtTo.Text = em(gintEmailToReview).MessageTo
  
  'shows message body
  If Len(em(gintEmailToReview).MessageBody) > 32000 Then
    txtMessageBody.Text = Left(em(gintEmailToReview).MessageBody, 32000)
  Else
    txtMessageBody.Text = em(gintEmailToReview).MessageBody
  End If
  
  'shows email number and byte size
  frmReview.Caption = "Email " & CStr(gintEmailToReview) & "  Size: " & em(gintEmailToReview).Bytes_total & " bytes"
  
  gblnUpdate = False
  
  'miscellaneous info
  txtWhy.Text = txtWhy.Text & "ReturnPath: " & em(gintEmailToReview).ReturnPath & vbCrLf
  txtWhy.Text = txtWhy.Text & "BCC: " & em(gintEmailToReview).BCC & vbCrLf
  txtWhy.Text = txtWhy.Text & "MessageID: " & em(gintEmailToReview).MessageID & vbCrLf
  txtWhy.Text = txtWhy.Text & "SendDate: " & em(gintEmailToReview).SendDate & vbCrLf
  txtWhy.Text = txtWhy.Text & "Sender: " & em(gintEmailToReview).Sender & vbCrLf
  txtWhy.Text = txtWhy.Text & "Size: " & em(gintEmailToReview).Size & vbCrLf
  txtWhy.Text = txtWhy.Text & "Comments: " & em(gintEmailToReview).Comments & vbCrLf
  txtWhy.Text = txtWhy.Text & "Encrypted: " & em(gintEmailToReview).Encrypted & vbCrLf
  txtWhy.Text = txtWhy.Text & "InReplyTo: " & em(gintEmailToReview).InReplyTo & vbCrLf
  txtWhy.Text = txtWhy.Text & "Received: " & em(gintEmailToReview).Received & vbCrLf
  txtWhy.Text = txtWhy.Text & "References: " & em(gintEmailToReview).References & vbCrLf
  
  'displays score
  If em(gintEmailToReview).Score < g_uScore.SpamMinimum Then
    lblScore.ForeColor = vbGreen
  Else
    lblScore.ForeColor = vbRed
  End If
  lblScore.Caption = em(gintEmailToReview).Score
  
  Exit Sub
  
myerror:
  If Err.number = 7 Then
    MsgBox Err.number & "  " & Err.Description & " Size: " & Len(em(gintEmailToReview).MessageBody)
  End If
  Resume Next
End Sub

Public Sub UpdateStatus()
 'updates txtWhy textbox with message about why the spam is spam
  'If em(gintEmailToReview).delete_code > 0 Then
    'lblStatus.Caption = "SPAM!"
    'lblStatus.ForeColor = vbRed
    'txtWhy.Visible = True
    'txtWhy.Text = ""

    'If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_SHORT) Then txtWhy.Text = txtWhy.Text & "Subject is too short." & vbCrLf
    'If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_LONG) Then txtWhy.Text = txtWhy.Text & "Subject is too long." & vbCrLf
    'If (em(gintEmailToReview).delete_code And FILTER_SUB_BAD_WORDS) Then
    '  txtWhy.Text = txtWhy.Text & "Subject has Spam phrases." & vbCrLf
    '  txtWhy.Text = txtWhy.Text & "*******************************" & vbCrLf & em(gintEmailToReview).sub_word & vbCrLf
    'End If
    'If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_MANY_CONSONANTS) Then txtWhy.Text = txtWhy.Text & "Subject has too many consonants." & vbCrLf
    'If (em(gintEmailToReview).delete_code And FILTER_MSG_BAD_WORDS) Then
    '  txtWhy.Text = txtWhy.Text & "Message has Spam phrases." & vbCrLf
    '  txtWhy.Text = txtWhy.Text & "*********************************  " & vbCrLf & em(gintEmailToReview).msg_word & vbCrLf
    'End If
    'If (em(gintEmailToReview).delete_code And FILTER_MSG_TOO_MANY_CONSONANTS) Then txtWhy.Text = txtWhy.Text & "Message has too many consonants." & vbCrLf
  'Else
    'l 'blStatus.Caption = "OK!"
    'lblStatus.ForeColor = vbBlack
    'txtWhy.Visible = False
  'End If
End Sub

Private Sub txtCC_LostFocus()
  If txtCC.SelLength > 0 Then
    mstrKeep = txtFrom.SelText
  End If
End Sub

Private Sub txtFrom_LostFocus()
  If txtFrom.SelLength > 0 Then
    mstrKeep = txtFrom.SelText
  End If
End Sub

Private Sub txtMessageBody_LostFocus()
  If txtMessageBody.SelLength > 0 Then
    mstrKeep = txtMessageBody.SelText
  End If
End Sub

Private Sub txtSubject_LostFocus()
  If txtSubject.SelLength > 0 Then
    mstrKeep = txtSubject.SelText
  End If
End Sub
