VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Filter Results:"
      Height          =   5835
      Left            =   7140
      TabIndex        =   10
      Top             =   60
      Width           =   3495
      Begin VB.TextBox txtWhy 
         Height          =   4995
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   3315
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
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CC:"
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   7035
      Begin VB.TextBox txtCC 
         Height          =   285
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Update Filter"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8340
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Message Body:"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   7035
      Begin VB.TextBox txtMessageBody 
         Height          =   3495
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From:"
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   7035
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject:"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7035
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
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
  Dim intCode As Integer
  
  'user has updated the filter list
  If gblnUpdate = True Then
    intCode = FilterSubject(em(gintEmailToReview).subject)
    frmMain.lstSubject.List(gintEmailToReview - 1) = "[ SPAM ]: " & CStr(gintEmailToReview) & ":  (" & CStr(em(gintEmailToReview).bytes_total) & ")  " & CStr(em(gintEmailToReview).subject)
    em(gintEmailToReview).delete_code = 4 'phrase or word added to list
  End If
  Unload frmAddWord
  Unload Me
End Sub

Private Sub cmdExtract_Click()
  If Len(mstrKeep) < 1 Then Exit Sub
  gstrString = mstrKeep
  frmAddWord.Show
End Sub


Private Sub Form_Load()
  On Error GoTo myerror
  
  If gintEmailToReview < 1 Then Exit Sub
  If gintEmailToReview > gintTotalEmails Then Exit Sub
  gstrString = ""
  If Len(em(gintEmailToReview).subject) > 32000 Then
   txtSubject.Text = Left(em(gintEmailToReview).subject, 32)
  Else
    txtSubject.Text = em(gintEmailToReview).subject
  End If
  txtCC.Text = em(gintEmailToReview).cc
  txtFrom.Text = em(gintEmailToReview).from
  
  '************************ generates Out Of Memory errors sometimes
  txtMessageBody.Text = em(gintEmailToReview).messagebody
  frmReview.Caption = "Email " & CStr(gintEmailToReview)
  gblnUpdate = False
  
  'updates txtWhy textbox with message about why the spam is spam
  UpdateStatus
  Exit Sub
  
myerror:
  If Err.number = 7 Then
    MsgBox Err.number & "  " & Err.Description & " Size: " & Len(em(gintEmailToReview).messagebody)
  End If
  Resume Next
End Sub

Public Sub UpdateStatus()
 'updates txtWhy textbox with message about why the spam is spam
  If em(gintEmailToReview).delete_code > 0 Then
    lblStatus.Caption = "SPAM!"
    lblStatus.ForeColor = vbRed
    txtWhy.Visible = True
    txtWhy.Text = ""

    If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_SHORT) Then txtWhy.Text = txtWhy.Text & "Subject is too short." & vbCrLf
    If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_LONG) Then txtWhy.Text = txtWhy.Text & "Subject is too long." & vbCrLf
    If (em(gintEmailToReview).delete_code And FILTER_SUB_BAD_WORDS) Then
      txtWhy.Text = txtWhy.Text & "Subject has Spam phrases." & vbCrLf
      txtWhy.Text = txtWhy.Text & "*******************************" & vbCrLf & em(gintEmailToReview).sub_word & vbCrLf
    End If
    If (em(gintEmailToReview).delete_code And FILTER_SUB_TOO_MANY_CONSONANTS) Then txtWhy.Text = txtWhy.Text & "Subject has too many consonants." & vbCrLf
    If (em(gintEmailToReview).delete_code And FILTER_MSG_BAD_WORDS) Then
      txtWhy.Text = txtWhy.Text & "Message has Spam phrases." & vbCrLf
      txtWhy.Text = txtWhy.Text & "*********************************  " & vbCrLf & em(gintEmailToReview).msg_word & vbCrLf
    End If
    If (em(gintEmailToReview).delete_code And FILTER_MSG_TOO_MANY_CONSONANTS) Then txtWhy.Text = txtWhy.Text & "Message has too many consonants." & vbCrLf
  Else
    lblStatus.Caption = "OK!"
    lblStatus.ForeColor = vbBlack
    txtWhy.Visible = False
  End If
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
