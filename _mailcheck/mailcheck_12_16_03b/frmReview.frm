VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "CC:"
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   9255
      Begin VB.TextBox txtCC 
         Height          =   285
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   9075
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Update Filter"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8340
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Message Body:"
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   9255
      Begin VB.TextBox txtMessageBody 
         Height          =   3495
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From:"
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   9255
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   9135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject:"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   9135
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
Dim mstrKeep As String 'stores select string to add to filters

Private Sub cmdClose_Click()
  Unload frmAddWord
  Unload Me
End Sub

Private Sub cmdExtract_Click()
  If Len(mstrKeep) < 1 Then Exit Sub
  gstrString = mstrKeep
  frmAddWord.Show
End Sub


Private Sub Form_Load()
  If gintEmailToReview < 1 Then Exit Sub
  If gintEmailToReview > gintTotalEmails Then Exit Sub
  gstrString = ""
  txtSubject.Text = em(gintEmailToReview).subject
  txtCC.Text = em(gintEmailToReview).cc
  txtFrom.Text = em(gintEmailToReview).from
  txtMessageBody.Text = em(gintEmailToReview).messagebody
  frmReview.Caption = "Email " & CStr(gintEmailToReview)
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
