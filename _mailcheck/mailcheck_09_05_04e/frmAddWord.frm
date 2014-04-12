VERSION 5.00
Begin VB.Form frmAddWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Word or Phrase to Filter List"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddSubject 
      Caption         =   "Add to Filter"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modified Word or Phrase to Add to List"
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   7515
      Begin VB.TextBox txtAdd 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   7395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Word or Phrase:"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7515
      Begin VB.TextBox txtPhrase 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   7395
      End
   End
End
Attribute VB_Name = "frmAddWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddSubject_Click()
  
  'adds word to spamwords file
  Dim X As Integer
  Open App.Path & "\spamwords.txt" For Append As #1
  ReDim Preserve word(UBound(word) + 1)
  Print #1, txtAdd.Text
  word(UBound(word)) = txtAdd.Text
  Close #1
  frmReview.gblnUpdate = True
 ' em(gintEmailToReview).delete_code = FILTER_SUB_BAD_WORDS
End Sub

Private Sub cmdClose_Click()
  If frmReview.gblnUpdate = True Then frmReview.UpdateStatus
  Unload frmAddWord
End Sub

Private Sub Form_Load()
  txtPhrase.Text = gstrString
 ' txtAdd.Text = CleanupString(gstrString)
End Sub

Private Sub txtPhrase_Change()
  'txtPhrase.Text = gstrString
  'txtAdd.Text = CleanupString(txtPhrase.Text)

End Sub
