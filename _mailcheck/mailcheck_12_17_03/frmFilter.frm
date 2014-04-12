VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Configuration"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7380
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Spam Word and Phrases List:"
      Height          =   2715
      Left            =   60
      TabIndex        =   1
      Top             =   4320
      Width           =   8535
      Begin VB.TextBox txtNum 
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3780
         TabIndex        =   5
         Top             =   2040
         Width           =   1155
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   660
         Width           =   1155
      End
      Begin VB.TextBox txtAdd 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   660
         Width           =   3435
      End
      Begin VB.ListBox lstWords 
         Height          =   2205
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Total Phrases:"
         Height          =   255
         Left            =   3780
         TabIndex        =   7
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject Filter"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.CheckBox chkSpamPhrases 
         Caption         =   "Spam Phrases and Words"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkMaxSubLen 
         Caption         =   "Maximum Subject Length:"
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkMinSubLen 
         Caption         =   "Minimum Subject Length:"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtMaxSubLen 
         Height          =   285
         Left            =   2460
         TabIndex        =   10
         Text            =   "50"
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtMinSubLen 
         Height          =   285
         Left            =   2460
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim X As Integer
  
  lstWords.AddItem LTrim(RTrim(txtAdd.Text))
  ReDim word(UBound(word) + 1)
  
  Open App.Path & "\spamwords.txt" For Output As #1
    For X = 1 To UBound(word)
      word(X) = lstWords.List(X - 1)
      Print #1, word(X)
    Next X
  Close #1
  txtNum.Text = UBound(word)
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
  Dim X As Integer
  
  If lstWords.ListIndex < 0 Then Exit Sub
  lstWords.RemoveItem lstWords.ListIndex
  
  ReDim word(UBound(word) - 1)
  
  Open App.Path & "\spamwords.txt" For Output As #1
    For X = 1 To UBound(word)
      word(X) = lstWords.List(X - 1)
      Print #1, word(X)
    Next X
  Close #1
  txtNum.Text = UBound(word)
End Sub

Private Sub Form_Load()
  Dim X As Integer
  
  'load list box
  For X = 1 To UBound(word)
    lstWords.AddItem word(X)
  Next X
  txtNum.Text = UBound(word)
End Sub

