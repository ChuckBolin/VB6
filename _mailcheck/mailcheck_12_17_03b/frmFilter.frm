VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Configuration"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Spam Word and Phrases List:"
      Height          =   2715
      Left            =   0
      TabIndex        =   1
      Top             =   3480
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
      Caption         =   "SPAM Filter:"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8535
      Begin VB.TextBox txtMaxMsgCons 
         Height          =   285
         Left            =   2460
         TabIndex        =   18
         Text            =   "6"
         Top             =   2280
         Width           =   555
      End
      Begin VB.TextBox txtMaxSubCons 
         Height          =   285
         Left            =   2460
         TabIndex        =   17
         Text            =   "6"
         Top             =   1200
         Width           =   555
      End
      Begin VB.CheckBox chkMsgCons 
         Caption         =   "Message Body Consonants:"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkMsgPhrases 
         Caption         =   "Message Body Phrases and Words"
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox chkSubCons 
         Caption         =   "Subject Consonants:"
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkSubPhrases 
         Caption         =   "Subject Phrases and Words"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Value           =   1  'Checked
         Width           =   2295
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
  
  If chkMaxSubLen.Value = vbChecked Then
    gblnMaxSubLen = True
    gintMaxSubLen = CInt(txtMaxSubLen.Text)
  Else
    gblnMaxSubLen = False
  End If
  
  If chkMinSubLen.Value = vbChecked Then
    gblnMinSubLen = True
    gintMinSubLen = CInt(txtMinSubLen.Text)
  Else
    gblnMinSubLen = False
  End If
  
  If chkSubPhrases.Value = vbChecked Then
    gblnSubPhrases = True
  Else
    gblnSubPhrases = False
  End If
  
  If chkSubCons.Value = vbChecked Then
    gblnSubConsonants = True
    gintMaxSubConsonants = CInt(txtMaxSubCons.Text)
  Else
    gblnSubConsonants = False
  End If
    
  If chkMsgPhrases.Value = vbChecked Then
    gblnMsgphrases = True
  Else
    gblnMsgphrases = False
  End If
    
  If chkMsgCons.Value = vbChecked Then
    gblnMsgConsonants = True
    gintMaxMsgConsonants = CInt(txtMaxMsgCons.Text)
  Else
    gblnMsgConsonants = False
  End If
  
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
  
  'load spam words and phrases list box
  For X = 1 To UBound(word)
    lstWords.AddItem word(X)
  Next X
  txtNum.Text = UBound(word)

  'configurs SPAM filter
  If gblnMaxSubLen = True Then
    chkMaxSubLen.Value = vbChecked
    txtMaxSubLen.Text = gintMaxSubLen
  Else
    chkMaxSubLen.Value = vbUnchecked
    txtMaxSubLen.Text = gintMaxSubLen
  End If
      
  If gblnMinSubLen = True Then
    chkMinSubLen.Value = vbChecked
    txtMinSubLen.Text = gintMinSubLen
  Else
    chkMinSubLen.Value = vbUnchecked
    txtMinSubLen.Text = gintMinSubLen
  End If
    
  If gblnSubPhrases = True Then
    chkSubPhrases.Value = vbChecked
  Else
    chkSubPhrases.Value = vbUnchecked
  End If
  
  If gblnSubConsonants = True Then
    chkSubCons.Value = vbChecked
    txtMaxSubCons.Text = gintMaxSubConsonants
  Else
    chkSubCons.Value = vbUnchecked
    txtMaxSubCons.Text = gintMaxSubConsonants
  End If
  
  If gblnMsgphrases = True Then
    chkMsgPhrases.Value = vbChecked
  Else
    chkMsgPhrases.Value = vbUnchecked
  End If
  
  If gblnMsgConsonants = True Then
    chkMsgCons.Value = vbChecked
    txtMaxMsgCons.Text = gintMaxMsgConsonants
  Else
    chkMsgCons.Value = vbChecked
    txtMaxMsgCons.Text = gintMaxMsgConsonants
  End If

End Sub

