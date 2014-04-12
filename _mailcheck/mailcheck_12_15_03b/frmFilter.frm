VERSION 5.00
Begin VB.Form frmFilter 
   Caption         =   "Filter Configuration"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Filter Spam Word and Phrases List:"
      Height          =   2715
      Left            =   60
      TabIndex        =   1
      Top             =   4320
      Width           =   8535
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
         Left            =   7260
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtAdd 
         Height          =   375
         Left            =   3780
         TabIndex        =   3
         Top             =   240
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject Filter"
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8295
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim x As Integer
  
  lstWords.AddItem LTrim(RTrim(txtAdd.Text))
  ReDim word(UBound(word) + 1)
  
  Open App.Path & "\spamwords.txt" For Output As #1
    For x = 1 To UBound(word)
      word(x) = lstWords.List(x - 1)
      Print #1, word(x)
    Next x
  Close #1
  
End Sub

Private Sub cmdDelete_Click()
  Dim x As Integer
  
  If lstWords.ListIndex < 0 Then Exit Sub
  lstWords.RemoveItem lstWords.ListIndex
  
  ReDim word(UBound(word) - 1)
  
  Open App.Path & "\spamwords.txt" For Output As #1
    For x = 1 To UBound(word)
      word(x) = lstWords.List(x - 1)
      Print #1, word(x)
    Next x
  Close #1
  
  
End Sub

Private Sub Form_Load()
  Dim x As Integer
  
  'load list box
  For x = 1 To UBound(word)
    lstWords.AddItem word(x)
  Next x
End Sub

