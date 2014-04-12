VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Top             =   2100
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Read"
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1620
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type header
  size As String * 20
End Type

Private Type Person
  fname As String * 10
  lname As String * 10
End Type

Private nCount As Integer
Private Him As Person
Private Head As header

Private Sub Command1_Click()

  Head.size = (nCount + 1) * 20
  
  Open App.Path & "\data.txt" For Random As #1 Len = Len(Him)
    Put #1, 1, Head
  Close #1
    
  Him.fname = Text1.Text
  Him.lname = Text2.Text
  nCount = nCount + 1
  Open App.Path & "\data.txt" For Random As #1 Len = Len(Him)
    Put #1, nCount + 2, Him
  Close #1
  HScroll1.Max = nCount
    
End Sub

Private Sub Command2_Click()
  Text1.Text = ""
  Text2.Text = ""
End Sub

Private Sub Command3_Click()
  If HScroll1.Max < 1 Then Exit Sub
  
  Open App.Path & "\data.txt" For Random As #1 Len = Len(Him)
    Get #1, HScroll1.Value + 2, Him
  Close #1
  Text1.Text = Him.fname
  Text2.Text = Him.lname

End Sub

Private Sub Form_Load()
  
  Open App.Path & "\data.txt" For Append As #1
    If LOF(1) > 20 Then
      nCount = (LOF(1) - 20) \ Len(Him)
    End If
  Close #1
  MsgBox nCount & " records in file."
  HScroll1.Min = 1
  HScroll1.Max = nCount
  
End Sub
